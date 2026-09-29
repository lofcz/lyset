//! Late-bound IDispatch helpers that keep the full EXCEPINFO.
//!
//! Word, unlike PowerPoint, reports why an open failed through the automation
//! exception (`bstrDescription`, `scode`). That text is the most useful signal
//! the verifier produces, so every call keeps it instead of flattening it.

use serde::Serialize;
use windows::core::{GUID, PCWSTR};
use windows::Win32::Foundation::DISP_E_PARAMNOTFOUND;
use windows::Win32::System::Com::{
    DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, DISPPARAMS, EXCEPINFO,
    IDispatch,
};
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;
use windows::Win32::System::Variant::{VARIANT, VT_ERROR};

const LOCALE_USER_DEFAULT: u32 = 0x0400;
/// HRESULT returned by Invoke when the server filled EXCEPINFO.
const DISP_E_EXCEPTION: i32 = 0x8002_0009u32 as i32;

/// One failed automation call, with everything Word told us.
#[derive(Serialize, Clone, Debug, Default)]
pub struct ComFailure {
    /// HRESULT of the Invoke call itself, as `0x…`.
    pub hresult: String,
    /// `EXCEPINFO.scode`, as `0x…`, when the server raised an exception.
    pub scode: Option<String>,
    /// Word's error number (low word of a `0x800A….` scode), e.g. 5981.
    #[serde(rename = "wordError")]
    pub word_error: Option<u32>,
    pub source: Option<String>,
    pub description: String,
}

impl std::fmt::Display for ComFailure {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{} ({}", self.description, self.scode.as_deref().unwrap_or(&self.hresult))?;
        if let Some(n) = self.word_error {
            write!(f, ", Word error {n}")?;
        }
        write!(f, ")")
    }
}

impl ComFailure {
    pub fn from_error(err: &windows::core::Error) -> Self {
        Self {
            hresult: hex(err.code().0),
            description: err.message().trim().to_string(),
            ..Default::default()
        }
    }

    /// The server process died or disconnected (killed by the watchdog, crashed).
    pub fn is_disconnect(&self) -> bool {
        matches!(
            self.hresult.as_str(),
            "0x800706BA" | "0x800706BE" | "0x80010108" | "0x80010012" | "0x800706BF"
        )
    }
}

fn hex(value: i32) -> String {
    format!("0x{:08X}", value as u32)
}

pub type CallResult<T> = Result<T, ComFailure>;

pub fn dispid(disp: &IDispatch, name: &str) -> CallResult<i32> {
    unsafe {
        let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let ptr = PCWSTR(wide.as_ptr());
        let mut id = 0i32;
        disp.GetIDsOfNames(&GUID::zeroed(), &ptr, 1, LOCALE_USER_DEFAULT, &mut id)
            .map_err(|err| ComFailure::from_error(&err))?;
        Ok(id)
    }
}

fn invoke(disp: &IDispatch, name: &str, flags: DISPATCH_FLAGS, mut args: Vec<VARIANT>) -> CallResult<VARIANT> {
    let id = dispid(disp, name)?;
    unsafe {
        args.reverse();
        let mut named = DISPID_PROPERTYPUT;
        let is_put = flags == DISPATCH_PROPERTYPUT;
        let params = DISPPARAMS {
            rgvarg: if args.is_empty() { std::ptr::null_mut() } else { args.as_mut_ptr() },
            rgdispidNamedArgs: if is_put { &mut named } else { std::ptr::null_mut() },
            cArgs: args.len() as u32,
            cNamedArgs: if is_put { 1 } else { 0 },
        };
        let mut result = VARIANT::default();
        let mut excep = EXCEPINFO::default();
        let mut arg_err = 0u32;
        match disp.Invoke(
            id,
            &GUID::zeroed(),
            LOCALE_USER_DEFAULT,
            flags,
            &params,
            Some(&mut result),
            Some(&mut excep),
            Some(&mut arg_err),
        ) {
            Ok(()) => Ok(result),
            Err(err) => Err(failure_from(name, &err, &mut excep)),
        }
    }
}

unsafe fn failure_from(name: &str, err: &windows::core::Error, excep: &mut EXCEPINFO) -> ComFailure {
    let mut failure = ComFailure::from_error(err);
    if err.code().0 == DISP_E_EXCEPTION {
        if let Some(fill) = excep.pfnDeferredFillIn {
            let _ = fill(excep);
        }
        let scode = if excep.scode != 0 { excep.scode } else { excep.wCode as i32 };
        failure.scode = Some(hex(scode));
        if (scode as u32) & 0xFFFF_0000 == 0x800A_0000 {
            failure.word_error = Some((scode as u32) & 0xFFFF);
        }
        let source = excep.bstrSource.to_string();
        if !source.is_empty() {
            failure.source = Some(source);
        }
        let description = excep.bstrDescription.to_string();
        if !description.trim().is_empty() {
            failure.description = description.trim().to_string();
        }
    }
    if failure.description.is_empty() {
        failure.description = format!("{name} failed");
    }
    failure
}

pub fn prop_get(disp: &IDispatch, name: &str) -> CallResult<VARIANT> {
    invoke(disp, name, DISPATCH_PROPERTYGET, Vec::new())
}

pub fn prop_put(disp: &IDispatch, name: &str, value: VARIANT) -> CallResult<()> {
    invoke(disp, name, DISPATCH_PROPERTYPUT, vec![value])?;
    Ok(())
}

pub fn method(disp: &IDispatch, name: &str, args: Vec<VARIANT>) -> CallResult<VARIANT> {
    invoke(disp, name, DISPATCH_METHOD, args)
}

pub fn as_dispatch(value: &VARIANT) -> CallResult<IDispatch> {
    IDispatch::try_from(value).map_err(|err| ComFailure::from_error(&err))
}

pub fn as_i32(value: &VARIANT) -> Option<i32> {
    i32::try_from(value).ok()
}

pub fn as_string(value: &VARIANT) -> Option<String> {
    windows::core::BSTR::try_from(value).ok().map(|b| b.to_string())
}

pub fn get_i32(disp: &IDispatch, path: &[&str]) -> Option<i32> {
    let (last, parents) = path.split_last()?;
    let mut current = disp.clone();
    for name in parents {
        current = as_dispatch(&prop_get(&current, name).ok()?).ok()?;
    }
    as_i32(&prop_get(&current, last).ok()?)
}

pub fn vt_bstr(text: &str) -> VARIANT {
    VARIANT::from(text)
}

pub fn vt_i4(value: i32) -> VARIANT {
    VARIANT::from(value)
}

pub fn vt_bool(value: bool) -> VARIANT {
    VARIANT::from(value)
}

/// An omitted optional positional argument (`VT_ERROR` / `DISP_E_PARAMNOTFOUND`).
pub fn vt_missing() -> VARIANT {
    let mut v = VARIANT::default();
    unsafe {
        let inner = &mut *v.Anonymous.Anonymous;
        inner.vt = VT_ERROR;
        inner.Anonymous.scode = DISP_E_PARAMNOTFOUND.0;
    }
    v
}
