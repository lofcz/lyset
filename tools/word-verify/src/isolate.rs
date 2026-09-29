//! Run Word on a private desktop so its windows and modal dialogs never reach
//! the interactive session, and so UI Automation can find them by process.

use windows::core::{w, PCWSTR};
use windows::Win32::Foundation::{CloseHandle, LocalFree, HANDLE, HLOCAL};
use windows::Win32::Security::Authorization::{ConvertStringSecurityDescriptorToSecurityDescriptorW, SDDL_REVISION_1};
use windows::Win32::Security::{PSECURITY_DESCRIPTOR, SECURITY_ATTRIBUTES};
use windows::Win32::System::JobObjects::{
    AssignProcessToJobObject, CreateJobObjectW, JobObjectExtendedLimitInformation, SetInformationJobObject,
    JOBOBJECT_EXTENDED_LIMIT_INFORMATION, JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE,
};
use windows::Win32::System::StationsAndDesktops::{
    CloseDesktop, CreateDesktopW, OpenDesktopW, SetThreadDesktop, DESKTOP_CONTROL_FLAGS, HDESK,
};
use windows::Win32::System::Threading::{
    OpenProcess, TerminateProcess, PROCESS_QUERY_LIMITED_INFORMATION, PROCESS_SET_QUOTA, PROCESS_TERMINATE,
};

const DESKTOP_NAME: PCWSTR = w!("lyset-word-verify-pv");
/// Protected View parses in a low-integrity / AppContainer sandbox process.
/// A desktop created with the default descriptor carries a medium mandatory
/// label, so that process cannot attach and every Protected View open fails
/// with "Word experienced an error trying to open the file". Grant the
/// sandbox principals access and label the desktop low.
const DESKTOP_SDDL: PCWSTR = w!("D:(A;;GA;;;SY)(A;;GA;;;BA)(A;;GA;;;AU)(A;;GA;;;AC)(A;;GA;;;S-1-15-2-2)S:(ML;;NW;;;LW)");
const DESKTOP_ACCESS: u32 = 0x01FF;

pub struct Isolation {
    desktop: HDESK,
    job: HANDLE,
    adopted: Option<HANDLE>,
    pub pid: u32,
}

impl Isolation {
    pub fn new() -> windows::core::Result<Self> {
        unsafe {
            let mut descriptor = PSECURITY_DESCRIPTOR::default();
            ConvertStringSecurityDescriptorToSecurityDescriptorW(DESKTOP_SDDL, SDDL_REVISION_1, &mut descriptor, None)?;
            let attributes = SECURITY_ATTRIBUTES {
                nLength: std::mem::size_of::<SECURITY_ATTRIBUTES>() as u32,
                lpSecurityDescriptor: descriptor.0,
                bInheritHandle: false.into(),
            };
            let created = CreateDesktopW(DESKTOP_NAME, None, None, DESKTOP_CONTROL_FLAGS(0), DESKTOP_ACCESS, Some(&attributes));
            let _ = LocalFree(Some(HLOCAL(descriptor.0)));
            let desktop = match created {
                Ok(desktop) => desktop,
                Err(_) => OpenDesktopW(DESKTOP_NAME, DESKTOP_CONTROL_FLAGS(0), false, DESKTOP_ACCESS)?,
            };
            let job = CreateJobObjectW(None, PCWSTR::null())?;
            let mut limit = JOBOBJECT_EXTENDED_LIMIT_INFORMATION::default();
            limit.BasicLimitInformation.LimitFlags = JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;
            SetInformationJobObject(
                job,
                JobObjectExtendedLimitInformation,
                &limit as *const _ as *const _,
                std::mem::size_of::<JOBOBJECT_EXTENDED_LIMIT_INFORMATION>() as u32,
            )?;
            Ok(Self { desktop, job, adopted: None, pid: 0 })
        }
    }

    /// Must run before CoInitialize on the calling thread: STA init creates a
    /// hidden window, after which SetThreadDesktop fails with ERROR_BUSY.
    pub fn bind_current_thread() -> windows::core::Result<()> {
        unsafe {
            let desk = OpenDesktopW(DESKTOP_NAME, DESKTOP_CONTROL_FLAGS(0), false, DESKTOP_ACCESS)?;
            SetThreadDesktop(desk)
        }
    }
}

impl Isolation {
    /// Track a Word process by pid (Word exposes no application HWND before a
    /// document window exists).
    pub fn adopt_pid(&mut self, pid: u32) -> Result<(), String> {
        self.release();
        unsafe {
            let process = OpenProcess(PROCESS_SET_QUOTA | PROCESS_TERMINATE | PROCESS_QUERY_LIMITED_INFORMATION, false, pid)
                .map_err(|err| err.to_string())?;
            let _ = AssignProcessToJobObject(self.job, process);
            self.adopted = Some(process);
        }
        self.pid = pid;
        Ok(())
    }

    /// Terminate and forget the adopted Word process, keeping the desktop.
    pub fn release(&mut self) {
        unsafe {
            if let Some(process) = self.adopted.take() {
                let _ = TerminateProcess(process, 0);
                let _ = CloseHandle(process);
            }
        }
        self.pid = 0;
    }
}

impl Drop for Isolation {
    fn drop(&mut self) {
        unsafe {
            if let Some(process) = self.adopted.take() {
                let _ = TerminateProcess(process, 0);
                let _ = CloseHandle(process);
            }
            if !self.job.is_invalid() {
                let _ = CloseHandle(self.job);
            }
            if !self.desktop.is_invalid() {
                let _ = CloseDesktop(self.desktop);
            }
        }
    }
}
