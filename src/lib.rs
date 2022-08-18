use std::{mem::ManuallyDrop, ptr};

use windows::{Win32::{System::{Com::{self, IDispatch, DISPPARAMS, VARIANT, VARIANT_0, VARIANT_0_0}, Ole}, Foundation::BSTR}, core::{InParam, PWSTR, HSTRING}};

// 在windows-rs 中并未搜索到此参数 使用本地定义 来源:
// https://docs.microsoft.com/en-us/windows/win32/intl/locale-user-default
const LOCALE_USER_DEFAULT:u32 = 0x0400;

pub struct Dmsoft{
    obj: IDispatch
}

#[allow(non_snake_case)]
impl Dmsoft{
    pub unsafe fn new() -> windows::core::Result<Self>{
        let guid = Com::CLSIDFromProgID(windows::w!("dm.dmsoft"))?;
        let r = Com::CoCreateInstance::<'_, InParam<'_,_>, IDispatch>(&guid, InParam::null(), Com::CLSCTX_ALL)?;
        Ok(Self{
            obj: r
        })
    }

    pub unsafe fn Ver(&self) -> windows::core::Result<String>{
        let rgdispid = {
            static mut RGDISPID:i32 = -1;
            if RGDISPID == -1{
                let name = windows::w!("Ver");
                let name = PWSTR::from_raw(name.as_ptr() as *mut _);
                self.obj.GetIDsOfNames(ptr::null(), &name, 1, LOCALE_USER_DEFAULT, &mut RGDISPID)?;
            };
            RGDISPID
        };

        let dispparams = DISPPARAMS{ rgvarg: ptr::null_mut(), rgdispidNamedArgs: ptr::null_mut(), cArgs: 0, cNamedArgs: 0 };
        let mut result = VARIANT::default();   
        self.obj.Invoke(rgdispid, ptr::null(), LOCALE_USER_DEFAULT, Ole::DISPATCH_METHOD as u16, &dispparams, &mut result, ptr::null_mut(), ptr::null_mut())?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let a = ManuallyDrop::into_inner(result.Anonymous.bstrVal);

        let s:String = a.try_into().unwrap();
        Ok(s)
        
    }

    pub unsafe fn SetPath(&self, path:&str) -> windows::core::Result<i32>{
        let rgdispid = {
            static mut RGDISPID:i32 = -1;
            if RGDISPID == -1{
                let name = windows::w!("SetPath");
                let name = PWSTR::from_raw(name.as_ptr() as *mut _);
                self.obj.GetIDsOfNames(ptr::null(), &name, 1, LOCALE_USER_DEFAULT, &mut RGDISPID)?;
            };
            RGDISPID
        };
       
        let s = HSTRING::from(path);
        let s = BSTR::from_raw(s.as_ptr());
        let mut args = VARIANT_0_0::default();
        args.vt = Ole::VT_BSTR.0 as u16;
        args.Anonymous.bstrVal = ManuallyDrop::new(s);
        let mut args = [VARIANT{ Anonymous: VARIANT_0{Anonymous:ManuallyDrop::new(args)} };1];
        
        let dispparams = DISPPARAMS{ rgvarg: args.as_mut_ptr(), rgdispidNamedArgs: ptr::null_mut(), cArgs: 1, cNamedArgs: 0 };
        let mut result = VARIANT::default();   
        self.obj.Invoke(rgdispid, ptr::null(), LOCALE_USER_DEFAULT, Ole::DISPATCH_METHOD as u16, &dispparams, &mut result, ptr::null_mut(), ptr::null_mut())?;
        
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let a = result.Anonymous.intVal;
        Ok(a)
        
    }

    // TODO: 其他函数映射
}


