




use std::{mem::ManuallyDrop, ptr};

use windows::{Win32::{System::{Com::{self, IDispatch, DISPPARAMS, VARIANT}, Ole}, Globalization}, core::{InParam, PWSTR}};



#[allow(unused_labels)]

fn main() {
    unsafe{
        Com::CoInitializeEx(std::ptr::null(), Default::default()).unwrap();

        let dm = Dmsoft::new().unwrap();

        let s = dm.Ver().unwrap();

        println!("{}", s);
    }
    
    println!("#################################")
    
}


struct Dmsoft{
    obj: IDispatch
}

#[allow(non_snake_case)]
impl Dmsoft{
    unsafe fn new() -> windows::core::Result<Self>{
        let guid = Com::CLSIDFromProgID(windows::w!("dm.dmsoft"))?;
        let r = Com::CoCreateInstance::<'_, InParam<'_,_>, IDispatch>(&guid, InParam::null(), Com::CLSCTX_ALL)?;
        Ok(Self{
            obj: r
        })
    }

    unsafe fn Ver(&self) -> windows::core::Result<String>{
        let rgdispid = {
            static mut RGDISPID:i32 = -1;
            if RGDISPID == -1{
                let name = windows::w!("Ver");
                let name = PWSTR::from_raw(name.as_ptr() as *mut _);
                self.obj.GetIDsOfNames(ptr::null(), &name, 1, 1, &mut RGDISPID)?;
            };
            RGDISPID
        };

        let dispparams = DISPPARAMS{ rgvarg: ptr::null_mut(), rgdispidNamedArgs: ptr::null_mut(), cArgs: 0, cNamedArgs: 0 };
        let mut var = VARIANT::default();   
        self.obj.Invoke(rgdispid, ptr::null(), Globalization::LOCALE_ALL, Ole::DISPATCH_METHOD as u16, &dispparams, &mut var, ptr::null_mut(), ptr::null_mut())?;
        let var = ManuallyDrop::into_inner(var.Anonymous.Anonymous);
        let a = ManuallyDrop::into_inner(var.Anonymous.bstrVal);

        let s:String = a.try_into().unwrap();
        Ok(s)
        
    }

    // TODO: 其他函数映射
}


