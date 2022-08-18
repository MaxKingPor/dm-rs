use std::{mem::ManuallyDrop, ptr, collections::HashMap, sync::Mutex};

use windows::{Win32::{System::{Com::{self, IDispatch, DISPPARAMS, VARIANT, VARIANT_0, VARIANT_0_0}, Ole::{self, VariantInit}}, Foundation::BSTR}, core::{InParam, PWSTR, HSTRING}};

// 在windows-rs 中并未搜索到此参数 使用本地定义 来源:
// https://docs.microsoft.com/en-us/windows/win32/intl/locale-user-default
const LOCALE_USER_DEFAULT:u32 = 0x0400;
const EMPTY_ARGS: [VARIANT;0] = [];
pub struct Dmsoft{
    obj: IDispatch,
    catch: Mutex<HashMap<&'static str, i32>>
}


#[derive(Debug)]
pub enum Error {
    WinError(windows::core::Error),
    IdError
}

type Result<T> = core::result::Result<T, Error>;
/// 大漠插件绑定
#[allow(non_snake_case)]
impl Dmsoft{

    pub unsafe fn new() -> windows::core::Result<Self>{
        let guid = Com::CLSIDFromProgID(windows::w!("dm.dmsoft"))?;
        let r = Com::CoCreateInstance::<'_, InParam<'_,_>, IDispatch>(&guid, InParam::null(), Com::CLSCTX_ALL)?;
        Ok(Self{
            obj: r,
            catch: Mutex::new(HashMap::new())
        })
    }

    #[allow(const_item_mutation)]
    pub unsafe fn Ver(&self) -> Result<String>{
        const NAME:&'static str = "Ver";
        let result = self.Invoke(NAME, &mut EMPTY_ARGS)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())

    }

    pub unsafe fn SetPath(&self, path:&str) -> Result<i32>{
        const NAME:&'static str = "SetPath";
        
        let mut args = [Dmsoft::bstrVal(path)];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let a = result.Anonymous.lVal;
        Ok(a)
        
    }

    pub unsafe fn Ocr(&self, x1:i32,y1:i32,x2:i32,y2:i32,color:&str, sim:f64) -> Result<String>{
        const NAME:&'static str = "Ocr";

        let mut args = [
            Dmsoft::doubleVar(sim),
            Dmsoft::bstrVal(color),
            Dmsoft::longVar(y2),
            Dmsoft::longVar(x2),
            Dmsoft::longVar(y1),
            Dmsoft::longVar(x1)
        ];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    // long dmsoft::FindStr(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,long * x,long * y)
    pub unsafe fn FindStr(&self, x1:i32, y1:i32, x2:i32,y2:i32, str:&str, color:&str, sim:f64, x:*mut i32, y:*mut i32)-> Result<i32> {
        const NAME:&'static str = "FindStr";
        let mut px = VARIANT::default();
        let mut py = VARIANT::default();
        let mut args = [
            Dmsoft::pvarVal(&mut py),
            Dmsoft::pvarVal(&mut px),
            Dmsoft::doubleVar(sim),
            Dmsoft::bstrVal(color),
            Dmsoft::bstrVal(str),
            Dmsoft::longVar(y2),
            Dmsoft::longVar(x2),
            Dmsoft::longVar(y1),
            Dmsoft::longVar(x1)
        ];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        if !x.is_null() {
            *x = px.Anonymous.Anonymous.Anonymous.lVal;
        }
        if !y.is_null() {
            *y = py.Anonymous.Anonymous.Anonymous.lVal;
        }
        
        Ok(result.Anonymous.lVal)
    
    }

    // long dmsoft::GetResultCount(const TCHAR * str)
    pub unsafe fn GetResultCount(&self, str: &str) -> Result<i32>{
        const NAME:&'static str = "GetResultCount";
        let mut args = [Dmsoft::bstrVal(str)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let a = result.Anonymous.lVal;
        Ok(a)
    }
    // TODO: 其他函数映射
}


/// 辅助函数
#[allow(non_snake_case)]
impl Dmsoft{
    pub unsafe fn Invoke(&self, name:&'static str, args: &mut [VARIANT]) -> Result<VARIANT>{
        let mut map = self.catch.lock().unwrap();
        let rgdispid = *map.entry(name).or_insert_with_key(|key|{
            let mut id = -1;
            let name = HSTRING::from(*key);
            let func_name = PWSTR::from_raw(name.as_ptr() as *mut _);
            // 在调试时解决 expect
            self.obj.GetIDsOfNames(ptr::null(), &func_name, 1, LOCALE_USER_DEFAULT, &mut id).expect("调用 GetIDsOfNames 获取ID 异常: ");
            id
        });
        drop(map);
        if rgdispid == -1 {
            return Err(Error::IdError);
        };

        let mut result = VARIANT::default();   
        let dispparams = DISPPARAMS{ rgvarg: args.as_mut_ptr(), rgdispidNamedArgs: ptr::null_mut(), cArgs: args.len() as u32, cNamedArgs: 0 };
        if let Err(e) = self.obj.Invoke(rgdispid, ptr::null(), LOCALE_USER_DEFAULT, Ole::DISPATCH_METHOD as u16, &dispparams, &mut result, ptr::null_mut(), ptr::null_mut()) {
            return Err(Error::WinError(e));
        };
        Ok(result)
    }

    pub unsafe fn bstrVal(var:&str) -> VARIANT{
        let s = BSTR::from_raw(HSTRING::from(var).as_ptr());
        let mut arg = VARIANT_0_0::default();
        arg.vt = Ole::VT_BSTR.0 as u16;
        arg.Anonymous.bstrVal = ManuallyDrop::new(s);
        VARIANT{ Anonymous: VARIANT_0{Anonymous:ManuallyDrop::new(arg)} }
    }


    pub unsafe fn longVar(var:i32) -> VARIANT{
        let mut arg = VARIANT_0_0::default();
        arg.vt = Ole::VT_I4.0 as u16;
        arg.Anonymous.lVal = var;
        VARIANT{ Anonymous: VARIANT_0{Anonymous:ManuallyDrop::new(arg)} }
    }
    pub unsafe fn pvarVal(var:*mut VARIANT) -> VARIANT{
        let mut arg = VARIANT_0_0::default();
        arg.vt = (Ole::VT_BYREF.0| Ole::VT_VARIANT.0) as u16;
        arg.Anonymous.pvarVal = var;
        VARIANT{ Anonymous: VARIANT_0{Anonymous:ManuallyDrop::new(arg)} }
    }
    pub unsafe fn doubleVar(var:f64) -> VARIANT{
        let mut arg = VARIANT_0_0::default();
        arg.vt = Ole::VT_R8.0 as u16;
        arg.Anonymous.dblVal = var;
        VARIANT{ Anonymous: VARIANT_0{Anonymous:ManuallyDrop::new(arg)} }
    }


}



