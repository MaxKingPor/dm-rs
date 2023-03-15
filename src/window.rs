use std::mem::ManuallyDrop;

use windows::Win32::System::Com::VARIANT;

use crate::{Dmsoft, Result};

#[allow(non_snake_case)]
impl Dmsoft {
    /// 解除绑定窗口,并释放系统资源.一般在OnScriptExit调用
    /// # The function prototype
    /// ```C++
    /// long dmsoft::UnBindWindow()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.UnBindWindow().unwrap();
    /// ```
    pub unsafe fn UnBindWindow(&self) -> Result<i32> {
        static NAME: &str = "UnBindWindow";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }
    
    /// 把窗口坐标转换为屏幕坐标
    /// # The function prototype
    /// ```C++
    /// long dmsoft::ClientToScreen(long x,long y,const TCHAR * color,double sim)
    /// ```
    /// # Args
    /// * `hwnd:i32`: 指定的窗口句柄
    /// * `x:&mut i32`: 窗口X坐标
    /// * `y:&mut i32`: 窗口Y坐标
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let hwnd = 0;
    /// let dm = Dmsoft::new();
    /// let status = dm.ClientToScreen(hwnd,0,0) .unwrap();
    /// ```
    pub unsafe fn ClientToScreen(&self, hwnd: i32, x: &mut i32, y: &mut i32) -> Result<i32> {
        static NAME: &str = "ClientToScreen";
        let mut px = VARIANT::default();
        let mut py = VARIANT::default();

        let mut args = [
            Dmsoft::pvarVal(&mut py),
            Dmsoft::pvarVal(&mut px),
            Dmsoft::longVar(hwnd),
        ];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        *x = px.Anonymous.Anonymous.Anonymous.lVal;
        *y = py.Anonymous.Anonymous.Anonymous.lVal;
        Ok(result.Anonymous.lVal)
    }

    /// 把屏幕坐标转换为窗口坐标
    /// # The function prototype
    /// ```C++
    /// long dmsoft::ScreenToClient(long x,long y,const TCHAR * color,double sim)
    /// ```
    /// # Args
    /// * `hwnd:i32`: 指定的窗口句柄
    /// * `x:&mut i32`: 窗口X坐标
    /// * `y:&mut i32`: 窗口Y坐标
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let hwnd = 0;
    /// let dm = Dmsoft::new();
    /// let status = dm.ScreenToClient(hwnd,0,0) .unwrap();
    /// ```
    pub unsafe fn ScreenToClient(&self, hwnd: i32, x: &mut i32, y: &mut i32) -> Result<i32> {
        static NAME: &str = "ScreenToClient";
        let mut px = VARIANT::default();
        let mut py = VARIANT::default();

        let mut args = [
            Dmsoft::pvarVal(&mut py),
            Dmsoft::pvarVal(&mut px),
            Dmsoft::longVar(hwnd),
        ];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        *x = px.Anonymous.Anonymous.Anonymous.lVal;
        *y = py.Anonymous.Anonymous.Anonymous.lVal;
        Ok(result.Anonymous.lVal)
    }
}
