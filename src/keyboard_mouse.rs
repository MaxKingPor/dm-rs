use std::mem::ManuallyDrop;

use crate::{Dmsoft, KeyMap, Result};
#[allow(non_snake_case)]
impl Dmsoft {
    /// 按住指定的虚拟键码
    /// # The function prototype
    /// ```C++
    /// long dmsoft::KeyDown(long vk)
    /// ```
    /// # Args
    /// * `vk:KeyMap<'a>`:  虚拟按键码
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.KeyDown(keymap::KEY_A).unwrap();
    /// ```
    pub unsafe fn KeyDown(&self, vk: KeyMap) -> Result<i32> {
        static NAME: &str = "KeyDown";
        let mut args = [Dmsoft::longVar(vk.get_id())];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 弹起来虚拟键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::KeyUp(long vk)
    /// ```
    /// # Args
    /// * `vk:KeyMap<'a>`:  虚拟按键码
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.KeyUp(keymap::KEY_A).unwrap();
    /// ```
    pub unsafe fn KeyUp(&self, vk: KeyMap) -> Result<i32> {
        static NAME: &str = "KeyUp";
        let mut args = [Dmsoft::longVar(vk.get_id())];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 按下鼠标左键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::LeftClick()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.LeftClick().unwrap();
    /// ```
    pub unsafe fn LeftClick(&self) -> Result<i32> {
        static NAME: &str = "LeftClick";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 按下鼠标右键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::RightClick()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.RightClick().unwrap();
    /// ```
    pub unsafe fn RightClick(&self) -> Result<i32> {
        static NAME: &str = "RightClick";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 按下鼠标中键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::MiddleClick()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.MiddleClick().unwrap();
    /// ```
    pub unsafe fn MiddleClick(&self) -> Result<i32> {
        static NAME: &str = "MiddleClick";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 双击鼠标左键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::LeftDoubleClick()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.LeftDoubleClick().unwrap();
    /// ```
    pub unsafe fn LeftDoubleClick(&self) -> Result<i32> {
        static NAME: &str = "LeftDoubleClick";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 按住鼠标左键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::LeftDown()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.LeftDown().unwrap();
    /// ```
    pub unsafe fn LeftDown(&self) -> Result<i32> {
        static NAME: &str = "LeftDown";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 弹起鼠标左键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::LeftUp()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.LeftUp().unwrap();
    /// ```
    pub unsafe fn LeftUp(&self) -> Result<i32> {
        static NAME: &str = "LeftUp";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 按住鼠标右键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::RightDown()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.RightDown().unwrap();
    /// ```
    pub unsafe fn RightDown(&self) -> Result<i32> {
        static NAME: &str = "RightDown";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 弹起鼠标右键
    /// # The function prototype
    /// ```C++
    /// long dmsoft::RightUp()
    /// ```
    /// # Args
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.RightUp().unwrap();
    /// ```
    pub unsafe fn RightUp(&self) -> Result<i32> {
        static NAME: &str = "RightUp";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 把鼠标移动到目的点(x,y)
    /// # The function prototype
    /// ```C++
    /// long dmsoft::MoveTo(long x,long y)
    /// ```
    /// # Args
    /// * `x:i32`: X坐标
    /// * `y:i32`: Y坐标
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.MoveTo(0,0).unwrap();
    /// ```
    pub unsafe fn MoveTo(&self, x: i32, y: i32) -> Result<i32> {
        static NAME: &str = "MoveTo";
        let mut args = [Dmsoft::longVar(y), Dmsoft::longVar(x)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 鼠标相对于上次的位置移动rx,ry.   
    ///
    /// 如果您要使前台鼠标移动的距离和指定的rx,ry一致,最好配合EnableMouseAccuracy函数来使用.
    ///
    /// # The function prototype
    /// ```C++
    /// long dmsoft::MoveR(long rx,long ry)
    /// ```
    /// # Args
    /// * `rx:i32`: 相对于上次的X偏移
    /// * `ry:i32`: 相对于上次的Y偏移
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.MoveR(0,0).unwrap();
    /// ```
    pub unsafe fn MoveR(&self, rx: i32, ry: i32) -> Result<i32> {
        static NAME: &str = "MoveR";
        let mut args = [Dmsoft::longVar(ry), Dmsoft::longVar(rx)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 按下指定的虚拟键码
    /// # The function prototype
    /// ```C++
    /// long dmsoft::KeyPress(long vk)
    /// ```
    /// # Args
    /// * `vk:KeyMap<'a>`:  虚拟按键码
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.KeyPress(keymap::KEY_A).unwrap();
    /// ```
    pub unsafe fn KeyPress(&self, vk: &KeyMap) -> Result<i32> {
        static NAME: &str = "KeyPress";
        let mut args = [Dmsoft::longVar(vk.get_id())];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }
}
