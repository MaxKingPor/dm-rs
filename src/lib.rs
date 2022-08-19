#![allow(dead_code)]
#![warn(missing_docs)]
#![doc = include_str!("../README.md")]

use std::{collections::HashMap, mem::ManuallyDrop, ptr, sync::Mutex};

use windows::{
    core::{InParam, HSTRING, PWSTR},
    Win32::{
        Foundation::BSTR,
        System::{
            Com::{self, IDispatch, DISPPARAMS, VARIANT, VARIANT_0, VARIANT_0_0},
            Ole,
        },
    },
};

#[cfg(feature = "keymap")]
pub mod keymap;

/// 在windows-rs 中并未搜索到此参数 使用本地定义 来源:
///
/// https://docs.microsoft.com/en-us/windows/win32/intl/locale-user-default
pub const LOCALE_USER_DEFAULT: u32 = 0x0400;

/// dm.dmsoft API 绑定
pub struct Dmsoft {
    /// dm.dmsoft 链接实例
    obj: IDispatch,
    /// Invoke ID 缓存
    catch: Mutex<HashMap<&'static str, i32>>,
}

/// 异常枚举
#[derive(Debug)]
pub enum Error {
    /// 调用Windows API 时产生的Error
    WinError(windows::core::Error),
    /// 从缓存区获取的 Invoke ID 为 `-1`
    IdError,
}

/// API Result
type Result<T> = core::result::Result<T, Error>;

/// 大漠插件绑定
#[allow(non_snake_case)]
impl Dmsoft {
    /// 新建一个 dm.dmsoft API 绑定实例
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// ```
    pub unsafe fn new() -> windows::core::Result<Self> {
        let guid = Com::CLSIDFromProgID(windows::w!("dm.dmsoft"))?;
        let r = Com::CoCreateInstance::<'_, InParam<'_, _>, IDispatch>(
            &guid,
            InParam::null(),
            Com::CLSCTX_ALL,
        )?;
        Ok(Self {
            obj: r,
            catch: Mutex::new(HashMap::new()),
        })
    }

    /// 返回当前插件版本号
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::Ver()
    /// ```
    /// # Args
    ///
    /// # Return
    /// `String` 当前插件的版本描述字符串
    ///
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// // 返回版本号
    /// let ver = dm.Ver().unwrap();
    /// println!("Ver: {}", ver);
    /// ```
    #[allow(const_item_mutation)]
    pub unsafe fn Ver(&self) -> Result<String> {
        const NAME: &'static str = "Ver";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 设置全局路径,设置了此路径后,所有接口调用中,相关的文件都相对于此路径. 比如图片,字库等.
    /// # The function prototype
    /// ```C++
    /// long dmsoft::SetPath(const TCHAR * path)
    /// ```
    ///
    /// # Args
    /// * `path` 字符串: 路径,可以是相对路径,也可以是绝对路径
    ///
    /// ## Return
    /// `i32`: 0: 失败 1: 成功
    ///
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// // 以下代码把全局路径设置到了c盘根目录
    /// let result = dm.SetPath(r"c:\").unwrap();
    /// // 如下是把全局路径设置到了相对于当前exe所在的路径
    /// let result = dm.SetPath(r".\MyData").unwrap();
    /// // 以上，如果exe在c:\test\a.exe 那么，就相当于把路径设置到了c:\test\MyData
    /// ```
    pub unsafe fn SetPath(&self, path: &str) -> Result<i32> {
        const NAME: &'static str = "SetPath";

        let mut args = [Dmsoft::bstrVal(path)];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let a = result.Anonymous.lVal;
        Ok(a)
    }

    /// 识别屏幕范围(x1,y1,x2,y2)内符合color_format的字符串,并且相似度为sim,sim取值范围(0.1-1.0),
    ///
    /// 这个值越大越精确,越大速度越快,越小速度越慢,请斟酌使用!
    ///
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::Ocr(long x1,long y1,long x2,long y2,const TCHAR * color,double sim)
    /// ```
    /// # Args
    /// * `x1:i32`: 区域的左上X坐标
    /// * `y1:i32`: 区域的左上Y坐标
    /// * `x2:i32`: 区域的右下X坐标
    /// * `y2:i32`: 区域的右下Y坐标
    /// * `color_format:&str`: 颜色格式串. 可以包含换行分隔符,语法是","后加分割字符串. 具体可以查看下面的示例.注意，RGB和HSV,以及灰度格式都支持.
    /// * `sim:f64`:相似度,取值范围0.1-1.0
    ///
    /// ## Return
    /// * `String` 返回识别到的字符串
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// // RGB单色识别
    /// let s = dm.Ocr(0,0,2000,2000,"9f2e3f-000000",1.0).unwrap();
    /// // RGB单色差色识别
    /// let s = dm.Ocr(0,0,2000,2000,"9f2e3f-030303",1.0).unwrap();
    /// // RGB多色识别(最多支持10种,每种颜色用"|"分割)
    /// let s = dm.Ocr(0,0,2000,2000,"9f2e3f-030303|2d3f2f-000000|3f9e4d-100000",1.0).unwrap();
    /// //HSV多色识别(最多支持10种,每种颜色用"|"分割)
    /// let s = dm.Ocr(0,0,2000,2000,"20.30.40-0.0.0|30.40.50-0.0.0",1.0).unwrap();
    /// //灰度多色识别(最多支持10种,每种颜色用"|"分割)
    /// let s = dm.Ocr(0,0,2000,2000,"#40-0|#70-10",1.0).unwrap();
    /// //识别后,每行字符串用指定字符分割 比如用"|"字符分割
    /// let s = dm.Ocr(0,0,2000,2000,"9f2e3f-000000,|",1.0).unwrap();
    /// ```
    pub unsafe fn Ocr(
        &self,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
        color: &str,
        sim: f64,
    ) -> Result<String> {
        const NAME: &'static str = "Ocr";

        let mut args = [
            Dmsoft::doubleVar(sim),
            Dmsoft::bstrVal(color),
            Dmsoft::longVar(y2),
            Dmsoft::longVar(x2),
            Dmsoft::longVar(y1),
            Dmsoft::longVar(x1),
        ];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 在屏幕范围(x1,y1,x2,y2)内,查找string(可以是任意个字符串的组合),并返回符合color_format的坐标位置,相似度sim同Ocr接口描述.
    ///
    /// (多色,差色查找类似于Ocr接口,不再重述)
    /// # The function prototype
    /// ```C++
    /// long dmsoft::FindStr(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,long * x,long * y)
    /// ```
    /// # Args
    /// * `x1:i32`: 区域的左上X坐标
    /// * `y1:i32`: 区域的左上Y坐标
    /// * `x2:i32`: 区域的右下X坐标
    /// * `y2:i32`: 区域的右下Y坐标
    /// * `string:&str`: 待查找的字符串,可以是字符串组合，比如"长安|洛阳|大雁塔",中间用"|"来分割字符串
    /// * `color_format:&str`: 颜色格式串, 可以包含换行分隔符,语法是","后加分割字符串. 具体可以查看下面的示例 .注意，RGB和HSV,以及灰度格式都支持.
    /// * `sim:f64`: 相似度,取值范围0.1-1.0
    /// * `x:*mut i32`: 返回X坐标没找到返回-1
    /// * `y:*mut i32`: 返回Y坐标没找到返回-1
    /// # Return
    /// * `i32`: 返回字符串的索引 没找到返回-1, 比如"长安|洛阳",若找到长安，则返回0
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let (mut x,mut y) = (0,0);
    ///
    /// let dm_ret = dm.FindStr(0,0,2000,2000,"长安","9f2e3f-000000",1.0,&mut x,&mut y).unwrap();
    /// if x >= 0 and y >= 0{
    ///     dm.MoveTo(x, y);
    /// };
    ///
    /// let dm_ret = dm.FindStr(0,0,2000,2000,"长安|洛阳","9f2e3f-000000",1.0,&mut x,&mut y).unwrap();
    /// if x >= 0 and y >= 0{
    ///     dm.MoveTo(x, y);
    /// };
    ///
    /// // 查找时,对多行文本进行换行,换行分隔符是"|". 语法是在","后增加换行字符串.任意字符串都可以.
    /// let dm_ret = dm.FindStr(0,0,2000,2000,"长安|洛阳","9f2e3f-000000,|",1.0,&mut x,&mut y).unwrap();
    /// if x >= 0 and y >= 0{
    ///     dm.MoveTo(x, y);
    /// };
    /// ```
    /// # Note:
    /// * 此函数的原理是先Ocr识别，然后再查找。所以速度比FindStrFast要慢，尤其是在字库 很大，或者模糊度不为1.0时。\
    /// * 一般字库字符数量小于100左右，模糊度为1.0时，用FindStr要快一些,否则用FindStrFast.
    pub unsafe fn FindStr(
        &self,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
        str: &str,
        color: &str,
        sim: f64,
        x: *mut i32,
        y: *mut i32,
    ) -> Result<i32> {
        const NAME: &'static str = "FindStr";
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
            Dmsoft::longVar(x1),
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

    /// 对插件部分接口的返回值进行解析,并返回ret中的坐标个数
    /// # The function prototype
    /// ```C++
    /// long dmsoft::GetResultCount(const TCHAR * str)
    /// ```
    ///
    /// # Args
    /// * `str:&str` 部分接口的返回串
    /// # Return
    /// `i32` 返回ret中的坐标个数
    ///
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    ///
    /// let s = dm.FindColorEx(0,0,2000,2000,"123456-000000|abcdef-202020",1.0,0).unwrap();
    /// let count = dm.GetResultCount(s).unwrap();
    /// ```
    pub unsafe fn GetResultCount(&self, str: &str) -> Result<i32> {
        const NAME: &'static str = "GetResultCount";
        let mut args = [Dmsoft::bstrVal(str)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let a = result.Anonymous.lVal;
        Ok(a)
    }

    /// 对插件部分接口的返回值进行解析,并根据指定的第index个坐标,返回具体的值
    /// # The function prototype
    /// ```C++
    /// long dmsoft::GetResultPos(const TCHAR * str,long index,long * x,long * y)
    /// ```
    /// # Args
    /// `ret:&str`: 部分接口的返回串
    /// `index:i32`: 第几个坐标
    /// `x:*mut i32`: 返回X坐标
    /// `y:*mut i32`: 返回Y坐标
    ///
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    ///
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let (mut x,mut y) = (0,0);
    ///
    /// let s = dm.FindColorEx(0,0,2000,2000,"123456-000000|abcdef-202020",1.0,0).unwrap();
    /// let count = dm.GetResultCount(s)
    /// for i in 0..count{
    ///     let dm_ret = dm.GetResultPos(s,i,&mut x,&mut y).unwrap();
    /// }
    /// ```
    pub unsafe fn GetResultPos(
        &self,
        str: &str,
        index: i32,
        x: *mut i32,
        y: *mut i32,
    ) -> Result<i32> {
        const NAME: &'static str = "GetResultPos";
        let mut px = VARIANT::default();
        let mut py = VARIANT::default();
        let mut args = [
            Dmsoft::pvarVal(&mut py),
            Dmsoft::pvarVal(&mut px),
            Dmsoft::longVar(index),
            Dmsoft::bstrVal(str),
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

    /// 未在API文档中找到此函数说明
    /// # The function prototype
    /// ```C++
    /// long dmsoft::StrStr(const TCHAR * s,const TCHAR * str)
    /// ```
    pub unsafe fn StrStr(&self, s: &str, str: &str) -> Result<i32> {
        const NAME: &'static str = "StrStr";
        let mut args = [Dmsoft::bstrVal(str), Dmsoft::bstrVal(s)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        Ok(result.Anonymous.lVal)
    }

    /// 未在API文档中找到此函数说明
    /// # The function prototype
    /// ```C++
    /// long dmsoft::SendCommand(const TCHAR * cmd)
    /// ```
    pub unsafe fn SendCommand(&self, cmd: &str) -> Result<i32> {
        const NAME: &'static str = "SendCommand";
        let mut args = [Dmsoft::bstrVal(cmd)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        Ok(result.Anonymous.lVal)
    }

    /// 表示使用哪个字库文件进行识别(index范围:0-99)
    ///
    /// 设置之后，永久生效，除非再次设定
    ///
    /// # The function prototype
    /// ```C++
    /// long dmsoft::UseDict(long index)
    /// ```
    ///
    /// # Args
    /// * `index:i32`: 字库编号(0-99)
    ///
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    ///
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let dm_ret = dm.UseDict(1).unwrap();
    /// ss = dm.Ocr(0,0,2000,2000,"FFFFFF-000000",1.0).unwrap();
    /// dm_ret = dm.UseDict(0).unwrap();
    /// ```
    pub unsafe fn UseDict(&self, index: i32) -> Result<i32> {
        const NAME: &'static str = "UseDict";
        let mut args = [Dmsoft::longVar(index)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        Ok(result.Anonymous.lVal)
    }

    /// 获取注册在系统中的dm.dll的路径.
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::GetBasePath()
    /// ```
    ///
    /// # Args
    /// # Return
    /// `String`: 返回dm.dll所在路径
    ///
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let base_path = dm.GetBasePath().unwrap();
    /// ```
    pub unsafe fn GetBasePath(&self) -> Result<String> {
        const NAME: &'static str = "GetBasePath";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 设置字库的密码,在SetDict前调用,目前的设计是,所有字库通用一个密码.
    /// # The function prototype
    /// ```C++
    /// long dmsoft::SetDictPwd(const TCHAR * pwd)
    /// ```
    /// # Args
    /// * `pwd: &str`: 字库密码
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// dm.SetDictPwd("1234").unwrap();
    /// ```
    /// # Note:
    /// * 如果使用了多字库,所有字库的密码必须一样. 此函数必须在SetDict之前调用,否则会解密失败.
    pub unsafe fn SetDictPwd(&self, pwd: &str) -> Result<i32> {
        const NAME: &'static str = "SetDictPwd";
        let mut args = [Dmsoft::bstrVal(pwd)];
        let result = self.Invoke(NAME, &mut args).unwrap();
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        Ok(result.Anonymous.lVal)
    }

    /// 识别位图中区域(x1,y1,x2,y2)的文字
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::OcrInFile(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * color,double sim)
    /// ```
    /// # Args
    /// * `x1:i32`: 区域的左上X坐标
    /// * `y1:i32`: 区域的左上Y坐标
    /// * `x2:i32`: 区域的右下X坐标
    /// * `y2:i32`: 区域的右下Y坐标
    /// * `pic_name:&str`: 图片文件名
    /// * `color_format:&str`: 颜色格式串. 注意，RGB和HSV,以及灰度格式都支持.
    /// * `sim:f64`: 相似度,取值范围0.1-1.0
    /// # Return
    /// * `String`: 返回识别到的字符串
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let result = dm.OcrInFile(0,0,2000,2000,"test.bmp","000000-000000",1.0).unwrap();
    /// ```
    pub unsafe fn OcrInFile(
        &self,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
        pic_name: &str,
        color: &str,
        sim: f64,
    ) -> Result<String> {
        const NAME: &'static str = "OcrInFile";

        let mut args = [
            Dmsoft::doubleVar(sim),
            Dmsoft::bstrVal(color),
            Dmsoft::bstrVal(pic_name),
            Dmsoft::longVar(y2),
            Dmsoft::longVar(x2),
            Dmsoft::longVar(y1),
            Dmsoft::longVar(x1),
        ];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 抓取指定区域(x1, y1, x2, y2)的图像,保存为file(24位位图)
    /// # The function prototype
    /// ```C++
    /// long dmsoft::Capture(long x1,long y1,long x2,long y2,const TCHAR * file_name)
    /// ```
    /// # Args
    /// * `x1:i32`: 区域的左上X坐标
    /// * `y1:i32`: 区域的左上Y坐标
    /// * `x2:i32`: 区域的右下X坐标
    /// * `y2:i32`: 区域的右下Y坐标
    /// * `file:&str`: 保存的文件名,保存的地方一般为SetPath中设置的目录 当然这里也可以指定全路径名.
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// dm.Capture(0,0,2000,2000,"screen.bmp").unwrap();
    /// ```
    pub unsafe fn Capture(
        &self,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
        file_name: &str,
    ) -> Result<i32> {
        const NAME: &'static str = "Capture";
        let mut args = [
            Dmsoft::bstrVal(file_name),
            Dmsoft::longVar(y2),
            Dmsoft::longVar(x2),
            Dmsoft::longVar(y1),
            Dmsoft::longVar(x1),
        ];
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
    pub unsafe fn KeyPress<'a>(&self, vk: KeyMap<'a>) -> Result<i32> {
        const NAME: &'static str = "KeyPress";
        let mut args = [Dmsoft::longVar(vk.get_id())];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

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
    pub unsafe fn KeyDown<'a>(&self, vk: KeyMap<'a>) -> Result<i32> {
        const NAME: &'static str = "KeyDown";
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
    pub unsafe fn KeyUp<'a>(&self, vk: KeyMap<'a>) -> Result<i32> {
        const NAME: &'static str = "KeyUp";
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
    pub unsafe fn LeftClick(&self)-> Result<i32>{
        const NAME: &'static str = "LeftClick";
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
    pub unsafe fn RightClick(&self) -> Result<i32>{
        const NAME: &'static str = "RightClick";
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
    pub unsafe fn MiddleClick(&self) -> Result<i32>{
        const NAME: &'static str = "MiddleClick";
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
    pub unsafe fn LeftDoubleClick(&self) -> Result<i32>{
        const NAME: &'static str = "LeftDoubleClick";
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
    pub unsafe fn LeftDown(&self) -> Result<i32>{
        const NAME: &'static str = "LeftDown";
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
    pub unsafe fn LeftUp(&self) -> Result<i32>{
        const NAME: &'static str = "LeftUp";
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
    pub unsafe fn RightDown(&self) -> Result<i32>{
        const NAME: &'static str = "RightDown";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        
        Ok(result.Anonymous.lVal)
    }





    // TODO: 其他函数映射
}

/// 辅助函数
#[allow(non_snake_case)]
impl Dmsoft {
    /// 通过COM Function 名称 快捷调用
    /// # Args
    /// * `name:&'static str`: COM Function name
    /// * `args: &mut [VARIANT]` COM Function arguments
    pub unsafe fn Invoke(&self, name: &'static str, args: &mut [VARIANT]) -> Result<VARIANT> {
        let mut map = self.catch.lock().unwrap();
        let rgdispid = *map.entry(name).or_insert_with_key(|key| {
            let mut id = -1;
            let name = HSTRING::from(*key);
            let func_name = PWSTR::from_raw(name.as_ptr() as *mut _);
            // 在调试时解决 expect
            self.obj
                .GetIDsOfNames(ptr::null(), &func_name, 1, LOCALE_USER_DEFAULT, &mut id)
                .expect("调用 GetIDsOfNames 获取ID 异常: ");
            id
        });
        drop(map);
        if rgdispid == -1 {
            return Err(Error::IdError);
        };

        let mut result = VARIANT::default();
        let dispparams = DISPPARAMS {
            rgvarg: args.as_mut_ptr(),
            rgdispidNamedArgs: ptr::null_mut(),
            cArgs: args.len() as u32,
            cNamedArgs: 0,
        };
        if let Err(e) = self.obj.Invoke(
            rgdispid,
            ptr::null(),
            LOCALE_USER_DEFAULT,
            Ole::DISPATCH_METHOD as u16,
            &dispparams,
            &mut result,
            ptr::null_mut(),
            ptr::null_mut(),
        ) {
            return Err(Error::WinError(e));
        };
        Ok(result)
    }

    /// 从 &str 构建一个 VT_BSTR VARIANT
    pub unsafe fn bstrVal(var: &str) -> VARIANT {
        let s = BSTR::from_raw(HSTRING::from(var).as_ptr());
        let mut arg = VARIANT_0_0::default();
        arg.vt = Ole::VT_BSTR.0 as u16;
        arg.Anonymous.bstrVal = ManuallyDrop::new(s);
        VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(arg),
            },
        }
    }

    /// 从 i32 构建一个 VT_I4 VARIANT
    pub unsafe fn longVar(var: i32) -> VARIANT {
        let mut arg = VARIANT_0_0::default();
        arg.vt = Ole::VT_I4.0 as u16;
        arg.Anonymous.lVal = var;
        VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(arg),
            },
        }
    }

    /// VARIANT 指针
    /// pvalVal中存放了另外一个VARIANTTARG的指针。这个被引用的VARIANTARG不能是VT_VARIANT | VT_BYREF类型。
    /// VT_BYREF|VT_VARIANT VARIANT
    pub unsafe fn pvarVal(var: *mut VARIANT) -> VARIANT {
        let mut arg = VARIANT_0_0::default();
        arg.vt = (Ole::VT_BYREF.0 | Ole::VT_VARIANT.0) as u16;
        arg.Anonymous.pvarVal = var;
        VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(arg),
            },
        }
    }

    /// 从 f64 构建一个 VT_R8 VARIANT
    pub unsafe fn doubleVar(var: f64) -> VARIANT {
        let mut arg = VARIANT_0_0::default();
        arg.vt = Ole::VT_R8.0 as u16;
        arg.Anonymous.dblVal = var;
        VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(arg),
            },
        }
    }
}

///
pub struct KeyMap<'a> {
    key_str: &'a str,
    id: i32,
}

impl<'a> KeyMap<'a> {
    ///
    pub fn new(key_str: &'a str, id: i32) -> Self {
        Self { key_str, id }
    }
    ///
    pub fn get_key_str(&self) -> &'a str {
        self.key_str
    }
    ///
    pub fn get_id(&self) -> i32 {
        self.id
    }
}
