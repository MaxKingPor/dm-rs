use std::mem::ManuallyDrop;

use windows::Win32::System::Com::VARIANT;

use crate::{Dmsoft, Result};
#[allow(non_snake_case)]
impl Dmsoft {
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
        static NAME: &str = "Capture";
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

    /// 查找指定区域内的颜色,颜色格式"RRGGBB-DRDGDB",注意,和按键的颜色格式相反
    /// # The function prototype
    /// ```C++
    /// long dmsoft::FindColor(long x1,long y1,long x2,long y2,const TCHAR * color,double sim,long dir,long * x,long * y)
    /// ```
    /// # Args
    /// * `x1:i32`: 区域的左上X坐标
    /// * `y1:i32`: 区域的左上Y坐标
    /// * `x2:i32`: 区域的右下X坐标
    /// * `y2:i32`: 区域的右下Y坐标
    /// * `color:&str`: 颜色 格式为"RRGGBB-DRDGDB",比如"123456-000000|aabbcc-202020". 也可以支持反色模式. 前面加@即可. 比如"@123456-000000|aabbcc-202020". 具体可以看下放注释. 注意，这里只支持RGB颜色.
    /// * `sim:f64`: 相似度,取值范围0.1-1.0
    /// * `dir:i32`: 查找方向
    ///     * 0: 从左到右,从上到下    
    ///     * 1: 从左到右,从下到上   
    ///     * 2: 从右到左,从上到下  
    ///     * 3: 从右到左,从下到上     
    ///     * 4：从中心往外查找   
    ///     * 5: 从上到下,从左到右   
    ///     * 6: 从上到下,从右到左  
    ///     * 7: 从下到上,从左到右  
    ///     * 8: 从下到上,从右到左
    /// * `intX:&mut i32`: 返回X坐标
    /// * `intY:&mut i32`: 返回Y坐标
    /// # Return
    /// `i32`: 0: 没找到 1: 找到
    /// # Examples
    /// ```
    /// let (mut x,mut y) = (0,0);
    /// let dm = Dmsoft::new();
    /// let status = dm.FindColor(0,0,2000,2000,"123456-000000|aabbcc-030303|ddeeff-202020",1.0,0,&mut x,&mut y).unwrap();
    /// ```
    ///
    pub unsafe fn FindColor(
        &self,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
        color: &str,
        sim: f64,
        dir: i32,
        x: &mut i32,
        y: &mut i32,
    ) -> Result<i32> {
        static NAME: &str = "FindColor";
        let mut px = VARIANT::default();
        let mut py = VARIANT::default();
        let mut args = [
            Dmsoft::pvarVal(&mut py),
            Dmsoft::pvarVal(&mut px),
            Dmsoft::longVar(dir),
            Dmsoft::doubleVar(sim),
            Dmsoft::bstrVal(color),
            Dmsoft::longVar(y2),
            Dmsoft::longVar(x2),
            Dmsoft::longVar(y1),
            Dmsoft::longVar(x1),
        ];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        *x = px.Anonymous.Anonymous.Anonymous.lVal;
        *y = px.Anonymous.Anonymous.Anonymous.lVal;

        Ok(result.Anonymous.lVal)
    }

    /// 获取(x,y)的颜色,颜色返回格式"RRGGBB",注意,和按键的颜色格式相反
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::GetColor(long x,long y)
    /// ```
    /// # Args
    /// * `x:i32`: X坐标
    /// * `y:i32`: Y坐标
    /// # Return
    /// `String` 颜色字符串(注意这里都是小写字符，和工具相匹配)
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.GetColor(0,0).unwrap();
    /// ```
    pub unsafe fn GetColor(&self, x: i32, y: i32) -> Result<String> {
        static NAME: &str = "GetColor";
        let mut args = [Dmsoft::longVar(y), Dmsoft::longVar(x)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 获取(x,y)的颜色,颜色返回格式"BBGGRR"
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::GetColorBGR(long x,long y)
    /// ```
    /// # Args
    /// * `x:i32`: X坐标
    /// * `y:i32`: Y坐标
    /// # Return
    /// `String` 颜色字符串(注意这里都是小写字符，和工具相匹配)
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.GetColorBGR(0,0).unwrap();
    /// ```
    pub unsafe fn GetColorBGR(&self, x: i32, y: i32) -> Result<String> {
        static NAME: &str = "GetColorBGR";
        let mut args = [Dmsoft::longVar(y), Dmsoft::longVar(x)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 把RGB的颜色格式转换为BGR(按键格式)
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::RGB2BGR(const TCHAR * rgb_color)
    /// ```
    /// # Args
    /// * `rgb_color:&str`: rgb格式的颜色字符串
    /// # Return
    /// `String` BGR格式的字符串
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.RGB2BGR("00FF00").unwrap();
    /// ```
    pub unsafe fn RGB2BGR(&self, rgb_color: &str) -> Result<String> {
        static NAME: &str = "RGB2BGR";
        let mut args = [Dmsoft::bstrVal(rgb_color)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 把BGR(按键格式)的颜色格式转换为RGB
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::BGR2RGB(const TCHAR * bgr_color)
    /// ```
    /// # Args
    /// * `bgr_color:&str`: bgr格式的颜色字符串
    /// # Return
    /// `String` RGB格式的字符串
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.BGR2RGB("00FF00").unwrap();
    /// ```
    pub unsafe fn BGR2RGB(&self, bgr_color: &str) -> Result<String> {
        static NAME: &str = "BGR2RGB";
        let mut args = [Dmsoft::bstrVal(bgr_color)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }

    /// 比较指定坐标点(x,y)的颜色
    /// # The function prototype
    /// ```C++
    /// long dmsoft::CmpColor(long x,long y,const TCHAR * color,double sim)
    /// ```
    /// # Args
    /// * `x:i32`: X坐标
    /// * `y:i32`: Y坐标
    /// * `color:&str`: 颜色字符串,可以支持偏色,多色,例如 "ffffff-202020|000000-000000" 这个表示白色偏色为202020,和黑色偏色为000000.颜色最多支持10种颜色组合. 注意，这里只支持RGB颜色.
    /// * `sim:f64`: 相似度(0.1-1.0)
    /// # Return
    /// `i32`: 0: 颜色匹配 1: 颜色不匹配
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.CmpColor(200,300,"000000-000000|ff00ff-101010",0.9).unwrap();
    /// ```
    pub unsafe fn CmpColor(&self, x: i32, y: i32, color: &str, sim: f64) -> Result<i32> {
        static NAME: &str = "UnBindWindow";
        let mut args = [
            Dmsoft::doubleVar(sim),
            Dmsoft::bstrVal(color),
            Dmsoft::longVar(y),
            Dmsoft::longVar(x),
        ];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 查找指定区域内的颜色,颜色格式"RRGGBB-DRDGDB",注意,和按键的颜色格式相反
    /// # The function prototype
    /// ```C++
    /// CString dmsoft::FindColorEx(long x1,long y1,long x2,long y2,const TCHAR * color,double sim,long dir)
    /// ```
    /// # Args
    /// * `x1:i32`: 区域的左上X坐标
    /// * `y1:i32`: 区域的左上Y坐标
    /// * `x2:i32`: 区域的右下X坐标
    /// * `y2:i32`: 区域的右下Y坐标
    /// * `color:&str`: 字符串:颜色 格式为"RRGGBB-DRDGDB" 比如"aabbcc-000000|123456-202020".也可以支持反色模式. 前面加@即可. 比如"@123456-000000|aabbcc-202020". 具体可以看下放注释.注意，这里只支持RGB颜色.
    /// * `sim:f64`: 相似度,取值范围0.1-1.0
    /// * `dir:i32`: 查找方向
    ///     * 0: 从左到右,从上到下    
    ///     * 1: 从左到右,从下到上   
    ///     * 2: 从右到左,从上到下  
    ///     * 3: 从右到左,从下到上     
    ///     * 4：从中心往外查找   
    ///     * 5: 从上到下,从左到右   
    ///     * 6: 从上到下,从右到左  
    ///     * 7: 从下到上,从左到右  
    ///     * 8: 从下到上,从右到左
    /// # Return
    /// `String`: 返回所有颜色信息的坐标值,然后通过GetResultCount等接口来解析 (由于内存限制,返回的颜色数量最多为1800个左右)
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.FindColorEx(0,0,2000,2000,"123456-000000|aabbcc-030303|ddeeff-202020",1.0,0).unwrap();
    /// ```
    /// # Note
    /// * 注: 反色模式是指匹配任意一个指定颜色之外的颜色. 比如"@123456|333333". 在匹配时,会匹配除了123456或者333333之外的颜色
    pub unsafe fn FindColorEx(
        &self,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
        color: &str,
        sim: f64,
        dir: i32,
    ) -> Result<String> {
        static NAME: &str = "FindColorEx";
        let mut args = [
            Dmsoft::longVar(dir),
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
}
