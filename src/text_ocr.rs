use std::mem::ManuallyDrop;

use windows::Win32::System::Com::VARIANT;

use crate::{Dmsoft, Result};

#[allow(non_snake_case)]
impl Dmsoft {
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
        static NAME: &str = "Ocr";

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
    /// * `x:&mut i32`: 返回X坐标没找到返回-1
    /// * `y:&mut i32`: 返回Y坐标没找到返回-1
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
        x: &mut i32,
        y: &mut i32,
    ) -> Result<i32> {
        static NAME: &str = "FindStr";
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
        *x = px.Anonymous.Anonymous.Anonymous.lVal;
        *y = py.Anonymous.Anonymous.Anonymous.lVal;

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
        static NAME: &str = "GetResultCount";
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
    /// `x:&mut i32`: 返回X坐标
    /// `y:&mut i32`: 返回Y坐标
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
        x: &mut i32,
        y: &mut i32,
    ) -> Result<i32> {
        static NAME: &str = "GetResultPos";
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

        *x = px.Anonymous.Anonymous.Anonymous.lVal;
        *y = py.Anonymous.Anonymous.Anonymous.lVal;

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
        static NAME: &str = "UseDict";
        let mut args = [Dmsoft::longVar(index)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        Ok(result.Anonymous.lVal)
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
        static NAME: &str = "SetDictPwd";
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
        static NAME: &str = "OcrInFile";

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

    /// 高级用户使用,在识别前,如果待识别区域有多行文字,可以设定行间距,默认的行间距是1,
    ///
    /// 如果根据情况设定,可以提高识别精度。一般不用设定。
    /// # The function prototype
    /// ```C++
    /// long dmsoft::SetMinRowGap(long row_gap)
    /// ```
    /// # Args
    /// * `row_gap:i32`: 最小行间距
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.SetMinRowGap(1) .unwrap();
    /// ```
    pub unsafe fn SetMinRowGap(&self, row_gap: i32) -> Result<i32> {
        static NAME: &str = "SetMinRowGap";
        let mut args = [Dmsoft::longVar(row_gap)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }

    /// 高级用户使用,在识别前,如果待识别区域有多行文字,可以设定列间距,默认的列间距是0,
    ///
    /// 如果根据情况设定,可以提高识别精度。一般不用设定。
    /// # The function prototype
    /// ```C++
    /// long dmsoft::SetMinColGap(long col_gap)
    /// ```
    /// # Args
    /// * `col_gap:i32`: 最小列间距
    /// # Return
    /// `i32`: 0: 失败 1: 成功
    /// # Examples
    /// ```
    /// let dm = Dmsoft::new();
    /// let status = dm.SetMinColGap(1) .unwrap();
    /// ```
    /// # Note
    /// * 此设置如果不为0,那么将不能识别连体字 慎用.
    pub unsafe fn SetMinColGap(&self, col_gap: i32) -> Result<i32> {
        static NAME: &str = "SetMinColGap";
        let mut args = [Dmsoft::longVar(col_gap)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);

        Ok(result.Anonymous.lVal)
    }
}
