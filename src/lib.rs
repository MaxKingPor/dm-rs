#![allow(dead_code)]
#![warn(missing_docs)]
#![doc = include_str!("../README.md")]
#![allow(clippy::missing_safety_doc, clippy::too_many_arguments)]

use std::{
    collections::HashMap, ffi::c_char, mem::ManuallyDrop, os::windows::prelude::OsStrExt,
    path::Path, ptr, sync::Mutex,
};

// use once_cell::sync::OnceCell;
use windows::{
    core::{BSTR, HSTRING, PCWSTR},
    Win32::System::Com::{self, IDispatch, DISPPARAMS, VARIANT, VARIANT_0, VARIANT_0_0},
};

#[cfg(feature = "reg")]
// #[link(name = "DmReg", kind = "static")]
extern "system" {
    /// 免注册使用大漠插件 path: 为 Ascii码表示dm插件所在的路径 status: 0表示STA，1表示MTA
    pub fn SetDllPathA(path: *const c_char, status: usize) -> usize;
    /// 免注册使用大漠插件 path: 为 Unicode码表示插件所在的路径 status: 0表示STA，1表示MTA
    pub fn SetDllPathW(path: *const c_char, status: usize) -> usize;
}

#[cfg(feature = "reg")]
#[allow(missing_docs)]
pub unsafe fn set_dll_path(dm_path: impl AsRef<Path>) -> usize {
    let v: Vec<_> = dm_path
        .as_ref()
        .canonicalize()
        .unwrap()
        .as_os_str()
        .encode_wide()
        // .skip(4)// 去除全路径前面斜杠
        .chain(Some(0))
        .collect();
    SetDllPathW(v.as_ptr() as _, 0)
}

// #[derive(Debug)]
enum VTVar {
    U8(u8),
    I8(i8),
    PU8(*mut u8),
    PI8(*mut i8),
    U16(u16),
    I16(i16),
    PU16(*mut u16),
    PI16(*mut i16),
    U32(u32),
    I32(i32),
    PU32(*mut u32),
    PI32(*mut i32),
    U64(u64),
    I64(i64),
    PU64(*mut u64),
    PI64(*mut i64),
    F32(f32),
    PF32(*mut f32),
    F64(f64),
    PF64(*mut f64),
    CY(CY),
    Pcy(*mut CY),
    String(String),
    Empty,
}
#[derive(Clone, Copy)]
#[repr(C)]
union CY {
    anonymous: Cy0,
    int64: i64,
}

#[repr(C)]
#[derive(Clone, Copy)]

struct Cy0 {
    lo: u32,
    hi: i32,
}

trait CastToVTVar {
    fn to_vtvar(self) -> VTVar;
}
#[allow(unused)]
impl CastToVTVar for VARIANT {
    fn to_vtvar(self) -> VTVar {
        unsafe {
            let anonymous = ManuallyDrop::into_inner(self.Anonymous.Anonymous);
            match anonymous {
                VARIANT_0_0 {
                    vt: Com::VT_EMPTY,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::Empty,
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_EMPTY.0 | Com::VT_BYREF.0 => todo!(),
                VARIANT_0_0 {
                    vt: Com::VT_UI1,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::U8(Anonymous.bVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_UI1.0 | Com::VT_BYREF.0 => VTVar::PU8(Anonymous.pbVal),
                VARIANT_0_0 {
                    vt: Com::VT_UI2,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::U16(Anonymous.uiVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_UI2.0 | Com::VT_BYREF.0 => VTVar::PU16(Anonymous.puiVal),
                VARIANT_0_0 {
                    vt: Com::VT_UI4,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::U32(Anonymous.ulVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_UI4.0 | Com::VT_BYREF.0 => VTVar::PU32(Anonymous.pulVal),
                VARIANT_0_0 {
                    vt: Com::VT_UI8,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::U64(Anonymous.ullVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_UI8.0 | Com::VT_BYREF.0 => VTVar::PU64(Anonymous.pullVal),

                VARIANT_0_0 {
                    vt: Com::VT_UINT,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::U32(Anonymous.uintVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_UINT.0 | Com::VT_BYREF.0 => VTVar::PU32(Anonymous.puintVal),

                VARIANT_0_0 {
                    vt: Com::VT_INT,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::I32(Anonymous.intVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_INT.0 | Com::VT_BYREF.0 => VTVar::PI32(Anonymous.pintVal),

                VARIANT_0_0 {
                    vt: Com::VT_I1,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::I8(Anonymous.cVal as _),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_I1.0 | Com::VT_BYREF.0 => VTVar::PI8(Anonymous.pcVal.0 as _),

                VARIANT_0_0 {
                    vt: Com::VT_I2,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::I16(Anonymous.iVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_I2.0 | Com::VT_BYREF.0 => VTVar::PI16(Anonymous.piVal),
                VARIANT_0_0 {
                    vt: Com::VT_I4,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::I32(Anonymous.lVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_I4.0 | Com::VT_BYREF.0 => VTVar::PI32(Anonymous.plVal),
                VARIANT_0_0 {
                    vt: Com::VT_I8,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::I64(Anonymous.llVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_I8.0 | Com::VT_BYREF.0 => VTVar::PI64(Anonymous.pllVal),

                VARIANT_0_0 {
                    vt: Com::VT_R4,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::F32(Anonymous.fltVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_R4.0 | Com::VT_BYREF.0 => VTVar::PF32(Anonymous.pfltVal),

                VARIANT_0_0 {
                    vt: Com::VT_R8,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::F64(Anonymous.dblVal),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_R8.0 | Com::VT_BYREF.0 => VTVar::PF64(Anonymous.pdblVal),

                VARIANT_0_0 {
                    vt: Com::VT_CY,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => VTVar::CY(*(&Anonymous.cyVal as *const _ as *const CY)),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_CY.0 | Com::VT_BYREF.0 => VTVar::Pcy(Anonymous.pcyVal as _),

                VARIANT_0_0 {
                    vt: Com::VT_BSTR,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => {
                    let str = ManuallyDrop::into_inner(Anonymous.bstrVal);
                    VTVar::String(str.try_into().unwrap())
                }
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_BSTR.0 | Com::VT_BYREF.0 => {
                    #[allow(clippy::drop_ref)]
                    {
                        let str = Anonymous.pbstrVal.as_mut().unwrap();
                        let a = String::try_from(str as &BSTR);
                        drop(str);
                        VTVar::String(a.unwrap())
                    }
                }
                VARIANT_0_0 {
                    vt: Com::VT_DECIMAL,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => todo!(),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_DECIMAL.0 | Com::VT_BYREF.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt: Com::VT_NULL,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => todo!(),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_NULL.0 | Com::VT_BYREF.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt: Com::VT_ERROR,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => todo!(),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_ERROR.0 | Com::VT_BYREF.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt: Com::VT_BOOL,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => todo!(),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_BOOL.0 | Com::VT_BYREF.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt: Com::VT_DATE,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => todo!(),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_DATE.0 | Com::VT_BYREF.0 => {
                    todo!()
                }
                VARIANT_0_0 {
                    vt: Com::VT_DISPATCH,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => todo!(),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_DISPATCH.0 | Com::VT_BYREF.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt: Com::VT_VARIANT,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => panic!("非法。VT_VARIANT必须通过引用传递。"),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_DISPATCH.0 | Com::VT_BYREF.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt: Com::VT_UNKNOWN,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => todo!(),
                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if b == Com::VT_UNKNOWN.0 | Com::VT_BYREF.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt: Com::VARENUM(b),
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } if (b & Com::VT_ARRAY.0) == Com::VT_ARRAY.0 => {
                    todo!()
                }

                VARIANT_0_0 {
                    vt,
                    wReserved1,
                    wReserved2,
                    wReserved3,
                    Anonymous,
                } => {
                    todo!()
                }
            }
        }
    }
}

#[cfg(feature = "keymap")]
pub mod keymap;

/// 在windows-rs 中并未搜索到此参数 使用本地定义 来源:
/// [Windows LOCALE_USER_DEFAULT](https://docs.microsoft.com/en-us/windows/win32/intl/locale-user-default)
pub const LOCALE_USER_DEFAULT: u32 = 0x0400;

/// dm.dmsoft API 绑定
#[derive(Debug)]
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
        Com::CoInitializeEx(None, Default::default()).unwrap();
        let guid = Com::CLSIDFromProgID(windows::w!("dm.dmsoft"))?;
        let r = Com::CoCreateInstance(&guid, None, Com::CLSCTX_ALL)?;
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
        static NAME: &str = "Ver";
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
        static NAME: &str = "SetPath";

        let mut args = [Dmsoft::bstrVal(path)];

        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let a = result.Anonymous.lVal;
        Ok(a)
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
        static NAME: &str = "GetBasePath";
        let result = self.Invoke(NAME, &mut [])?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        let result = ManuallyDrop::into_inner(result.Anonymous.bstrVal);
        Ok(result.try_into().unwrap())
    }
}

mod keyboard_mouse;
mod pic_color;
mod text_ocr;
mod window;
/// # 其他
#[allow(non_snake_case)]
impl Dmsoft {
    /// 未在API文档中找到此函数说明
    /// # The function prototype
    /// ```C++
    /// long dmsoft::StrStr(const TCHAR * s,const TCHAR * str)
    /// ```
    pub unsafe fn StrStr(&self, s: &str, str: &str) -> Result<i32> {
        static NAME: &str = "StrStr";
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
        static NAME: &str = "SendCommand";
        let mut args = [Dmsoft::bstrVal(cmd)];
        let result = self.Invoke(NAME, &mut args)?;
        let result = ManuallyDrop::into_inner(result.Anonymous.Anonymous);
        Ok(result.Anonymous.lVal)
    }

    /// 未在API文档中找到此函数说明
    /// # The function prototype
    /// ```C++
    /// long dmsoft::ShowScrMsg(long x1,long y1,long x2,long y2,const TCHAR * msg,const TCHAR * color)
    /// ```
    pub unsafe fn ShowScrMsg(
        &self,
        x1: i32,
        y1: i32,
        x2: i32,
        y2: i32,
        msg: &str,
        color: &str,
    ) -> Result<i32> {
        static NAME: &str = "ShowScrMsg";
        let mut args = [
            Dmsoft::bstrVal(color),
            Dmsoft::bstrVal(msg),
            Dmsoft::longVar(y2),
            Dmsoft::longVar(x2),
            Dmsoft::longVar(y1),
            Dmsoft::longVar(x1),
        ];

        let result = self.Invoke(NAME, &mut args)?;
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
            let name = HSTRING::from(*key);
            let func_name = PCWSTR::from_raw(name.as_ptr() as *mut _);
            // 在调试时解决 expect
            let mut result = 0;
            self.obj
                .GetIDsOfNames(ptr::null(), &func_name, 1, LOCALE_USER_DEFAULT, &mut result)
                .expect("调用 GetIDsOfNames 获取ID 异常: ");
            result
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
            Com::DISPATCH_METHOD,
            &dispparams,
            Some(&mut result),
            None,
            None,
        ) {
            return Err(Error::WinError(e));
        };
        Ok(result)
    }

    /// 从 &str 构建一个 VT_BSTR VARIANT
    pub unsafe fn bstrVal(var: &str) -> VARIANT {
        let s = BSTR::from_raw(HSTRING::from(var).as_ptr());
        let mut arg = VARIANT_0_0 {
            vt: Com::VT_BSTR,
            ..Default::default()
        };
        arg.Anonymous.bstrVal = ManuallyDrop::new(s);
        VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(arg),
            },
        }
    }

    /// 从 i32 构建一个 VT_I4 VARIANT
    pub unsafe fn longVar(var: i32) -> VARIANT {
        let mut arg = VARIANT_0_0 {
            vt: Com::VT_I4,
            ..Default::default()
        };
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
        let mut arg = VARIANT_0_0 {
            vt: Com::VARENUM(Com::VT_BYREF.0 | Com::VT_VARIANT.0),
            ..Default::default()
        };

        arg.Anonymous.pvarVal = var;
        VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(arg),
            },
        }
    }

    /// 从 f64 构建一个 VT_R8 VARIANT
    pub unsafe fn doubleVar(var: f64) -> VARIANT {
        let mut arg = VARIANT_0_0 {
            vt: Com::VT_R8,
            ..Default::default()
        };

        arg.Anonymous.dblVal = var;
        VARIANT {
            Anonymous: VARIANT_0 {
                Anonymous: ManuallyDrop::new(arg),
            },
        }
    }
}

///
#[derive(Debug)]
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
