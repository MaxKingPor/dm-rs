#ifndef __INCLUDE_OBJ_H__
#define __INCLUDE_OBJ_H__

#if !defined(__cplusplus)
#error C++ compiler required
#endif

#include <objbase.h>

class dmsoft
{
private:
    IDispatch * obj;

public:
    dmsoft();
    virtual ~dmsoft();

    virtual CString Ver();
    virtual long SetPath(const TCHAR * path);
    virtual CString Ocr(long x1,long y1,long x2,long y2,const TCHAR * color,double sim);
    virtual long FindStr(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,long * x,long * y);
    virtual long GetResultCount(const TCHAR * str);
    virtual long GetResultPos(const TCHAR * str,long index,long * x,long * y);
    virtual long StrStr(const TCHAR * s,const TCHAR * str);
    virtual long SendCommand(const TCHAR * cmd);
    virtual long UseDict(long index);
    virtual CString GetBasePath();
    virtual long SetDictPwd(const TCHAR * pwd);
    virtual CString OcrInFile(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * color,double sim);
    virtual long Capture(long x1,long y1,long x2,long y2,const TCHAR * file_name);
    virtual long KeyPress(long vk);
    virtual long KeyDown(long vk);
    virtual long KeyUp(long vk);
    virtual long LeftClick();
    virtual long RightClick();
    virtual long MiddleClick();
    virtual long LeftDoubleClick();
    virtual long LeftDown();
    virtual long LeftUp();
    virtual long RightDown();
    virtual long RightUp();
    virtual long MoveTo(long x,long y);
    virtual long MoveR(long rx,long ry);
    virtual CString GetColor(long x,long y);
    virtual CString GetColorBGR(long x,long y);
    virtual CString RGB2BGR(const TCHAR * rgb_color);
    virtual CString BGR2RGB(const TCHAR * bgr_color);
    virtual long UnBindWindow();
    virtual long CmpColor(long x,long y,const TCHAR * color,double sim);
    virtual long ClientToScreen(long hwnd,long * x,long * y);
    virtual long ScreenToClient(long hwnd,long * x,long * y);
    virtual long ShowScrMsg(long x1,long y1,long x2,long y2,const TCHAR * msg,const TCHAR * color);
    virtual long SetMinRowGap(long row_gap);
    virtual long SetMinColGap(long col_gap);
    virtual long FindColor(long x1,long y1,long x2,long y2,const TCHAR * color,double sim,long dir,long * x,long * y);
    virtual CString FindColorEx(long x1,long y1,long x2,long y2,const TCHAR * color,double sim,long dir);
    virtual long SetWordLineHeight(long line_height);
    virtual long SetWordGap(long word_gap);
    virtual long SetRowGapNoDict(long row_gap);
    virtual long SetColGapNoDict(long col_gap);
    virtual long SetWordLineHeightNoDict(long line_height);
    virtual long SetWordGapNoDict(long word_gap);
    virtual long GetWordResultCount(const TCHAR * str);
    virtual long GetWordResultPos(const TCHAR * str,long index,long * x,long * y);
    virtual CString GetWordResultStr(const TCHAR * str,long index);
    virtual CString GetWords(long x1,long y1,long x2,long y2,const TCHAR * color,double sim);
    virtual CString GetWordsNoDict(long x1,long y1,long x2,long y2,const TCHAR * color);
    virtual long SetShowErrorMsg(long show);
    virtual long GetClientSize(long hwnd,long * width,long * height);
    virtual long MoveWindow(long hwnd,long x,long y);
    virtual CString GetColorHSV(long x,long y);
    virtual CString GetAveRGB(long x1,long y1,long x2,long y2);
    virtual CString GetAveHSV(long x1,long y1,long x2,long y2);
    virtual long GetForegroundWindow();
    virtual long GetForegroundFocus();
    virtual long GetMousePointWindow();
    virtual long GetPointWindow(long x,long y);
    virtual CString EnumWindow(long parent,const TCHAR * title,const TCHAR * class_name,long filter);
    virtual long GetWindowState(long hwnd,long flag);
    virtual long GetWindow(long hwnd,long flag);
    virtual long GetSpecialWindow(long flag);
    virtual long SetWindowText(long hwnd,const TCHAR * text);
    virtual long SetWindowSize(long hwnd,long width,long height);
    virtual long GetWindowRect(long hwnd,long * x1,long * y1,long * x2,long * y2);
    virtual CString GetWindowTitle(long hwnd);
    virtual CString GetWindowClass(long hwnd);
    virtual long SetWindowState(long hwnd,long flag);
    virtual long CreateFoobarRect(long hwnd,long x,long y,long w,long h);
    virtual long CreateFoobarRoundRect(long hwnd,long x,long y,long w,long h,long rw,long rh);
    virtual long CreateFoobarEllipse(long hwnd,long x,long y,long w,long h);
    virtual long CreateFoobarCustom(long hwnd,long x,long y,const TCHAR * pic,const TCHAR * trans_color,double sim);
    virtual long FoobarFillRect(long hwnd,long x1,long y1,long x2,long y2,const TCHAR * color);
    virtual long FoobarDrawText(long hwnd,long x,long y,long w,long h,const TCHAR * text,const TCHAR * color,long align);
    virtual long FoobarDrawPic(long hwnd,long x,long y,const TCHAR * pic,const TCHAR * trans_color);
    virtual long FoobarUpdate(long hwnd);
    virtual long FoobarLock(long hwnd);
    virtual long FoobarUnlock(long hwnd);
    virtual long FoobarSetFont(long hwnd,const TCHAR * font_name,long size,long flag);
    virtual long FoobarTextRect(long hwnd,long x,long y,long w,long h);
    virtual long FoobarPrintText(long hwnd,const TCHAR * text,const TCHAR * color);
    virtual long FoobarClearText(long hwnd);
    virtual long FoobarTextLineGap(long hwnd,long gap);
    virtual long Play(const TCHAR * file_name);
    virtual long FaqCapture(long x1,long y1,long x2,long y2,long quality,long delay,long time);
    virtual long FaqRelease(long handle);
    virtual CString FaqSend(const TCHAR * server,long handle,long request_type,long time_out);
    virtual long Beep(long fre,long delay);
    virtual long FoobarClose(long hwnd);
    virtual long MoveDD(long dx,long dy);
    virtual long FaqGetSize(long handle);
    virtual long LoadPic(const TCHAR * pic_name);
    virtual long FreePic(const TCHAR * pic_name);
    virtual long GetScreenData(long x1,long y1,long x2,long y2);
    virtual long FreeScreenData(long handle);
    virtual long WheelUp();
    virtual long WheelDown();
    virtual long SetMouseDelay(const TCHAR * tpe,long delay);
    virtual long SetKeypadDelay(const TCHAR * tpe,long delay);
    virtual CString GetEnv(long index,const TCHAR * name);
    virtual long SetEnv(long index,const TCHAR * name,const TCHAR * value);
    virtual long SendString(long hwnd,const TCHAR * str);
    virtual long DelEnv(long index,const TCHAR * name);
    virtual CString GetPath();
    virtual long SetDict(long index,const TCHAR * dict_name);
    virtual long FindPic(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,double sim,long dir,long * x,long * y);
    virtual CString FindPicEx(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,double sim,long dir);
    virtual long SetClientSize(long hwnd,long width,long height);
    virtual LONGLONG ReadInt(long hwnd,const TCHAR * addr,long tpe);
    virtual float ReadFloat(long hwnd,const TCHAR * addr);
    virtual double ReadDouble(long hwnd,const TCHAR * addr);
    virtual CString FindInt(long hwnd,const TCHAR * addr_range,LONGLONG int_value_min,LONGLONG int_value_max,long tpe);
    virtual CString FindFloat(long hwnd,const TCHAR * addr_range,float float_value_min,float float_value_max);
    virtual CString FindDouble(long hwnd,const TCHAR * addr_range,double double_value_min,double double_value_max);
    virtual CString FindString(long hwnd,const TCHAR * addr_range,const TCHAR * string_value,long tpe);
    virtual LONGLONG GetModuleBaseAddr(long hwnd,const TCHAR * module_name);
    virtual CString MoveToEx(long x,long y,long w,long h);
    virtual CString MatchPicName(const TCHAR * pic_name);
    virtual long AddDict(long index,const TCHAR * dict_info);
    virtual long EnterCri();
    virtual long LeaveCri();
    virtual long WriteInt(long hwnd,const TCHAR * addr,long tpe,LONGLONG v);
    virtual long WriteFloat(long hwnd,const TCHAR * addr,float v);
    virtual long WriteDouble(long hwnd,const TCHAR * addr,double v);
    virtual long WriteString(long hwnd,const TCHAR * addr,long tpe,const TCHAR * v);
    virtual long AsmAdd(const TCHAR * asm_ins);
    virtual long AsmClear();
    virtual LONGLONG AsmCall(long hwnd,long mode);
    virtual long FindMultiColor(long x1,long y1,long x2,long y2,const TCHAR * first_color,const TCHAR * offset_color,double sim,long dir,long * x,long * y);
    virtual CString FindMultiColorEx(long x1,long y1,long x2,long y2,const TCHAR * first_color,const TCHAR * offset_color,double sim,long dir);
    virtual CString Assemble(LONGLONG base_addr,long is_64bit);
    virtual CString DisAssemble(const TCHAR * asm_code,LONGLONG base_addr,long is_64bit);
    virtual long SetWindowTransparent(long hwnd,long v);
    virtual CString ReadData(long hwnd,const TCHAR * addr,long length);
    virtual long WriteData(long hwnd,const TCHAR * addr,const TCHAR * data);
    virtual CString FindData(long hwnd,const TCHAR * addr_range,const TCHAR * data);
    virtual long SetPicPwd(const TCHAR * pwd);
    virtual long Log(const TCHAR * info);
    virtual CString FindStrE(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim);
    virtual CString FindColorE(long x1,long y1,long x2,long y2,const TCHAR * color,double sim,long dir);
    virtual CString FindPicE(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,double sim,long dir);
    virtual CString FindMultiColorE(long x1,long y1,long x2,long y2,const TCHAR * first_color,const TCHAR * offset_color,double sim,long dir);
    virtual long SetExactOcr(long exact_ocr);
    virtual CString ReadString(long hwnd,const TCHAR * addr,long tpe,long length);
    virtual long FoobarTextPrintDir(long hwnd,long dir);
    virtual CString OcrEx(long x1,long y1,long x2,long y2,const TCHAR * color,double sim);
    virtual long SetDisplayInput(const TCHAR * mode);
    virtual long GetTime();
    virtual long GetScreenWidth();
    virtual long GetScreenHeight();
    virtual long BindWindowEx(long hwnd,const TCHAR * display,const TCHAR * mouse,const TCHAR * keypad,const TCHAR * public_desc,long mode);
    virtual CString GetDiskSerial(long index);
    virtual CString Md5(const TCHAR * str);
    virtual CString GetMac();
    virtual long ActiveInputMethod(long hwnd,const TCHAR * id);
    virtual long CheckInputMethod(long hwnd,const TCHAR * id);
    virtual long FindInputMethod(const TCHAR * id);
    virtual long GetCursorPos(long * x,long * y);
    virtual long BindWindow(long hwnd,const TCHAR * display,const TCHAR * mouse,const TCHAR * keypad,long mode);
    virtual long FindWindow(const TCHAR * class_name,const TCHAR * title_name);
    virtual long GetScreenDepth();
    virtual long SetScreen(long width,long height,long depth);
    virtual long ExitOs(long tpe);
    virtual CString GetDir(long tpe);
    virtual long GetOsType();
    virtual long FindWindowEx(long parent,const TCHAR * class_name,const TCHAR * title_name);
    virtual long SetExportDict(long index,const TCHAR * dict_name);
    virtual CString GetCursorShape();
    virtual long DownCpu(long tpe,long rate);
    virtual CString GetCursorSpot();
    virtual long SendString2(long hwnd,const TCHAR * str);
    virtual long FaqPost(const TCHAR * server,long handle,long request_type,long time_out);
    virtual CString FaqFetch();
    virtual CString FetchWord(long x1,long y1,long x2,long y2,const TCHAR * color,const TCHAR * word);
    virtual long CaptureJpg(long x1,long y1,long x2,long y2,const TCHAR * file_name,long quality);
    virtual long FindStrWithFont(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,const TCHAR * font_name,long font_size,long flag,long * x,long * y);
    virtual CString FindStrWithFontE(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,const TCHAR * font_name,long font_size,long flag);
    virtual CString FindStrWithFontEx(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,const TCHAR * font_name,long font_size,long flag);
    virtual CString GetDictInfo(const TCHAR * str,const TCHAR * font_name,long font_size,long flag);
    virtual long SaveDict(long index,const TCHAR * file_name);
    virtual long GetWindowProcessId(long hwnd);
    virtual CString GetWindowProcessPath(long hwnd);
    virtual long LockInput(long locks);
    virtual CString GetPicSize(const TCHAR * pic_name);
    virtual long GetID();
    virtual long CapturePng(long x1,long y1,long x2,long y2,const TCHAR * file_name);
    virtual long CaptureGif(long x1,long y1,long x2,long y2,const TCHAR * file_name,long delay,long time);
    virtual long ImageToBmp(const TCHAR * pic_name,const TCHAR * bmp_name);
    virtual long FindStrFast(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,long * x,long * y);
    virtual CString FindStrFastEx(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim);
    virtual CString FindStrFastE(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim);
    virtual long EnableDisplayDebug(long enable_debug);
    virtual long CapturePre(const TCHAR * file_name);
    virtual long RegEx(const TCHAR * code,const TCHAR * Ver,const TCHAR * ip);
    virtual CString GetMachineCode();
    virtual long SetClipboard(const TCHAR * data);
    virtual CString GetClipboard();
    virtual long GetNowDict();
    virtual long Is64Bit();
    virtual long GetColorNum(long x1,long y1,long x2,long y2,const TCHAR * color,double sim);
    virtual CString EnumWindowByProcess(const TCHAR * process_name,const TCHAR * title,const TCHAR * class_name,long filter);
    virtual long GetDictCount(long index);
    virtual long GetLastError();
    virtual CString GetNetTime();
    virtual long EnableGetColorByCapture(long en);
    virtual long CheckUAC();
    virtual long SetUAC(long uac);
    virtual long DisableFontSmooth();
    virtual long CheckFontSmooth();
    virtual long SetDisplayAcceler(long level);
    virtual long FindWindowByProcess(const TCHAR * process_name,const TCHAR * class_name,const TCHAR * title_name);
    virtual long FindWindowByProcessId(long process_id,const TCHAR * class_name,const TCHAR * title_name);
    virtual CString ReadIni(const TCHAR * section,const TCHAR * key,const TCHAR * file_name);
    virtual long WriteIni(const TCHAR * section,const TCHAR * key,const TCHAR * v,const TCHAR * file_name);
    virtual long RunApp(const TCHAR * path,long mode);
    virtual long delay(long mis);
    virtual long FindWindowSuper(const TCHAR * spec1,long flag1,long type1,const TCHAR * spec2,long flag2,long type2);
    virtual CString ExcludePos(const TCHAR * all_pos,long tpe,long x1,long y1,long x2,long y2);
    virtual CString FindNearestPos(const TCHAR * all_pos,long tpe,long x,long y);
    virtual CString SortPosDistance(const TCHAR * all_pos,long tpe,long x,long y);
    virtual long FindPicMem(long x1,long y1,long x2,long y2,const TCHAR * pic_info,const TCHAR * delta_color,double sim,long dir,long * x,long * y);
    virtual CString FindPicMemEx(long x1,long y1,long x2,long y2,const TCHAR * pic_info,const TCHAR * delta_color,double sim,long dir);
    virtual CString FindPicMemE(long x1,long y1,long x2,long y2,const TCHAR * pic_info,const TCHAR * delta_color,double sim,long dir);
    virtual CString AppendPicAddr(const TCHAR * pic_info,long addr,long size);
    virtual long WriteFile(const TCHAR * file_name,const TCHAR * content);
    virtual long Stop(long id);
    virtual long SetDictMem(long index,long addr,long size);
    virtual CString GetNetTimeSafe();
    virtual long ForceUnBindWindow(long hwnd);
    virtual CString ReadIniPwd(const TCHAR * section,const TCHAR * key,const TCHAR * file_name,const TCHAR * pwd);
    virtual long WriteIniPwd(const TCHAR * section,const TCHAR * key,const TCHAR * v,const TCHAR * file_name,const TCHAR * pwd);
    virtual long DecodeFile(const TCHAR * file_name,const TCHAR * pwd);
    virtual long KeyDownChar(const TCHAR * key_str);
    virtual long KeyUpChar(const TCHAR * key_str);
    virtual long KeyPressChar(const TCHAR * key_str);
    virtual long KeyPressStr(const TCHAR * key_str,long delay);
    virtual long EnableKeypadPatch(long en);
    virtual long EnableKeypadSync(long en,long time_out);
    virtual long EnableMouseSync(long en,long time_out);
    virtual long DmGuard(long en,const TCHAR * tpe);
    virtual long FaqCaptureFromFile(long x1,long y1,long x2,long y2,const TCHAR * file_name,long quality);
    virtual CString FindIntEx(long hwnd,const TCHAR * addr_range,LONGLONG int_value_min,LONGLONG int_value_max,long tpe,long steps,long multi_thread,long mode);
    virtual CString FindFloatEx(long hwnd,const TCHAR * addr_range,float float_value_min,float float_value_max,long steps,long multi_thread,long mode);
    virtual CString FindDoubleEx(long hwnd,const TCHAR * addr_range,double double_value_min,double double_value_max,long steps,long multi_thread,long mode);
    virtual CString FindStringEx(long hwnd,const TCHAR * addr_range,const TCHAR * string_value,long tpe,long steps,long multi_thread,long mode);
    virtual CString FindDataEx(long hwnd,const TCHAR * addr_range,const TCHAR * data,long steps,long multi_thread,long mode);
    virtual long EnableRealMouse(long en,long mousedelay,long mousestep);
    virtual long EnableRealKeypad(long en);
    virtual long SendStringIme(const TCHAR * str);
    virtual long FoobarDrawLine(long hwnd,long x1,long y1,long x2,long y2,const TCHAR * color,long style,long width);
    virtual CString FindStrEx(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim);
    virtual long IsBind(long hwnd);
    virtual long SetDisplayDelay(long t);
    virtual long GetDmCount();
    virtual long DisableScreenSave();
    virtual long DisablePowerSave();
    virtual long SetMemoryHwndAsProcessId(long en);
    virtual long FindShape(long x1,long y1,long x2,long y2,const TCHAR * offset_color,double sim,long dir,long * x,long * y);
    virtual CString FindShapeE(long x1,long y1,long x2,long y2,const TCHAR * offset_color,double sim,long dir);
    virtual CString FindShapeEx(long x1,long y1,long x2,long y2,const TCHAR * offset_color,double sim,long dir);
    virtual CString FindStrS(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,long * x,long * y);
    virtual CString FindStrExS(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim);
    virtual CString FindStrFastS(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim,long * x,long * y);
    virtual CString FindStrFastExS(long x1,long y1,long x2,long y2,const TCHAR * str,const TCHAR * color,double sim);
    virtual CString FindPicS(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,double sim,long dir,long * x,long * y);
    virtual CString FindPicExS(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,double sim,long dir);
    virtual long ClearDict(long index);
    virtual CString GetMachineCodeNoMac();
    virtual long GetClientRect(long hwnd,long * x1,long * y1,long * x2,long * y2);
    virtual long EnableFakeActive(long en);
    virtual long GetScreenDataBmp(long x1,long y1,long x2,long y2,long * data,long * size);
    virtual long EncodeFile(const TCHAR * file_name,const TCHAR * pwd);
    virtual CString GetCursorShapeEx(long tpe);
    virtual long FaqCancel();
    virtual CString IntToData(LONGLONG int_value,long tpe);
    virtual CString FloatToData(float float_value);
    virtual CString DoubleToData(double double_value);
    virtual CString StringToData(const TCHAR * string_value,long tpe);
    virtual long SetMemoryFindResultToFile(const TCHAR * file_name);
    virtual long EnableBind(long en);
    virtual long SetSimMode(long mode);
    virtual long LockMouseRect(long x1,long y1,long x2,long y2);
    virtual long SendPaste(long hwnd);
    virtual long IsDisplayDead(long x1,long y1,long x2,long y2,long t);
    virtual long GetKeyState(long vk);
    virtual long CopyFile(const TCHAR * src_file,const TCHAR * dst_file,long over);
    virtual long IsFileExist(const TCHAR * file_name);
    virtual long DeleteFile(const TCHAR * file_name);
    virtual long MoveFile(const TCHAR * src_file,const TCHAR * dst_file);
    virtual long CreateFolder(const TCHAR * folder_name);
    virtual long DeleteFolder(const TCHAR * folder_name);
    virtual long GetFileLength(const TCHAR * file_name);
    virtual CString ReadFile(const TCHAR * file_name);
    virtual long WaitKey(long key_code,long time_out);
    virtual long DeleteIni(const TCHAR * section,const TCHAR * key,const TCHAR * file_name);
    virtual long DeleteIniPwd(const TCHAR * section,const TCHAR * key,const TCHAR * file_name,const TCHAR * pwd);
    virtual long EnableSpeedDx(long en);
    virtual long EnableIme(long en);
    virtual long Reg(const TCHAR * code,const TCHAR * Ver);
    virtual CString SelectFile();
    virtual CString SelectDirectory();
    virtual long LockDisplay(long locks);
    virtual long FoobarSetSave(long hwnd,const TCHAR * file_name,long en,const TCHAR * header);
    virtual CString EnumWindowSuper(const TCHAR * spec1,long flag1,long type1,const TCHAR * spec2,long flag2,long type2,long sort);
    virtual long DownloadFile(const TCHAR * url,const TCHAR * save_file,long timeout);
    virtual long EnableKeypadMsg(long en);
    virtual long EnableMouseMsg(long en);
    virtual long RegNoMac(const TCHAR * code,const TCHAR * Ver);
    virtual long RegExNoMac(const TCHAR * code,const TCHAR * Ver,const TCHAR * ip);
    virtual long SetEnumWindowDelay(long delay);
    virtual long FindMulColor(long x1,long y1,long x2,long y2,const TCHAR * color,double sim);
    virtual CString GetDict(long index,long font_index);
    virtual long GetBindWindow();
    virtual long FoobarStartGif(long hwnd,long x,long y,const TCHAR * pic_name,long repeat_limit,long delay);
    virtual long FoobarStopGif(long hwnd,long x,long y,const TCHAR * pic_name);
    virtual long FreeProcessMemory(long hwnd);
    virtual CString ReadFileData(const TCHAR * file_name,long start_pos,long end_pos);
    virtual LONGLONG VirtualAllocEx(long hwnd,LONGLONG addr,long size,long tpe);
    virtual long VirtualFreeEx(long hwnd,LONGLONG addr);
    virtual CString GetCommandLine(long hwnd);
    virtual long TerminateProcess(long pid);
    virtual CString GetNetTimeByIp(const TCHAR * ip);
    virtual CString EnumProcess(const TCHAR * name);
    virtual CString GetProcessInfo(long pid);
    virtual LONGLONG ReadIntAddr(long hwnd,LONGLONG addr,long tpe);
    virtual CString ReadDataAddr(long hwnd,LONGLONG addr,long length);
    virtual double ReadDoubleAddr(long hwnd,LONGLONG addr);
    virtual float ReadFloatAddr(long hwnd,LONGLONG addr);
    virtual CString ReadStringAddr(long hwnd,LONGLONG addr,long tpe,long length);
    virtual long WriteDataAddr(long hwnd,LONGLONG addr,const TCHAR * data);
    virtual long WriteDoubleAddr(long hwnd,LONGLONG addr,double v);
    virtual long WriteFloatAddr(long hwnd,LONGLONG addr,float v);
    virtual long WriteIntAddr(long hwnd,LONGLONG addr,long tpe,LONGLONG v);
    virtual long WriteStringAddr(long hwnd,LONGLONG addr,long tpe,const TCHAR * v);
    virtual long Delays(long min_s,long max_s);
    virtual long FindColorBlock(long x1,long y1,long x2,long y2,const TCHAR * color,double sim,long count,long width,long height,long * x,long * y);
    virtual CString FindColorBlockEx(long x1,long y1,long x2,long y2,const TCHAR * color,double sim,long count,long width,long height);
    virtual long OpenProcess(long pid);
    virtual CString EnumIniSection(const TCHAR * file_name);
    virtual CString EnumIniSectionPwd(const TCHAR * file_name,const TCHAR * pwd);
    virtual CString EnumIniKey(const TCHAR * section,const TCHAR * file_name);
    virtual CString EnumIniKeyPwd(const TCHAR * section,const TCHAR * file_name,const TCHAR * pwd);
    virtual long SwitchBindWindow(long hwnd);
    virtual long InitCri();
    virtual long SendStringIme2(long hwnd,const TCHAR * str,long mode);
    virtual CString EnumWindowByProcessId(long pid,const TCHAR * title,const TCHAR * class_name,long filter);
    virtual CString GetDisplayInfo();
    virtual long EnableFontSmooth();
    virtual CString OcrExOne(long x1,long y1,long x2,long y2,const TCHAR * color,double sim);
    virtual long SetAero(long en);
    virtual long FoobarSetTrans(long hwnd,long trans,const TCHAR * color,double sim);
    virtual long EnablePicCache(long en);
    virtual long FaqIsPosted();
    virtual long LoadPicByte(long addr,long size,const TCHAR * name);
    virtual long MiddleDown();
    virtual long MiddleUp();
    virtual long FaqCaptureString(const TCHAR * str);
    virtual long VirtualProtectEx(long hwnd,LONGLONG addr,long size,long tpe,long old_protect);
    virtual long SetMouseSpeed(long speed);
    virtual long GetMouseSpeed();
    virtual long EnableMouseAccuracy(long en);
    virtual long SetExcludeRegion(long tpe,const TCHAR * info);
    virtual long EnableShareDict(long en);
    virtual long DisableCloseDisplayAndSleep();
    virtual long Int64ToInt32(LONGLONG v);
    virtual long GetLocale();
    virtual long SetLocale();
    virtual long ReadDataToBin(long hwnd,const TCHAR * addr,long length);
    virtual long WriteDataFromBin(long hwnd,const TCHAR * addr,long data,long length);
    virtual long ReadDataAddrToBin(long hwnd,LONGLONG addr,long length);
    virtual long WriteDataAddrFromBin(long hwnd,LONGLONG addr,long data,long length);
    virtual long SetParam64ToPointer();
    virtual long GetDPI();
    virtual long SetDisplayRefreshDelay(long t);
    virtual long IsFolderExist(const TCHAR * folder);
    virtual long GetCpuType();
    virtual long ReleaseRef();
    virtual long SetExitThread(long en);
    virtual long GetFps();
    virtual CString VirtualQueryEx(long hwnd,LONGLONG addr,long pmbi);
    virtual LONGLONG AsmCallEx(long hwnd,long mode,const TCHAR * base_addr);
    virtual LONGLONG GetRemoteApiAddress(long hwnd,LONGLONG base_addr,const TCHAR * fun_name);
    virtual CString ExecuteCmd(const TCHAR * cmd,const TCHAR * current_dir,long time_out);
    virtual long SpeedNormalGraphic(long en);
    virtual long UnLoadDriver();
    virtual long GetOsBuildNumber();
    virtual long HackSpeed(double rate);
    virtual CString GetRealPath(const TCHAR * path);
    virtual long ShowTaskBarIcon(long hwnd,long is_show);
    virtual long AsmSetTimeout(long time_out,long param);
    virtual CString DmGuardParams(const TCHAR * cmd,const TCHAR * sub_cmd,const TCHAR * param);
    virtual long GetModuleSize(long hwnd,const TCHAR * module_name);
    virtual long IsSurrpotVt();
    virtual CString GetDiskModel(long index);
    virtual CString GetDiskReversion(long index);
    virtual long EnableFindPicMultithread(long en);
    virtual long GetCpuUsage();
    virtual long GetMemoryUsage();
    virtual CString Hex32(long v);
    virtual CString Hex64(LONGLONG v);
    virtual long GetWindowThreadId(long hwnd);
    virtual long DmGuardExtract(const TCHAR * tpe,const TCHAR * path);
    virtual long DmGuardLoadCustom(const TCHAR * tpe,const TCHAR * path);
    virtual long SetShowAsmErrorMsg(long show);
    virtual CString GetSystemInfo(const TCHAR * tpe,long method);
    virtual long SetFindPicMultithreadCount(long count);
    virtual long FindPicSim(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,long sim,long dir,long * x,long * y);
    virtual CString FindPicSimEx(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,long sim,long dir);
    virtual long FindPicSimMem(long x1,long y1,long x2,long y2,const TCHAR * pic_info,const TCHAR * delta_color,long sim,long dir,long * x,long * y);
    virtual CString FindPicSimMemEx(long x1,long y1,long x2,long y2,const TCHAR * pic_info,const TCHAR * delta_color,long sim,long dir);
    virtual CString FindPicSimE(long x1,long y1,long x2,long y2,const TCHAR * pic_name,const TCHAR * delta_color,long sim,long dir);
    virtual CString FindPicSimMemE(long x1,long y1,long x2,long y2,const TCHAR * pic_info,const TCHAR * delta_color,long sim,long dir);
    virtual long SetInputDm(long input_dm,long rx,long ry);
};

#endif