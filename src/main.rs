#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::mem::{size_of, zeroed};
use std::ptr::{null, null_mut};
use std::ffi::c_void;
use windows_sys::core::GUID;
use windows_sys::Win32::Foundation::{
    CloseHandle, GetLastError, HINSTANCE, HWND, INVALID_HANDLE_VALUE, LPARAM, LRESULT,
    RECT, RPC_E_CHANGED_MODE, SYSTEMTIME, WPARAM, GENERIC_READ, GENERIC_WRITE,
};
use windows_sys::Win32::Globalization::{
    CP_ACP, CP_UTF8, GetDateFormatW, GetTimeFormatW, GetUserDefaultUILanguage, LOCALE_USER_DEFAULT,
    MultiByteToWideChar, WideCharToMultiByte,
};
use windows_sys::Win32::Graphics::Gdi::{
    CreateFontIndirectW, DeleteObject, GetStockObject, GetSysColorBrush, COLOR_3DFACE, COLOR_WINDOW,
    DEFAULT_GUI_FONT, HBRUSH, HFONT, LOGFONTW,
};
use windows_sys::Win32::Storage::FileSystem::{
    CreateFileW, ReadFile, WriteFile, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL,
    FILE_SHARE_DELETE, FILE_SHARE_READ, FILE_SHARE_WRITE, OPEN_EXISTING,
};
use windows_sys::Win32::System::LibraryLoader::GetModuleHandleW;
use windows_sys::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, CLSCTX_INPROC_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows_sys::Win32::System::SystemInformation::GetLocalTime;
use windows_sys::Win32::System::SystemServices::LANG_JAPANESE;
use windows_sys::Win32::UI::Controls::{
    InitCommonControlsEx, EM_CANUNDO, EM_GETMODIFY, EM_GETSEL, EM_LINEFROMCHAR,
    EM_GETLINECOUNT, EM_LINEINDEX, EM_REPLACESEL, EM_SETMODIFY, EM_SETSEL, ICC_BAR_CLASSES,
    INITCOMMONCONTROLSEX, SB_SETPARTS, SB_SETTEXTW, STATUSCLASSNAMEW,
};
use windows_sys::Win32::UI::Controls::Dialogs::{
    ChooseFontW, FindTextW, ReplaceTextW, CF_INITTOLOGFONTSTRUCT, CF_SCREENFONTS, CHOOSEFONTW,
    FINDMSGSTRINGW, FINDREPLACEW, FR_DIALOGTERM, FR_DOWN, FR_FINDNEXT, FR_MATCHCASE, FR_REPLACE,
    FR_REPLACEALL, FR_WHOLEWORD,
};
use windows_sys::Win32::UI::Input::KeyboardAndMouse::{SetFocus, VK_F3, VK_F5};
use windows_sys::Win32::UI::Shell::Common::COMDLG_FILTERSPEC;
use windows_sys::Win32::UI::Shell::{
    DragAcceptFiles, DragFinish, DragQueryFileW, FILEOPENDIALOGOPTIONS,
    FOS_FILEMUSTEXIST, FOS_FORCEFILESYSTEM, FOS_OVERWRITEPROMPT, HDROP, SIGDN_FILESYSPATH,
};
use windows_sys::Win32::UI::WindowsAndMessaging::{
    AppendMenuW, CheckMenuItem, CreateAcceleratorTableW, CreateMenu, CreatePopupMenu, CreateWindowExW,
    DefWindowProcW, DestroyAcceleratorTable, DestroyWindow, DispatchMessageW, GetMessageW,
    GetMenu, GetWindowLongPtrW, GetWindowTextLengthW, GetWindowTextW, IsDialogMessageW, LoadCursorW, MessageBoxW,
    PostQuitMessage, RegisterClassW, SendMessageW, SetMenu, SetWindowLongPtrW, SetWindowTextW,
    ShowWindow, TranslateAcceleratorW, TranslateMessage, RegisterWindowMessageW, ACCEL, CREATESTRUCTW, CS_HREDRAW,
    CS_VREDRAW, CW_USEDEFAULT, FVIRTKEY, FCONTROL, GWLP_USERDATA, HMENU, IDC_ARROW, MB_ICONERROR,
    MB_ICONWARNING, MB_OK, MB_YESNOCANCEL, MF_POPUP, MF_SEPARATOR, MF_STRING, MSG, SW_SHOW,
    SW_HIDE, EN_CHANGE, IDOK, IDCANCEL, MF_BYCOMMAND, MF_CHECKED, MF_UNCHECKED,
    WM_CLEAR, WM_CLOSE, WM_COMMAND, WM_COPY, WM_CREATE, WM_CUT, WM_DESTROY, WM_NCDESTROY,
    WM_DROPFILES, WM_PASTE, WM_SETFONT, WM_SETFOCUS, WM_SIZE, WM_TIMER, WM_UNDO, WNDCLASSW,
    WS_CHILD, WS_EX_CLIENTEDGE, WS_HSCROLL, WS_OVERLAPPEDWINDOW, WS_VISIBLE, WS_VSCROLL,
    WS_CAPTION, WS_POPUP, WS_SYSMENU, WS_TABSTOP, BS_DEFPUSHBUTTON, ES_AUTOHSCROLL, ES_AUTOVSCROLL,
    ES_LEFT, ES_MULTILINE, ES_NOHIDESEL, WM_CTLCOLORDLG, WM_CTLCOLORSTATIC,
    GetClientRect, GetWindowRect, MoveWindow, SetTimer,
};

const APP_CLASS: &str = "XpNotepadRustMainWindow";
const APP_NAME: &str = "RetNotepad";

const ID_FILE_NEW: usize = 1001;
const ID_FILE_OPEN: usize = 1002;
const ID_FILE_SAVE: usize = 1003;
const ID_FILE_SAVE_AS: usize = 1004;
const ID_FILE_EXIT: usize = 1005;
const ID_EDIT_UNDO: usize = 2001;
const ID_EDIT_CUT: usize = 2002;
const ID_EDIT_COPY: usize = 2003;
const ID_EDIT_PASTE: usize = 2004;
const ID_EDIT_DELETE: usize = 2005;
const ID_EDIT_SELECT_ALL: usize = 2006;
const ID_EDIT_TIME_DATE: usize = 2007;
const ID_EDIT_FIND: usize = 2008;
const ID_EDIT_FIND_NEXT: usize = 2009;
const ID_EDIT_REPLACE: usize = 2010;
const ID_EDIT_GOTO: usize = 2011;
const ID_HELP_ABOUT: usize = 3001;
const ID_VIEW_STATUS_BAR: usize = 4001;
const ID_FORMAT_WORD_WRAP: usize = 5001;
const ID_FORMAT_FONT: usize = 5002;
const ID_EDIT_CTRL: usize = 1;
const ID_TIMER_STATUS: usize = 1;
const ID_GOTO_EDIT: usize = 7001;

struct Locale {
    file: &'static str,
    edit: &'static str,
    view: &'static str,
    format: &'static str,
    help: &'static str,
    new_file: &'static str,
    open: &'static str,
    save: &'static str,
    save_as: &'static str,
    exit: &'static str,
    undo: &'static str,
    cut: &'static str,
    copy: &'static str,
    paste: &'static str,
    delete: &'static str,
    find: &'static str,
    find_next: &'static str,
    replace: &'static str,
    goto_line: &'static str,
    select_all: &'static str,
    time_date: &'static str,
    status_bar: &'static str,
    word_wrap: &'static str,
    font: &'static str,
    about: &'static str,
    untitled: &'static str,
    app_caption: &'static str,
    open_filter: &'static str,
    save_filter: &'static str,
    save_changes: &'static str,
    load_failed: &'static str,
    save_failed: &'static str,
    io_error_title: &'static str,
    about_title: &'static str,
    version_label: &'static str,
    status_pos: &'static str,
}

const JA: Locale = Locale {
    file: "ファイル(&F)",
    edit: "編集(&E)",
    view: "表示(&V)",
    format: "書式(&O)",
    help: "ヘルプ(&H)",
    new_file: "新規(&N)\tCtrl+N",
    open: "開く(&O)...\tCtrl+O",
    save: "上書き保存(&S)\tCtrl+S",
    save_as: "名前を付けて保存(&A)...",
    exit: "終了(&X)",
    undo: "元に戻す(&U)\tCtrl+Z",
    cut: "切り取り(&T)\tCtrl+X",
    copy: "コピー(&C)\tCtrl+C",
    paste: "貼り付け(&P)\tCtrl+V",
    delete: "削除(&L)\tDel",
    find: "検索(&F)...\tCtrl+F",
    find_next: "次を検索(&N)\tF3",
    replace: "置換(&R)...\tCtrl+H",
    goto_line: "行へ移動(&G)...\tCtrl+G",
    select_all: "すべて選択(&A)\tCtrl+A",
    time_date: "日付と時刻(&D)\tF5",
    status_bar: "ステータス バー(&S)",
    word_wrap: "右端で折り返す(&W)",
    font: "フォント(&F)...",
    about: "バージョン情報(&A)",
    untitled: "無題",
    app_caption: "RetNotepad",
    open_filter: "テキスト ファイル (*.txt)\0*.txt\0すべてのファイル (*.*)\0*.*\0\0",
    save_filter: "テキスト ファイル (*.txt)\0*.txt\0すべてのファイル (*.*)\0*.*\0\0",
    save_changes: "変更を保存しますか？",
    load_failed: "ファイルを開けませんでした。",
    save_failed: "ファイルを保存できませんでした。",
    io_error_title: "I/O エラー",
    about_title: "RetNotepad",
    version_label: "バージョン",
    status_pos: "行 {line}, 列 {col}",
};

const EN: Locale = Locale {
    file: "&File",
    edit: "&Edit",
    view: "&View",
    format: "F&ormat",
    help: "&Help",
    new_file: "&New\tCtrl+N",
    open: "&Open...\tCtrl+O",
    save: "&Save\tCtrl+S",
    save_as: "Save &As...",
    exit: "E&xit",
    undo: "&Undo\tCtrl+Z",
    cut: "Cu&t\tCtrl+X",
    copy: "&Copy\tCtrl+C",
    paste: "&Paste\tCtrl+V",
    delete: "&Delete\tDel",
    find: "&Find...\tCtrl+F",
    find_next: "Find &Next\tF3",
    replace: "&Replace...\tCtrl+H",
    goto_line: "&Go To...\tCtrl+G",
    select_all: "Select &All\tCtrl+A",
    time_date: "Time/&Date\tF5",
    status_bar: "&Status Bar",
    word_wrap: "&Word Wrap",
    font: "&Font...",
    about: "&About",
    untitled: "Untitled",
    app_caption: "RetNotepad",
    open_filter: "Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0",
    save_filter: "Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0\0",
    save_changes: "Do you want to save changes?",
    load_failed: "Failed to open the file.",
    save_failed: "Failed to save the file.",
    io_error_title: "I/O Error",
    about_title: "RetNotepad",
    version_label: "Version",
    status_pos: "Ln {line}, Col {col}",
};

struct AppState {
    edit_hwnd: HWND,
    status_hwnd: HWND,
    show_status_bar: bool,
    word_wrap: bool,
    hfont: HFONT,
    logfont: LOGFONTW,
    current_encoding: TextEncoding,
    current_path: Vec<u16>,
    locale: &'static Locale,
    find_msg: u32,
    find_dialog: Option<Box<FindReplaceState>>,
    last_find: Vec<u16>,
    last_replace: Vec<u16>,
    last_find_flags: u32,
}

#[derive(Clone, Copy)]
struct AppSettings {
    x: i32,
    y: i32,
    width: i32,
    height: i32,
    show_status_bar: bool,
    word_wrap: bool,
    encoding: TextEncoding,
    logfont: LOGFONTW,
    has_logfont: bool,
}

struct AppLaunchData {
    locale: &'static Locale,
    initial_path: Vec<u16>,
    settings: AppSettings,
}

struct FindReplaceState {
    dlg_hwnd: HWND,
    fr: FINDREPLACEW,
    find_buf: [u16; 256],
    replace_buf: [u16; 256],
}

struct GoToDialogState {
    done: bool,
    result_line: i32,
    max_line: i32,
    locale: *const Locale,
    edit_hwnd: HWND,
    font: HFONT,
    bg_brush: HBRUSH,
}

#[derive(Clone, Copy)]
enum TextEncoding {
    Utf8Bom,
    Utf8,
    ShiftJis,
    Utf16Le,
    Utf16Be,
}

const CID_ENCODING_COMBO: u32 = 1001;
const ENC_UTF8_BOM: u32 = 1;
const ENC_UTF8: u32 = 2;
const ENC_SHIFT_JIS: u32 = 3;
const ENC_UTF16_LE: u32 = 4;
const ENC_UTF16_BE: u32 = 5;

const CLSID_FILE_OPEN_DIALOG: GUID = GUID::from_u128(0xdc1c5a9c_e88a_4dde_a5a1_60f82a20aef7);
const CLSID_FILE_SAVE_DIALOG: GUID = GUID::from_u128(0xc0b4e2f3_ba21_4773_8dba_335ec946eb8b);
const IID_IFILE_DIALOG: GUID = GUID::from_u128(0x42f85136_db7e_439c_85f1_e4075d135fc8);
const IID_IFILE_DIALOG_CUSTOMIZE: GUID =
    GUID::from_u128(0xe6fdd21a_163f_4975_9c8c_a69f1ba37034);

#[repr(C)]
struct IUnknownVTbl {
    query_interface: unsafe extern "system" fn(*mut c_void, *const GUID, *mut *mut c_void) -> i32,
    add_ref: unsafe extern "system" fn(*mut c_void) -> u32,
    release: unsafe extern "system" fn(*mut c_void) -> u32,
}

#[repr(C)]
struct IModalWindowVTbl {
    base: IUnknownVTbl,
    show: unsafe extern "system" fn(*mut c_void, HWND) -> i32,
}

#[repr(C)]
struct IFileDialogVTbl {
    base: IModalWindowVTbl,
    set_file_types:
        unsafe extern "system" fn(*mut c_void, u32, *const COMDLG_FILTERSPEC) -> i32,
    set_file_type_index: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    get_file_type_index: unsafe extern "system" fn(*mut c_void, *mut u32) -> i32,
    advise: unsafe extern "system" fn(*mut c_void, *mut c_void, *mut u32) -> i32,
    unadvise: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    set_options: unsafe extern "system" fn(*mut c_void, FILEOPENDIALOGOPTIONS) -> i32,
    get_options: unsafe extern "system" fn(*mut c_void, *mut FILEOPENDIALOGOPTIONS) -> i32,
    set_default_folder: unsafe extern "system" fn(*mut c_void, *mut c_void) -> i32,
    set_folder: unsafe extern "system" fn(*mut c_void, *mut c_void) -> i32,
    get_folder: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32,
    get_current_selection: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32,
    set_file_name: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    get_file_name: unsafe extern "system" fn(*mut c_void, *mut *mut u16) -> i32,
    set_title: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    set_ok_button_label: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    set_file_name_label: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    get_result: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32,
    add_place: unsafe extern "system" fn(*mut c_void, *mut c_void, i32) -> i32,
    set_default_extension: unsafe extern "system" fn(*mut c_void, *const u16) -> i32,
    close: unsafe extern "system" fn(*mut c_void, i32) -> i32,
    set_client_guid: unsafe extern "system" fn(*mut c_void, *const GUID) -> i32,
    clear_client_data: unsafe extern "system" fn(*mut c_void) -> i32,
    set_filter: unsafe extern "system" fn(*mut c_void, *mut c_void) -> i32,
}

#[repr(C)]
struct IFileDialog {
    lp_vtbl: *const IFileDialogVTbl,
}

#[repr(C)]
struct IFileDialogCustomizeVTbl {
    base: IUnknownVTbl,
    enable_open_drop_down: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    add_menu: unsafe extern "system" fn(*mut c_void, u32, *const u16) -> i32,
    add_push_button: unsafe extern "system" fn(*mut c_void, u32, *const u16) -> i32,
    add_combo_box: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    add_radio_button_list: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    add_check_button: unsafe extern "system" fn(*mut c_void, u32, *const u16, i32) -> i32,
    add_edit_box: unsafe extern "system" fn(*mut c_void, u32, *const u16) -> i32,
    add_separator: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    add_text: unsafe extern "system" fn(*mut c_void, u32, *const u16) -> i32,
    set_control_label: unsafe extern "system" fn(*mut c_void, u32, *const u16) -> i32,
    get_control_state: unsafe extern "system" fn(*mut c_void, u32, *mut u32) -> i32,
    set_control_state: unsafe extern "system" fn(*mut c_void, u32, u32) -> i32,
    get_edit_box_text: unsafe extern "system" fn(*mut c_void, u32, *mut *mut u16) -> i32,
    set_edit_box_text: unsafe extern "system" fn(*mut c_void, u32, *const u16) -> i32,
    get_check_button_state: unsafe extern "system" fn(*mut c_void, u32, *mut i32) -> i32,
    set_check_button_state: unsafe extern "system" fn(*mut c_void, u32, i32) -> i32,
    add_control_item: unsafe extern "system" fn(*mut c_void, u32, u32, *const u16) -> i32,
    remove_control_item: unsafe extern "system" fn(*mut c_void, u32, u32) -> i32,
    remove_all_control_items: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    get_control_item_state: unsafe extern "system" fn(*mut c_void, u32, u32, *mut u32) -> i32,
    set_control_item_state: unsafe extern "system" fn(*mut c_void, u32, u32, u32) -> i32,
    get_selected_control_item: unsafe extern "system" fn(*mut c_void, u32, *mut u32) -> i32,
    set_selected_control_item: unsafe extern "system" fn(*mut c_void, u32, u32) -> i32,
    start_visual_group: unsafe extern "system" fn(*mut c_void, u32, *const u16) -> i32,
    end_visual_group: unsafe extern "system" fn(*mut c_void) -> i32,
    make_prominent: unsafe extern "system" fn(*mut c_void, u32) -> i32,
    set_control_item_text: unsafe extern "system" fn(*mut c_void, u32, u32, *const u16) -> i32,
}

#[repr(C)]
struct IFileDialogCustomize {
    lp_vtbl: *const IFileDialogCustomizeVTbl,
}

#[repr(C)]
struct IShellItemVTbl {
    base: IUnknownVTbl,
    bind_to_handler:
        unsafe extern "system" fn(*mut c_void, *mut c_void, *const GUID, *const GUID, *mut *mut c_void) -> i32,
    get_parent: unsafe extern "system" fn(*mut c_void, *mut *mut c_void) -> i32,
    get_display_name: unsafe extern "system" fn(*mut c_void, i32, *mut *mut u16) -> i32,
    get_attributes: unsafe extern "system" fn(*mut c_void, u32, *mut u32) -> i32,
    compare: unsafe extern "system" fn(*mut c_void, *mut c_void, u32, *mut i32) -> i32,
}

#[repr(C)]
struct IShellItem {
    lp_vtbl: *const IShellItemVTbl,
}


fn encoding_to_item_id(enc: TextEncoding) -> u32 {
    match enc {
        TextEncoding::Utf8Bom => ENC_UTF8_BOM,
        TextEncoding::Utf8 => ENC_UTF8,
        TextEncoding::ShiftJis => ENC_SHIFT_JIS,
        TextEncoding::Utf16Le => ENC_UTF16_LE,
        TextEncoding::Utf16Be => ENC_UTF16_BE,
    }
}

fn item_id_to_encoding(id: u32) -> TextEncoding {
    match id {
        ENC_UTF8_BOM => TextEncoding::Utf8Bom,
        ENC_UTF8 => TextEncoding::Utf8,
        ENC_SHIFT_JIS => TextEncoding::ShiftJis,
        ENC_UTF16_LE => TextEncoding::Utf16Le,
        ENC_UTF16_BE => TextEncoding::Utf16Be,
        _ => TextEncoding::Utf8Bom,
    }
}

fn utf16_from_ptr(ptr: *const u16) -> Vec<u16> {
    if ptr.is_null() {
        return Vec::new();
    }
    let mut len = 0usize;
    unsafe {
        while *ptr.add(len) != 0 {
            len += 1;
        }
        let mut out = Vec::with_capacity(len);
        for i in 0..len {
            out.push(*ptr.add(i));
        }
        out
    }
}

fn copy_into_wide_buf(buf: &mut [u16], src: &[u16]) {
    buf.fill(0);
    let len = src.len().min(buf.len().saturating_sub(1));
    if len > 0 {
        buf[..len].copy_from_slice(&src[..len]);
    }
}

fn lower_ascii_u16(ch: u16) -> u16 {
    if (b'A' as u16..=b'Z' as u16).contains(&ch) {
        ch + 32
    } else {
        ch
    }
}

fn is_word_char_u16(ch: u16) -> bool {
    (b'0' as u16..=b'9' as u16).contains(&ch)
        || (b'A' as u16..=b'Z' as u16).contains(&ch)
        || (b'a' as u16..=b'z' as u16).contains(&ch)
        || ch == b'_' as u16
}

fn wide_null(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(Some(0)).collect()
}

fn default_app_launch_data(locale: &'static Locale) -> AppLaunchData {
    AppLaunchData {
        locale,
        initial_path: Vec::new(),
        settings: AppSettings {
            x: CW_USEDEFAULT,
            y: CW_USEDEFAULT,
            width: 900,
            height: 650,
            show_status_bar: true,
            word_wrap: false,
            encoding: TextEncoding::Utf8Bom,
            logfont: unsafe { zeroed() },
            has_logfont: false,
        },
    }
}


unsafe fn multi_to_wide(cp: u32, bytes: &[u8]) -> Vec<u16> {
    if bytes.is_empty() {
        return Vec::new();
    }
    let len = MultiByteToWideChar(cp, 0, bytes.as_ptr(), bytes.len() as i32, null_mut(), 0);
    if len <= 0 {
        return Vec::new();
    }
    let mut out = vec![0u16; len as usize];
    MultiByteToWideChar(
        cp,
        0,
        bytes.as_ptr(),
        bytes.len() as i32,
        out.as_mut_ptr(),
        len,
    );
    out
}

unsafe fn wide_to_multi(cp: u32, text: &[u16]) -> Vec<u8> {
    if text.is_empty() {
        return Vec::new();
    }
    let len = WideCharToMultiByte(
        cp,
        0,
        text.as_ptr(),
        text.len() as i32,
        null_mut(),
        0,
        null(),
        null_mut(),
    );
    if len <= 0 {
        return Vec::new();
    }
    let mut out = vec![0u8; len as usize];
    WideCharToMultiByte(
        cp,
        0,
        text.as_ptr(),
        text.len() as i32,
        out.as_mut_ptr(),
        len,
        null(),
        null_mut(),
    );
    out
}

fn choose_locale() -> &'static Locale {
    unsafe {
        let lang = GetUserDefaultUILanguage();
        if (lang & 0x03ff) as u32 == LANG_JAPANESE {
            &JA
        } else {
            &EN
        }
    }
}

fn loword(v: usize) -> usize {
    v & 0xFFFF
}

fn hiword(v: usize) -> usize {
    (v >> 16) & 0xFFFF
}

unsafe fn get_state(hwnd: HWND) -> Option<&'static mut AppState> {
    let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
    if ptr.is_null() {
        None
    } else {
        Some(&mut *ptr)
    }
}

fn file_name_from_path(path: &[u16]) -> Vec<u16> {
    if path.is_empty() {
        return vec![];
    }
    let mut start = 0usize;
    for (i, c) in path.iter().enumerate() {
        if *c == b'\\' as u16 || *c == b'/' as u16 {
            start = i + 1;
        }
    }
    path[start..].to_vec()
}

unsafe fn update_title(hwnd: HWND, state: &AppState) {
    let is_modified = SendMessageW(state.edit_hwnd, EM_GETMODIFY as u32, 0, 0) != 0;
    let file = if state.current_path.is_empty() {
        wide_null(state.locale.untitled)
    } else {
        let mut f = file_name_from_path(&state.current_path);
        if f.last().copied() != Some(0) {
            f.push(0);
        }
        f
    };
    let app = wide_null(state.locale.app_caption);
    let mut title = Vec::<u16>::with_capacity(file.len() + app.len() + 8);
    if is_modified {
        title.push('*' as u16);
    }
    title.extend_from_slice(&file[..file.len().saturating_sub(1)]);
    title.extend(" - ".encode_utf16());
    title.extend_from_slice(&app[..app.len().saturating_sub(1)]);
    title.push(0);
    SetWindowTextW(hwnd, title.as_ptr());
}

unsafe fn update_status_bar(state: &AppState) {
    let mut sel_start: u32 = 0;
    let mut sel_end: u32 = 0;
    SendMessageW(
        state.edit_hwnd,
        EM_GETSEL as u32,
        &mut sel_start as *mut u32 as usize,
        &mut sel_end as *mut u32 as isize,
    );
    let caret = sel_end as i32;
    let line = SendMessageW(state.edit_hwnd, EM_LINEFROMCHAR as u32, caret as usize, 0) as i32;
    let line_index = SendMessageW(state.edit_hwnd, EM_LINEINDEX as u32, line as usize, 0) as i32;
    let col = (caret - line_index + 1).max(1);
    let text = state
        .locale
        .status_pos
        .replace("{line}", &line.saturating_add(1).to_string())
        .replace("{col}", &col.to_string());
    let wide = wide_null(&text);
    SendMessageW(state.status_hwnd, SB_SETTEXTW, 1, wide.as_ptr() as isize);
}

unsafe fn create_edit_control(hwnd: HWND, word_wrap: bool) -> HWND {
    let edit_class = wide_null("EDIT");
    let mut style = WS_CHILD
        | WS_VISIBLE
        | WS_VSCROLL
        | ES_LEFT as u32
        | ES_MULTILINE as u32
        | ES_AUTOVSCROLL as u32
        | ES_NOHIDESEL as u32;
    if !word_wrap {
        style |= WS_HSCROLL | ES_AUTOHSCROLL as u32;
    }
    CreateWindowExW(
        WS_EX_CLIENTEDGE,
        edit_class.as_ptr(),
        null(),
        style,
        0,
        0,
        0,
        0,
        hwnd,
        ID_EDIT_CTRL as HMENU,
        null_mut(),
        null(),
    )
}

unsafe fn apply_font_to_edit(state: &AppState) {
    if !state.hfont.is_null() {
        SendMessageW(state.edit_hwnd, WM_SETFONT, state.hfont as usize, 1);
    }
}

unsafe fn layout_children(hwnd: HWND, state: &AppState) {
    let mut rc: RECT = zeroed();
    GetClientRect(hwnd, &mut rc);
    let width = (rc.right - rc.left).max(0);
    let height = (rc.bottom - rc.top).max(0);
    let sb_h = if state.show_status_bar { 22 } else { 0 };
    let edit_h = (height - sb_h).max(0);
    MoveWindow(state.edit_hwnd, 0, 0, width, edit_h, 1);
    MoveWindow(state.status_hwnd, 0, edit_h, width, 22, 1);
    if state.show_status_bar {
        let mut parts = [(width - 170).max(0), -1];
        SendMessageW(
            state.status_hwnd,
            SB_SETPARTS,
            parts.len(),
            parts.as_mut_ptr() as isize,
        );
    }
}

unsafe fn open_path_into_editor(
    hwnd: HWND,
    state: &mut AppState,
    path: &[u16],
    enc: TextEncoding,
) {
    match read_file_bytes(path) {
        Ok(bytes) => {
            let text = decode_with_encoding(&bytes, enc);
            set_edit_text(state.edit_hwnd, &text);
            SendMessageW(state.edit_hwnd, EM_SETMODIFY as u32, 0, 0);
            state.current_path = path.to_vec();
            state.current_encoding = enc;
            update_title(hwnd, state);
            update_status_bar(state);
        }
        Err(err) => {
            let detail = format!("{} (code: {})", state.locale.load_failed, err);
            message_box(
                hwnd,
                &detail,
                state.locale.io_error_title,
                MB_OK | MB_ICONERROR,
            );
        }
    }
}

unsafe fn toggle_word_wrap(hwnd: HWND, state: &mut AppState) {
    let text = get_edit_text(state.edit_hwnd);
    let modified = SendMessageW(state.edit_hwnd, EM_GETMODIFY as u32, 0, 0);
    state.word_wrap = !state.word_wrap;
    DestroyWindow(state.edit_hwnd);
    state.edit_hwnd = create_edit_control(hwnd, state.word_wrap);
    apply_font_to_edit(state);
    set_edit_text(state.edit_hwnd, &text);
    SendMessageW(state.edit_hwnd, EM_SETMODIFY as u32, modified as usize, 0);
    SetFocus(state.edit_hwnd);

    if state.word_wrap {
        CheckMenuItem(
            GetMenu(hwnd),
            ID_FORMAT_WORD_WRAP as u32,
            MF_BYCOMMAND | MF_CHECKED,
        );
    } else {
        CheckMenuItem(
            GetMenu(hwnd),
            ID_FORMAT_WORD_WRAP as u32,
            MF_BYCOMMAND | MF_UNCHECKED,
        );
    }
    layout_children(hwnd, state);
    update_title(hwnd, state);
    update_status_bar(state);
}

unsafe fn choose_display_font(hwnd: HWND, state: &mut AppState) {
    let mut lf = state.logfont;
    let mut cf: CHOOSEFONTW = zeroed();
    cf.lStructSize = size_of::<CHOOSEFONTW>() as u32;
    cf.hwndOwner = hwnd;
    cf.lpLogFont = &mut lf;
    cf.Flags = CF_SCREENFONTS | CF_INITTOLOGFONTSTRUCT;
    if ChooseFontW(&mut cf) == 0 {
        return;
    }
    let new_font = CreateFontIndirectW(&lf);
    if new_font.is_null() {
        return;
    }
    if !state.hfont.is_null() {
        DeleteObject(state.hfont as _);
    }
    state.hfont = new_font;
    state.logfont = lf;
    apply_font_to_edit(state);
}

unsafe fn insert_date_time(state: &mut AppState) {
    let mut st: SYSTEMTIME = zeroed();
    GetLocalTime(&mut st);
    let mut tbuf = [0u16; 128];
    let mut dbuf = [0u16; 128];
    let tlen = GetTimeFormatW(
        LOCALE_USER_DEFAULT,
        0,
        &st,
        null(),
        tbuf.as_mut_ptr(),
        tbuf.len() as i32,
    );
    let dlen = GetDateFormatW(
        LOCALE_USER_DEFAULT,
        0,
        &st,
        null(),
        dbuf.as_mut_ptr(),
        dbuf.len() as i32,
    );
    if tlen <= 1 || dlen <= 1 {
        return;
    }
    let mut ins = Vec::<u16>::new();
    ins.extend_from_slice(&tbuf[..(tlen as usize - 1)]);
    ins.push(' ' as u16);
    ins.extend_from_slice(&dbuf[..(dlen as usize - 1)]);
    ins.push(0);
    SendMessageW(state.edit_hwnd, EM_REPLACESEL as u32, 1, ins.as_ptr() as isize);
}

unsafe fn message_box(hwnd: HWND, text: &str, title: &str, flags: u32) -> i32 {
    let t = wide_null(text);
    let c = wide_null(title);
    MessageBoxW(hwnd, t.as_ptr(), c.as_ptr(), flags)
}

unsafe fn read_file_bytes(path: &[u16]) -> Result<Vec<u8>, u32> {
    let handle = CreateFileW(
        path.as_ptr(),
        GENERIC_READ,
        FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
        null(),
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        null_mut(),
    );
    if handle == INVALID_HANDLE_VALUE {
        return Err(GetLastError());
    }
    let mut buffer = Vec::<u8>::new();
    let mut chunk = vec![0u8; 64 * 1024];
    loop {
        let mut read = 0u32;
        if ReadFile(
            handle,
            chunk.as_mut_ptr(),
            chunk.len() as u32,
            &mut read,
            null_mut(),
        ) == 0
        {
            let err = GetLastError();
            CloseHandle(handle);
            return Err(err);
        }
        if read == 0 {
            break;
        }
        buffer.extend_from_slice(&chunk[..read as usize]);
    }
    CloseHandle(handle);
    Ok(buffer)
}

unsafe fn write_file_bytes(path: &[u16], data: &[u8]) -> Result<(), u32> {
    let handle = CreateFileW(
        path.as_ptr(),
        GENERIC_WRITE,
        0,
        null(),
        CREATE_ALWAYS,
        FILE_ATTRIBUTE_NORMAL,
        null_mut(),
    );
    if handle == INVALID_HANDLE_VALUE {
        return Err(GetLastError());
    }
    let mut total = 0usize;
    while total < data.len() {
        let mut written = 0u32;
        let remain = (data.len() - total).min(u32::MAX as usize);
        if WriteFile(
            handle,
            data[total..].as_ptr() as *const _,
            remain as u32,
            &mut written,
            null_mut(),
        ) == 0
        {
            let err = GetLastError();
            CloseHandle(handle);
            return Err(err);
        }
        if written == 0 {
            break;
        }
        total += written as usize;
    }
    CloseHandle(handle);
    Ok(())
}

unsafe fn decode_with_encoding(bytes: &[u8], enc: TextEncoding) -> Vec<u16> {
    match enc {
        TextEncoding::Utf8Bom => {
            let src = if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
                &bytes[3..]
            } else {
                bytes
            };
            multi_to_wide(CP_UTF8, src)
        }
        TextEncoding::Utf8 => {
            let src = if bytes.starts_with(&[0xEF, 0xBB, 0xBF]) {
                &bytes[3..]
            } else {
                bytes
            };
            multi_to_wide(CP_UTF8, src)
        }
        TextEncoding::ShiftJis => multi_to_wide(CP_ACP, bytes),
        TextEncoding::Utf16Le => {
            let src = if bytes.starts_with(&[0xFF, 0xFE]) {
                &bytes[2..]
            } else {
                bytes
            };
            let mut out = Vec::new();
            let mut i = 0usize;
            while i + 1 < src.len() {
                out.push(u16::from_le_bytes([src[i], src[i + 1]]));
                i += 2;
            }
            out
        }
        TextEncoding::Utf16Be => {
            let src = if bytes.starts_with(&[0xFE, 0xFF]) {
                &bytes[2..]
            } else {
                bytes
            };
            let mut out = Vec::new();
            let mut i = 0usize;
            while i + 1 < src.len() {
                out.push(u16::from_be_bytes([src[i], src[i + 1]]));
                i += 2;
            }
            out
        }
    }
}

unsafe fn encode_with_encoding(text: &[u16], enc: TextEncoding) -> Vec<u8> {
    match enc {
        TextEncoding::Utf8Bom => {
            let mut out = vec![0xEF, 0xBB, 0xBF];
            out.extend_from_slice(&wide_to_multi(CP_UTF8, text));
            out
        }
        TextEncoding::Utf8 => wide_to_multi(CP_UTF8, text),
        TextEncoding::ShiftJis => wide_to_multi(CP_ACP, text),
        TextEncoding::Utf16Le => {
            let mut out = Vec::with_capacity(text.len() * 2);
            for ch in text {
                let b = ch.to_le_bytes();
                out.push(b[0]);
                out.push(b[1]);
            }
            out
        }
        TextEncoding::Utf16Be => {
            let mut out = Vec::with_capacity(text.len() * 2);
            for ch in text {
                let b = ch.to_be_bytes();
                out.push(b[0]);
                out.push(b[1]);
            }
            out
        }
    }
}

unsafe fn succeeded(hr: i32) -> bool {
    hr >= 0
}

unsafe fn pwstr_to_wide_owned(ptr: *mut u16) -> Vec<u16> {
    if ptr.is_null() {
        return Vec::new();
    }
    let mut len = 0usize;
    while *ptr.add(len) != 0 {
        len += 1;
    }
    let mut v = Vec::with_capacity(len + 1);
    for i in 0..len {
        v.push(*ptr.add(i));
    }
    v.push(0);
    CoTaskMemFree(ptr as *const c_void);
    v
}

unsafe fn get_dialog_result_path(dialog: *mut IFileDialog) -> Option<Vec<u16>> {
    let mut item_ptr: *mut c_void = null_mut();
    let hr = ((*(*dialog).lp_vtbl).get_result)(dialog as *mut c_void, &mut item_ptr);
    if !succeeded(hr) || item_ptr.is_null() {
        return None;
    }
    let item = item_ptr as *mut IShellItem;
    let mut pwsz: *mut u16 = null_mut();
    let hr_path = ((*(*item).lp_vtbl).get_display_name)(
        item as *mut c_void,
        SIGDN_FILESYSPATH,
        &mut pwsz,
    );
    ((*(*item).lp_vtbl).base.release)(item as *mut c_void);
    if !succeeded(hr_path) {
        return None;
    }
    Some(pwstr_to_wide_owned(pwsz))
}

unsafe fn configure_encoding_combo(
    dialog: *mut IFileDialog,
    locale: &Locale,
    initial_enc: TextEncoding,
) -> Option<*mut IFileDialogCustomize> {
    let mut customize_ptr: *mut c_void = null_mut();
    let hr_qi = ((*(*dialog).lp_vtbl).base.base.query_interface)(
        dialog as *mut c_void,
        &IID_IFILE_DIALOG_CUSTOMIZE,
        &mut customize_ptr,
    );
    if !succeeded(hr_qi) || customize_ptr.is_null() {
        return None;
    }
    let customize = customize_ptr as *mut IFileDialogCustomize;
    let enc_label = if std::ptr::eq(locale as *const Locale, &JA as *const Locale) {
        wide_null("エンコード(&E):")
    } else {
        wide_null("Encoding(&E):")
    };
    let s1 = wide_null("UTF-8 (BOM)");
    let s2 = wide_null("UTF-8");
    let s3 = wide_null("Shift-JIS (ASCII)");
    let s4 = wide_null("UTF-16 LE");
    let s5 = wide_null("UTF-16 BE");

    ((*(*customize).lp_vtbl).add_combo_box)(customize as *mut c_void, CID_ENCODING_COMBO);
    ((*(*customize).lp_vtbl).set_control_label)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        enc_label.as_ptr(),
    );
    ((*(*customize).lp_vtbl).add_control_item)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        ENC_UTF8_BOM,
        s1.as_ptr(),
    );
    ((*(*customize).lp_vtbl).add_control_item)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        ENC_UTF8,
        s2.as_ptr(),
    );
    ((*(*customize).lp_vtbl).add_control_item)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        ENC_SHIFT_JIS,
        s3.as_ptr(),
    );
    ((*(*customize).lp_vtbl).add_control_item)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        ENC_UTF16_LE,
        s4.as_ptr(),
    );
    ((*(*customize).lp_vtbl).add_control_item)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        ENC_UTF16_BE,
        s5.as_ptr(),
    );
    ((*(*customize).lp_vtbl).set_selected_control_item)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        encoding_to_item_id(initial_enc),
    );
    Some(customize)
}

unsafe fn get_selected_encoding(customize: *mut IFileDialogCustomize) -> TextEncoding {
    let mut id = ENC_UTF8_BOM;
    let hr = ((*(*customize).lp_vtbl).get_selected_control_item)(
        customize as *mut c_void,
        CID_ENCODING_COMBO,
        &mut id,
    );
    if succeeded(hr) {
        item_id_to_encoding(id)
    } else {
        TextEncoding::Utf8Bom
    }
}

unsafe fn run_common_dialog_with_encoding(
    hwnd: HWND,
    locale: &Locale,
    initial_path: &[u16],
    initial_enc: TextEncoding,
    for_save: bool,
) -> Option<(Vec<u16>, TextEncoding)> {
    let clsid = if for_save {
        &CLSID_FILE_SAVE_DIALOG
    } else {
        &CLSID_FILE_OPEN_DIALOG
    };
    let mut dialog_ptr: *mut c_void = null_mut();
    let hr = CoCreateInstance(
        clsid,
        null_mut(),
        CLSCTX_INPROC_SERVER,
        &IID_IFILE_DIALOG,
        &mut dialog_ptr,
    );
    if !succeeded(hr) || dialog_ptr.is_null() {
        return None;
    }
    let dialog = dialog_ptr as *mut IFileDialog;

    let mut options: FILEOPENDIALOGOPTIONS = 0;
    ((*(*dialog).lp_vtbl).get_options)(dialog as *mut c_void, &mut options);
    options |= FOS_FORCEFILESYSTEM;
    options |= if for_save {
        FOS_OVERWRITEPROMPT
    } else {
        FOS_FILEMUSTEXIST
    };
    ((*(*dialog).lp_vtbl).set_options)(dialog as *mut c_void, options);

    let t1_name = wide_null("Text Documents (*.txt)");
    let t1_spec = wide_null("*.txt");
    let t2_name = wide_null("All Files (*.*)");
    let t2_spec = wide_null("*.*");
    let filters = [
        COMDLG_FILTERSPEC {
            pszName: t1_name.as_ptr(),
            pszSpec: t1_spec.as_ptr(),
        },
        COMDLG_FILTERSPEC {
            pszName: t2_name.as_ptr(),
            pszSpec: t2_spec.as_ptr(),
        },
    ];
    ((*(*dialog).lp_vtbl).set_file_types)(dialog as *mut c_void, filters.len() as u32, filters.as_ptr());
    ((*(*dialog).lp_vtbl).set_file_type_index)(dialog as *mut c_void, 1);

    let title = if for_save {
        wide_null(locale.save_as)
    } else {
        wide_null(locale.open)
    };
    ((*(*dialog).lp_vtbl).set_title)(dialog as *mut c_void, title.as_ptr());

    if for_save && !initial_path.is_empty() {
        ((*(*dialog).lp_vtbl).set_file_name)(dialog as *mut c_void, initial_path.as_ptr());
    }

    let customize = configure_encoding_combo(dialog, locale, initial_enc);
    let hr_show = ((*(*dialog).lp_vtbl).base.show)(dialog as *mut c_void, hwnd);
    if !succeeded(hr_show) {
        if let Some(c) = customize {
            ((*(*c).lp_vtbl).base.release)(c as *mut c_void);
        }
        ((*(*dialog).lp_vtbl).base.base.release)(dialog as *mut c_void);
        return None;
    }

    let enc = if let Some(c) = customize {
        let e = get_selected_encoding(c);
        ((*(*c).lp_vtbl).base.release)(c as *mut c_void);
        e
    } else {
        initial_enc
    };
    let path = get_dialog_result_path(dialog);
    ((*(*dialog).lp_vtbl).base.base.release)(dialog as *mut c_void);
    path.map(|p| (p, enc))
}

unsafe fn run_open_dialog_with_encoding(
    hwnd: HWND,
    locale: &Locale,
    initial_enc: TextEncoding,
) -> Option<(Vec<u16>, TextEncoding)> {
    run_common_dialog_with_encoding(hwnd, locale, &[], initial_enc, false)
}

unsafe fn run_save_dialog_with_encoding(
    hwnd: HWND,
    locale: &Locale,
    initial_path: &[u16],
    initial_enc: TextEncoding,
) -> Option<(Vec<u16>, TextEncoding)> {
    run_common_dialog_with_encoding(hwnd, locale, initial_path, initial_enc, true)
}

fn find_range(
    text: &[u16],
    pat: &[u16],
    start: usize,
    down: bool,
    match_case: bool,
    whole_word: bool,
) -> Option<usize> {
    if pat.is_empty() || text.is_empty() || pat.len() > text.len() {
        return None;
    }
    let end = text.len() - pat.len();
    let matched_at = |i: usize| -> bool {
        for j in 0..pat.len() {
            let a = text[i + j];
            let b = pat[j];
            let eq = if match_case {
                a == b
            } else {
                lower_ascii_u16(a) == lower_ascii_u16(b)
            };
            if !eq {
                return false;
            }
        }
        if whole_word {
            let left_ok = i == 0 || !is_word_char_u16(text[i - 1]);
            let right_idx = i + pat.len();
            let right_ok = right_idx >= text.len() || !is_word_char_u16(text[right_idx]);
            left_ok && right_ok
        } else {
            true
        }
    };

    if down {
        for i in start..=end {
            if matched_at(i) {
                return Some(i);
            }
        }
    } else {
        let mut i = start.min(end);
        loop {
            if matched_at(i) {
                return Some(i);
            }
            if i == 0 {
                break;
            }
            i -= 1;
        }
    }
    None
}

unsafe fn select_range(state: &mut AppState, start: usize, end: usize) {
    SendMessageW(state.edit_hwnd, EM_SETSEL as u32, start, end as isize);
    SetFocus(state.edit_hwnd);
    update_status_bar(state);
}

unsafe fn do_find_next_with_flags(state: &mut AppState, pattern: &[u16], flags: u32) -> bool {
    if pattern.is_empty() {
        return false;
    }
    let text = get_edit_text(state.edit_hwnd);
    let mut sel_start: u32 = 0;
    let mut sel_end: u32 = 0;
    SendMessageW(
        state.edit_hwnd,
        EM_GETSEL as u32,
        &mut sel_start as *mut u32 as usize,
        &mut sel_end as *mut u32 as isize,
    );
    let down = (flags & FR_DOWN) != 0;
    let start = if down {
        sel_end as usize
    } else {
        sel_start.saturating_sub(1) as usize
    };
    let match_case = (flags & FR_MATCHCASE) != 0;
    let whole_word = (flags & FR_WHOLEWORD) != 0;
    if let Some(idx) = find_range(&text, pattern, start, down, match_case, whole_word) {
        select_range(state, idx, idx + pattern.len());
        true
    } else {
        false
    }
}

unsafe fn create_find_replace_dialog(hwnd: HWND, state: &mut AppState, replace_mode: bool) {
    if state.find_dialog.is_some() {
        return;
    }
    let mut fr_state = Box::new(FindReplaceState {
        dlg_hwnd: null_mut(),
        fr: zeroed(),
        find_buf: [0; 256],
        replace_buf: [0; 256],
    });
    copy_into_wide_buf(&mut fr_state.find_buf, &state.last_find);
    copy_into_wide_buf(&mut fr_state.replace_buf, &state.last_replace);
    fr_state.fr.lStructSize = size_of::<FINDREPLACEW>() as u32;
    fr_state.fr.hwndOwner = hwnd;
    fr_state.fr.Flags = FR_DOWN | (state.last_find_flags & (FR_MATCHCASE | FR_WHOLEWORD));
    fr_state.fr.lpstrFindWhat = fr_state.find_buf.as_mut_ptr();
    fr_state.fr.wFindWhatLen = fr_state.find_buf.len() as u16;
    fr_state.fr.lpstrReplaceWith = fr_state.replace_buf.as_mut_ptr();
    fr_state.fr.wReplaceWithLen = fr_state.replace_buf.len() as u16;
    let dlg = if replace_mode {
        ReplaceTextW(&mut fr_state.fr)
    } else {
        FindTextW(&mut fr_state.fr)
    };
    if !dlg.is_null() {
        fr_state.dlg_hwnd = dlg;
        state.find_dialog = Some(fr_state);
    }
}

unsafe fn handle_find_replace_notification(hwnd: HWND, state: &mut AppState, lparam: LPARAM) {
    let fr = lparam as *mut FINDREPLACEW;
    if fr.is_null() {
        return;
    }
    if ((*fr).Flags & FR_DIALOGTERM) != 0 {
        state.find_dialog = None;
        return;
    }
    let flags = (*fr).Flags;
    let find_text = utf16_from_ptr((*fr).lpstrFindWhat);
    let repl_text = utf16_from_ptr((*fr).lpstrReplaceWith);
    state.last_find = find_text.clone();
    state.last_replace = repl_text.clone();
    state.last_find_flags = flags & (FR_DOWN | FR_MATCHCASE | FR_WHOLEWORD);

    if (flags & FR_FINDNEXT) != 0 {
        if !do_find_next_with_flags(state, &find_text, state.last_find_flags | FR_DOWN) {
            let msg = if std::ptr::eq(state.locale as *const Locale, &JA as *const Locale) {
                "見つかりませんでした。"
            } else {
                "Text not found."
            };
            message_box(hwnd, msg, state.locale.app_caption, MB_OK);
        }
        return;
    }

    if (flags & FR_REPLACE) != 0 {
        let mut sel_start: u32 = 0;
        let mut sel_end: u32 = 0;
        SendMessageW(
            state.edit_hwnd,
            EM_GETSEL as u32,
            &mut sel_start as *mut u32 as usize,
            &mut sel_end as *mut u32 as isize,
        );
        let text = get_edit_text(state.edit_hwnd);
        let s = sel_start as usize;
        let e = sel_end as usize;
        let selected = if s < e && e <= text.len() { &text[s..e] } else { &[] };
        let match_case = (state.last_find_flags & FR_MATCHCASE) != 0;
        let mut equal = selected.len() == find_text.len();
        if equal {
            for (a, b) in selected.iter().zip(find_text.iter()) {
                let eq = if match_case {
                    *a == *b
                } else {
                    lower_ascii_u16(*a) == lower_ascii_u16(*b)
                };
                if !eq {
                    equal = false;
                    break;
                }
            }
        }
        if equal {
            let mut repl = repl_text.clone();
            repl.push(0);
            SendMessageW(state.edit_hwnd, EM_REPLACESEL as u32, 1, repl.as_ptr() as isize);
            update_title(hwnd, state);
        }
        let _ = do_find_next_with_flags(state, &find_text, state.last_find_flags | FR_DOWN);
        return;
    }

    if (flags & FR_REPLACEALL) != 0 {
        SendMessageW(state.edit_hwnd, EM_SETSEL as u32, 0, 0);
        let mut guard = 0usize;
        while guard < 200000 {
            if !do_find_next_with_flags(state, &find_text, FR_DOWN | state.last_find_flags) {
                break;
            }
            let mut repl = repl_text.clone();
            repl.push(0);
            SendMessageW(state.edit_hwnd, EM_REPLACESEL as u32, 1, repl.as_ptr() as isize);
            guard += 1;
        }
        update_title(hwnd, state);
    }
}

unsafe extern "system" fn goto_wnd_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    match msg {
        WM_CREATE => {
            let cs = lparam as *const CREATESTRUCTW;
            if cs.is_null() {
                return 0;
            }
            let st = (*cs).lpCreateParams as *mut GoToDialogState;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, st as isize);
            let txt = if std::ptr::eq((*st).locale, &JA as *const Locale) {
                wide_null("行番号:")
            } else {
                wide_null("Line number:")
            };
            let cls_static = wide_null("STATIC");
            let cls_edit = wide_null("EDIT");
            let cls_btn = wide_null("BUTTON");
            let ok = wide_null("OK");
            let cancel = if std::ptr::eq((*st).locale, &JA as *const Locale) {
                wide_null("キャンセル")
            } else {
                wide_null("Cancel")
            };
            let label = CreateWindowExW(
                0,
                cls_static.as_ptr(),
                txt.as_ptr(),
                WS_CHILD | WS_VISIBLE,
                12,
                12,
                80,
                20,
                hwnd,
                null_mut(),
                null_mut(),
                null(),
            );
            let edit = CreateWindowExW(
                WS_EX_CLIENTEDGE,
                cls_edit.as_ptr(),
                wide_null("1").as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                92,
                10,
                120,
                24,
                hwnd,
                ID_GOTO_EDIT as HMENU,
                null_mut(),
                null(),
            );
            let ok_btn = CreateWindowExW(
                0,
                cls_btn.as_ptr(),
                ok.as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON as u32,
                56,
                46,
                70,
                24,
                hwnd,
                IDOK as HMENU,
                null_mut(),
                null(),
            );
            let cancel_btn = CreateWindowExW(
                0,
                cls_btn.as_ptr(),
                cancel.as_ptr(),
                WS_CHILD | WS_VISIBLE | WS_TABSTOP,
                138,
                46,
                74,
                24,
                hwnd,
                IDCANCEL as HMENU,
                null_mut(),
                null(),
            );
            (*st).edit_hwnd = edit;
            if !(*st).font.is_null() {
                SendMessageW(hwnd, WM_SETFONT, (*st).font as usize, 1);
                SendMessageW(label, WM_SETFONT, (*st).font as usize, 1);
                SendMessageW(edit, WM_SETFONT, (*st).font as usize, 1);
                SendMessageW(ok_btn, WM_SETFONT, (*st).font as usize, 1);
                SendMessageW(cancel_btn, WM_SETFONT, (*st).font as usize, 1);
            }
            SetFocus(edit);
            0
        }
        WM_COMMAND => {
            let st = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut GoToDialogState;
            if st.is_null() {
                return 0;
            }
            match loword(wparam) as i32 {
                IDOK => {
                    let mut buf = [0u16; 32];
                    let len = GetWindowTextW((*st).edit_hwnd, buf.as_mut_ptr(), buf.len() as i32);
                    let s = String::from_utf16_lossy(&buf[..len.max(0) as usize]);
                    if let Ok(v) = s.trim().parse::<i32>() {
                        if v >= 1 && v <= (*st).max_line {
                            (*st).result_line = v;
                            (*st).done = true;
                            DestroyWindow(hwnd);
                            return 0;
                        }
                    }
                    let msg = if std::ptr::eq((*st).locale, &JA as *const Locale) {
                        "有効な行番号を入力してください。"
                    } else {
                        "Enter a valid line number."
                    };
                    MessageBoxW(hwnd, wide_null(msg).as_ptr(), wide_null("Go To").as_ptr(), MB_OK | MB_ICONWARNING);
                }
                IDCANCEL => {
                    (*st).done = true;
                    DestroyWindow(hwnd);
                }
                _ => {}
            }
            0
        }
        WM_CTLCOLORDLG | WM_CTLCOLORSTATIC => {
            let st = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut GoToDialogState;
            if !st.is_null() {
                return (*st).bg_brush as isize;
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        WM_CLOSE => {
            let st = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut GoToDialogState;
            if !st.is_null() {
                (*st).done = true;
            }
            DestroyWindow(hwnd);
            0
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn show_goto_line_dialog(
    owner: HWND,
    locale: &Locale,
    max_line: i32,
    font: HFONT,
    bg_brush: HBRUSH,
) -> Option<i32> {
    let cls = wide_null("XpNotepadGotoLineDialog");
    let wnd = WNDCLASSW {
        lpfnWndProc: Some(goto_wnd_proc),
        hInstance: GetModuleHandleW(null()),
        hCursor: LoadCursorW(null_mut(), IDC_ARROW),
        hbrBackground: bg_brush,
        lpszClassName: cls.as_ptr(),
        ..zeroed()
    };
    RegisterClassW(&wnd);

    let mut st = Box::new(GoToDialogState {
        done: false,
        result_line: -1,
        max_line,
        locale: locale as *const Locale,
        edit_hwnd: null_mut(),
        font,
        bg_brush,
    });
    let st_ptr: *mut GoToDialogState = &mut *st;
    let title = if std::ptr::eq(locale as *const Locale, &JA as *const Locale) {
        wide_null("行へ移動")
    } else {
        wide_null("Go To Line")
    };
    let dlg_w = 240;
    let dlg_h = 110;
    let mut x = CW_USEDEFAULT;
    let mut y = CW_USEDEFAULT;
    if !owner.is_null() {
        let mut rc: RECT = zeroed();
        if GetWindowRect(owner, &mut rc) != 0 {
            let ow = rc.right - rc.left;
            let oh = rc.bottom - rc.top;
            x = rc.left + ((ow - dlg_w) / 2).max(0);
            y = rc.top + ((oh - dlg_h) / 2).max(0);
        }
    }
    let hwnd = CreateWindowExW(
        0,
        cls.as_ptr(),
        title.as_ptr(),
        WS_POPUP | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        x,
        y,
        dlg_w,
        dlg_h,
        owner,
        null_mut(),
        GetModuleHandleW(null()),
        st_ptr as *const c_void,
    );
    if hwnd.is_null() {
        return None;
    }
    let mut msg: MSG = zeroed();
    while !st.done && GetMessageW(&mut msg, null_mut(), 0, 0) > 0 {
        if IsDialogMessageW(hwnd, &msg) == 0 {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }
    let result = st.result_line;
    if result > 0 {
        Some(result)
    } else {
        None
    }
}

unsafe fn set_edit_text(edit: HWND, text: &[u16]) {
    let mut buf = text.to_vec();
    if buf.last().copied() != Some(0) {
        buf.push(0);
    }
    SetWindowTextW(edit, buf.as_ptr());
}

unsafe fn get_edit_text(edit: HWND) -> Vec<u16> {
    let len = GetWindowTextLengthW(edit);
    if len <= 0 {
        return Vec::new();
    }
    let mut buf = vec![0u16; len as usize + 1];
    GetWindowTextW(edit, buf.as_mut_ptr(), len + 1);
    buf.truncate(len as usize);
    buf
}

unsafe fn do_open_file(hwnd: HWND, state: &mut AppState) {
    if let Some((path, enc)) = run_open_dialog_with_encoding(hwnd, state.locale, state.current_encoding)
    {
        open_path_into_editor(hwnd, state, &path, enc);
    }
}

unsafe fn do_save_to_path(
    hwnd: HWND,
    state: &mut AppState,
    path: &[u16],
    enc: TextEncoding,
) -> bool {
    let text = get_edit_text(state.edit_hwnd);
    let bytes = encode_with_encoding(&text, enc);
    match write_file_bytes(path, &bytes) {
        Ok(_) => {
            SendMessageW(state.edit_hwnd, EM_SETMODIFY as u32, 0, 0);
            state.current_path = path.to_vec();
            state.current_encoding = enc;
            update_title(hwnd, state);
            true
        }
        Err(_) => {
            message_box(
                hwnd,
                state.locale.save_failed,
                state.locale.io_error_title,
                MB_OK | MB_ICONERROR,
            );
            false
        }
    }
}

unsafe fn do_save_as_with_encoding(hwnd: HWND, state: &mut AppState, enc: TextEncoding) -> bool {
    let Some((path, enc_from_dlg)) = run_save_dialog_with_encoding(
        hwnd,
        state.locale,
        &state.current_path,
        enc,
    ) else {
        return false;
    };
    do_save_to_path(hwnd, state, &path, enc_from_dlg)
}

unsafe fn do_save_as(hwnd: HWND, state: &mut AppState) -> bool {
    do_save_as_with_encoding(hwnd, state, state.current_encoding)
}

unsafe fn do_save(hwnd: HWND, state: &mut AppState) -> bool {
    do_save_as_with_encoding(hwnd, state, state.current_encoding)
}

unsafe fn confirm_save_if_needed(hwnd: HWND, state: &mut AppState) -> bool {
    let modified = SendMessageW(state.edit_hwnd, EM_GETMODIFY as u32, 0, 0) != 0;
    if !modified {
        return true;
    }
    let text = wide_null(state.locale.save_changes);
    let caption = wide_null(state.locale.app_caption);
    let answer = MessageBoxW(
        hwnd,
        text.as_ptr(),
        caption.as_ptr(),
        MB_YESNOCANCEL | MB_ICONWARNING,
    );
    match answer {
        6 => do_save(hwnd, state), // IDYES
        7 => true,                 // IDNO
        _ => false,                // IDCANCEL or error
    }
}

unsafe fn build_menu(locale: &Locale) -> HMENU {
    let main = CreateMenu();
    let file_menu = CreatePopupMenu();
    let edit_menu = CreatePopupMenu();
    let view_menu = CreatePopupMenu();
    let format_menu = CreatePopupMenu();
    let help_menu = CreatePopupMenu();

    let m_file = wide_null(locale.file);
    let m_edit = wide_null(locale.edit);
    let m_view = wide_null(locale.view);
    let m_format = wide_null(locale.format);
    let m_help = wide_null(locale.help);
    let i_new = wide_null(locale.new_file);
    let i_open = wide_null(locale.open);
    let i_save = wide_null(locale.save);
    let i_save_as = wide_null(locale.save_as);
    let i_exit = wide_null(locale.exit);
    let i_undo = wide_null(locale.undo);
    let i_cut = wide_null(locale.cut);
    let i_copy = wide_null(locale.copy);
    let i_paste = wide_null(locale.paste);
    let i_delete = wide_null(locale.delete);
    let i_find = wide_null(locale.find);
    let i_find_next = wide_null(locale.find_next);
    let i_replace = wide_null(locale.replace);
    let i_goto = wide_null(locale.goto_line);
    let i_select_all = wide_null(locale.select_all);
    let i_time_date = wide_null(locale.time_date);
    let i_status_bar = wide_null(locale.status_bar);
    let i_word_wrap = wide_null(locale.word_wrap);
    let i_font = wide_null(locale.font);
    let i_about = wide_null(locale.about);

    AppendMenuW(file_menu, MF_STRING, ID_FILE_NEW, i_new.as_ptr());
    AppendMenuW(file_menu, MF_STRING, ID_FILE_OPEN, i_open.as_ptr());
    AppendMenuW(file_menu, MF_STRING, ID_FILE_SAVE, i_save.as_ptr());
    AppendMenuW(file_menu, MF_STRING, ID_FILE_SAVE_AS, i_save_as.as_ptr());
    AppendMenuW(file_menu, MF_SEPARATOR, 0, null());
    AppendMenuW(file_menu, MF_STRING, ID_FILE_EXIT, i_exit.as_ptr());

    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_UNDO, i_undo.as_ptr());
    AppendMenuW(edit_menu, MF_SEPARATOR, 0, null());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_CUT, i_cut.as_ptr());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_COPY, i_copy.as_ptr());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_PASTE, i_paste.as_ptr());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_DELETE, i_delete.as_ptr());
    AppendMenuW(edit_menu, MF_SEPARATOR, 0, null());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_FIND, i_find.as_ptr());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_FIND_NEXT, i_find_next.as_ptr());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_REPLACE, i_replace.as_ptr());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_GOTO, i_goto.as_ptr());
    AppendMenuW(edit_menu, MF_SEPARATOR, 0, null());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_SELECT_ALL, i_select_all.as_ptr());
    AppendMenuW(edit_menu, MF_STRING, ID_EDIT_TIME_DATE, i_time_date.as_ptr());

    AppendMenuW(view_menu, MF_STRING, ID_VIEW_STATUS_BAR, i_status_bar.as_ptr());
    AppendMenuW(format_menu, MF_STRING, ID_FORMAT_WORD_WRAP, i_word_wrap.as_ptr());
    AppendMenuW(format_menu, MF_STRING, ID_FORMAT_FONT, i_font.as_ptr());

    AppendMenuW(help_menu, MF_STRING, ID_HELP_ABOUT, i_about.as_ptr());

    AppendMenuW(main, MF_POPUP, file_menu as usize, m_file.as_ptr());
    AppendMenuW(main, MF_POPUP, edit_menu as usize, m_edit.as_ptr());
    AppendMenuW(main, MF_POPUP, format_menu as usize, m_format.as_ptr());
    AppendMenuW(main, MF_POPUP, view_menu as usize, m_view.as_ptr());
    AppendMenuW(main, MF_POPUP, help_menu as usize, m_help.as_ptr());
    main
}

unsafe extern "system" fn wnd_proc(
    hwnd: HWND,
    msg: u32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if let Some(st) = get_state(hwnd) {
        if msg == st.find_msg {
            handle_find_replace_notification(hwnd, st, lparam);
            return 0;
        }
    }
    match msg {
        WM_CREATE => {
            let cs = lparam as *const CREATESTRUCTW;
            let launch = if cs.is_null() || (*cs).lpCreateParams.is_null() {
                Box::new(default_app_launch_data(&JA))
            } else {
                Box::from_raw((*cs).lpCreateParams as *mut AppLaunchData)
            };
            let locale = launch.locale;
            let edit = create_edit_control(hwnd, launch.settings.word_wrap);
            let status = CreateWindowExW(
                0,
                STATUSCLASSNAMEW,
                null(),
                WS_CHILD | WS_VISIBLE,
                0,
                0,
                0,
                0,
                hwnd,
                null_mut(),
                null_mut(),
                null(),
            );
            let state = Box::new(AppState {
                edit_hwnd: edit,
                status_hwnd: status,
                show_status_bar: launch.settings.show_status_bar,
                word_wrap: launch.settings.word_wrap,
                hfont: null_mut(),
                logfont: launch.settings.logfont,
                current_encoding: launch.settings.encoding,
                current_path: launch.initial_path.clone(),
                locale,
                find_msg: RegisterWindowMessageW(FINDMSGSTRINGW),
                find_dialog: None,
                last_find: Vec::new(),
                last_replace: Vec::new(),
                last_find_flags: FR_DOWN,
            });
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, Box::into_raw(state) as isize);
            DragAcceptFiles(hwnd, 1);
            if let Some(st) = get_state(hwnd) {
                if launch.settings.has_logfont {
                    let font = CreateFontIndirectW(&launch.settings.logfont);
                    if !font.is_null() {
                        st.hfont = font;
                        apply_font_to_edit(st);
                    }
                }
                let menu = build_menu(st.locale);
                SetMenu(hwnd, menu);
                CheckMenuItem(
                    GetMenu(hwnd),
                    ID_VIEW_STATUS_BAR as u32,
                    MF_BYCOMMAND | if st.show_status_bar { MF_CHECKED } else { MF_UNCHECKED },
                );
                CheckMenuItem(
                    GetMenu(hwnd),
                    ID_FORMAT_WORD_WRAP as u32,
                    MF_BYCOMMAND | if st.word_wrap { MF_CHECKED } else { MF_UNCHECKED },
                );
                if !st.show_status_bar {
                    ShowWindow(st.status_hwnd, SW_HIDE);
                }
                if launch.settings.width > 0
                    && launch.settings.height > 0
                    && launch.settings.x != CW_USEDEFAULT
                    && launch.settings.y != CW_USEDEFAULT
                {
                    MoveWindow(
                        hwnd,
                        launch.settings.x,
                        launch.settings.y,
                        launch.settings.width,
                        launch.settings.height,
                        1,
                    );
                }
                update_title(hwnd, st);
                update_status_bar(st);
                layout_children(hwnd, st);
                if !st.current_path.is_empty() {
                    let path = st.current_path.clone();
                    open_path_into_editor(hwnd, st, &path, st.current_encoding);
                }
            }
            SetTimer(hwnd, ID_TIMER_STATUS, 120, None);
            0
        }
        WM_SIZE => {
            if let Some(st) = get_state(hwnd) {
                layout_children(hwnd, st);
            }
            0
        }
        WM_SETFOCUS => {
            if let Some(st) = get_state(hwnd) {
                SetFocus(st.edit_hwnd);
            }
            0
        }
        WM_COMMAND => {
            if let Some(st) = get_state(hwnd) {
                if loword(wparam) == ID_EDIT_CTRL && hiword(wparam) == EN_CHANGE as usize {
                    update_title(hwnd, st);
                    update_status_bar(st);
                    return 0;
                }
                match loword(wparam) {
                    ID_FILE_NEW => {
                        if confirm_save_if_needed(hwnd, st) {
                            set_edit_text(st.edit_hwnd, &[]);
                            SendMessageW(st.edit_hwnd, EM_SETMODIFY as u32, 0, 0);
                            st.current_path.clear();
                            update_title(hwnd, st);
                            update_status_bar(st);
                        }
                    }
                    ID_FILE_OPEN => {
                        if confirm_save_if_needed(hwnd, st) {
                            do_open_file(hwnd, st);
                        }
                    }
                    ID_FILE_SAVE => {
                        do_save(hwnd, st);
                    }
                    ID_FILE_SAVE_AS => {
                        do_save_as(hwnd, st);
                    }
                    ID_FILE_EXIT => {
                        SendMessageW(hwnd, WM_CLOSE, 0, 0);
                    }
                    ID_EDIT_UNDO => {
                        if SendMessageW(st.edit_hwnd, EM_CANUNDO as u32, 0, 0) != 0 {
                            SendMessageW(st.edit_hwnd, WM_UNDO, 0, 0);
                            update_title(hwnd, st);
                        }
                    }
                    ID_EDIT_CUT => {
                        SendMessageW(st.edit_hwnd, WM_CUT, 0, 0);
                        update_title(hwnd, st);
                    }
                    ID_EDIT_COPY => {
                        SendMessageW(st.edit_hwnd, WM_COPY, 0, 0);
                    }
                    ID_EDIT_PASTE => {
                        SendMessageW(st.edit_hwnd, WM_PASTE, 0, 0);
                        update_title(hwnd, st);
                    }
                    ID_EDIT_DELETE => {
                        SendMessageW(st.edit_hwnd, WM_CLEAR, 0, 0);
                        update_title(hwnd, st);
                    }
                    ID_EDIT_FIND => {
                        create_find_replace_dialog(hwnd, st, false);
                    }
                    ID_EDIT_FIND_NEXT => {
                        if st.last_find.is_empty() {
                            create_find_replace_dialog(hwnd, st, false);
                        } else if !do_find_next_with_flags(st, &st.last_find.clone(), st.last_find_flags | FR_DOWN) {
                            let msg = if std::ptr::eq(st.locale as *const Locale, &JA as *const Locale) {
                                "見つかりませんでした。"
                            } else {
                                "Text not found."
                            };
                            message_box(hwnd, msg, st.locale.app_caption, MB_OK);
                        }
                    }
                    ID_EDIT_REPLACE => {
                        create_find_replace_dialog(hwnd, st, true);
                    }
                    ID_EDIT_GOTO => {
                        let line_count = SendMessageW(st.edit_hwnd, EM_GETLINECOUNT as u32, 0, 0) as i32;
                        let font = GetStockObject(DEFAULT_GUI_FONT) as HFONT;
                        let bg = GetSysColorBrush(COLOR_3DFACE);
                        if let Some(target) =
                            show_goto_line_dialog(hwnd, st.locale, line_count.max(1), font, bg)
                        {
                            let idx = SendMessageW(st.edit_hwnd, EM_LINEINDEX as u32, (target - 1) as usize, 0);
                            if idx >= 0 {
                                SendMessageW(st.edit_hwnd, EM_SETSEL as u32, idx as usize, idx);
                                SetFocus(st.edit_hwnd);
                                update_status_bar(st);
                            }
                        }
                    }
                    ID_EDIT_SELECT_ALL => {
                        SendMessageW(st.edit_hwnd, EM_SETSEL as u32, 0, -1);
                        update_status_bar(st);
                    }
                    ID_EDIT_TIME_DATE => {
                        insert_date_time(st);
                        update_title(hwnd, st);
                        update_status_bar(st);
                    }
                    ID_VIEW_STATUS_BAR => {
                        st.show_status_bar = !st.show_status_bar;
                        if st.show_status_bar {
                            ShowWindow(st.status_hwnd, SW_SHOW);
                            CheckMenuItem(
                                GetMenu(hwnd),
                                ID_VIEW_STATUS_BAR as u32,
                                MF_BYCOMMAND | MF_CHECKED,
                            );
                        } else {
                            ShowWindow(st.status_hwnd, SW_HIDE);
                            CheckMenuItem(
                                GetMenu(hwnd),
                                ID_VIEW_STATUS_BAR as u32,
                                MF_BYCOMMAND | MF_UNCHECKED,
                            );
                        }
                        layout_children(hwnd, st);
                    }
                    ID_FORMAT_WORD_WRAP => {
                        toggle_word_wrap(hwnd, st);
                    }
                    ID_FORMAT_FONT => {
                        choose_display_font(hwnd, st);
                    }
                    ID_HELP_ABOUT => {
                        let about = format!(
                            "{}\n{} {}",
                            st.locale.about_title,
                            st.locale.version_label,
                            env!("CARGO_PKG_VERSION")
                        );
                        message_box(
                            hwnd,
                            &about,
                            st.locale.app_caption,
                            MB_OK,
                        );
                    }
                    _ => {}
                }
            }
            0
        }
        WM_DROPFILES => {
            if let Some(st) = get_state(hwnd) {
                let hdrop = wparam as HDROP;
                let file_count = DragQueryFileW(hdrop, u32::MAX, null_mut(), 0);
                if file_count > 0 && confirm_save_if_needed(hwnd, st) {
                    let len = DragQueryFileW(hdrop, 0, null_mut(), 0);
                    if len > 0 {
                        let mut path = vec![0u16; len as usize + 1];
                        let copied = DragQueryFileW(hdrop, 0, path.as_mut_ptr(), (len + 1) as u32);
                        path.truncate(copied as usize);
                        path.push(0);
                        open_path_into_editor(hwnd, st, &path, st.current_encoding);
                    }
                }
                DragFinish(hdrop);
            }
            0
        }
        WM_TIMER => {
            if wparam == ID_TIMER_STATUS {
                if let Some(st) = get_state(hwnd) {
                    update_status_bar(st);
                }
            }
            0
        }
        WM_CLOSE => {
            if let Some(st) = get_state(hwnd) {
                if !confirm_save_if_needed(hwnd, st) {
                    return 0;
                }
            }
            DestroyWindow(hwnd);
            0
        }
        WM_DESTROY => {
            PostQuitMessage(0);
            0
        }
        WM_NCDESTROY => {
            let ptr = GetWindowLongPtrW(hwnd, GWLP_USERDATA) as *mut AppState;
            if !ptr.is_null() {
                let mut state = Box::from_raw(ptr);
                if !state.hfont.is_null() {
                    DeleteObject(state.hfont as _);
                    state.hfont = null_mut();
                }
                SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
            }
            DefWindowProcW(hwnd, msg, wparam, lparam)
        }
        _ => DefWindowProcW(hwnd, msg, wparam, lparam),
    }
}

unsafe fn create_accelerators() -> windows_sys::Win32::UI::WindowsAndMessaging::HACCEL {
    let accels = [
        ACCEL {
            fVirt: (FVIRTKEY | FCONTROL) as u8,
            key: 'N' as u16,
            cmd: ID_FILE_NEW as u16,
        },
        ACCEL {
            fVirt: (FVIRTKEY | FCONTROL) as u8,
            key: 'O' as u16,
            cmd: ID_FILE_OPEN as u16,
        },
        ACCEL {
            fVirt: (FVIRTKEY | FCONTROL) as u8,
            key: 'S' as u16,
            cmd: ID_FILE_SAVE as u16,
        },
        ACCEL {
            fVirt: (FVIRTKEY | FCONTROL) as u8,
            key: 'A' as u16,
            cmd: ID_EDIT_SELECT_ALL as u16,
        },
        ACCEL {
            fVirt: (FVIRTKEY | FCONTROL) as u8,
            key: 'F' as u16,
            cmd: ID_EDIT_FIND as u16,
        },
        ACCEL {
            fVirt: FVIRTKEY as u8,
            key: VK_F3,
            cmd: ID_EDIT_FIND_NEXT as u16,
        },
        ACCEL {
            fVirt: (FVIRTKEY | FCONTROL) as u8,
            key: 'H' as u16,
            cmd: ID_EDIT_REPLACE as u16,
        },
        ACCEL {
            fVirt: (FVIRTKEY | FCONTROL) as u8,
            key: 'G' as u16,
            cmd: ID_EDIT_GOTO as u16,
        },
        ACCEL {
            fVirt: FVIRTKEY as u8,
            key: VK_F5,
            cmd: ID_EDIT_TIME_DATE as u16,
        },
    ];
    CreateAcceleratorTableW(accels.as_ptr(), accels.len() as i32)
}

fn main() {
    unsafe {
        let com_hr = CoInitializeEx(null(), COINIT_APARTMENTTHREADED as u32);
        let com_initialized = com_hr >= 0;
        if com_hr < 0 && com_hr != RPC_E_CHANGED_MODE {
            return;
        }
        let locale = choose_locale();
        let instance: HINSTANCE = GetModuleHandleW(null());
        let mut icc = INITCOMMONCONTROLSEX {
            dwSize: size_of::<INITCOMMONCONTROLSEX>() as u32,
            dwICC: ICC_BAR_CLASSES,
        };
        InitCommonControlsEx(&mut icc);

        let launch = Box::new(default_app_launch_data(locale));
        let launch_x = launch.settings.x;
        let launch_y = launch.settings.y;
        let launch_w = launch.settings.width;
        let launch_h = launch.settings.height;
        let launch_ptr = Box::into_raw(launch);

        let class_name = wide_null(APP_CLASS);
        let window_title = wide_null(APP_NAME);
        let wnd = WNDCLASSW {
            style: CS_HREDRAW | CS_VREDRAW,
            lpfnWndProc: Some(wnd_proc),
            hInstance: instance,
            hCursor: LoadCursorW(null_mut(), IDC_ARROW),
            hbrBackground: (COLOR_WINDOW as isize + 1) as _,
            lpszClassName: class_name.as_ptr(),
            ..zeroed()
        };
        RegisterClassW(&wnd);

        let hwnd = CreateWindowExW(
            0,
            class_name.as_ptr(),
            window_title.as_ptr(),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            launch_x,
            launch_y,
            launch_w,
            launch_h,
            null_mut(),
            null_mut(),
            instance,
            launch_ptr as *const _,
        );
        if hwnd.is_null() {
            let _ = Box::from_raw(launch_ptr);
            return;
        }

        ShowWindow(hwnd, SW_SHOW);

        let haccel = create_accelerators();
        let mut msg: MSG = zeroed();
        while GetMessageW(&mut msg, null_mut(), 0, 0) > 0 {
            if let Some(st) = get_state(hwnd) {
                if let Some(fd) = st.find_dialog.as_ref() {
                    if !fd.dlg_hwnd.is_null() && IsDialogMessageW(fd.dlg_hwnd, &msg) != 0 {
                        continue;
                    }
                }
            }
            if haccel.is_null() || TranslateAcceleratorW(hwnd, haccel, &msg) == 0 {
                TranslateMessage(&msg);
                DispatchMessageW(&msg);
            }
        }
        if !haccel.is_null() {
            DestroyAcceleratorTable(haccel);
        }
        if com_initialized {
            CoUninitialize();
        }
    }
}
