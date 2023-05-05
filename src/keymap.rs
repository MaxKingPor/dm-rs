//! 大漠键盘映射

#![allow(dead_code)]
#![warn(missing_docs)]

use super::KeyMap;

macro_rules! key {
    ($ident:ident, $key_str:expr, $id:expr) => {
        impl KeyMap<'_> {
            #[doc= concat!( "按键: ", stringify!($key_str), "\tid: ",stringify!($id)) ]
            pub const $ident: KeyMap<'static> = KeyMap {
                key_str: $key_str,
                id: $id,
            };
        }
    };
}

// pub const ONC:KeyMap = KeyMap{key_str: "1", id:49};

key!(KEY_0, "0", 48);
key!(KEY_1, "1", 49);
key!(KEY_2, "2", 50);
key!(KEY_3, "3", 51);
key!(KEY_4, "4", 52);
key!(KEY_5, "5", 53);
key!(KEY_6, "6", 54);
key!(KEY_7, "7", 55);
key!(KEY_8, "8", 56);
key!(KEY_9, "9", 57);

key!(KEY_MINUS, "-", 189);
key!(KEY_EQUAL, "=", 187);

key!(KEY_BACK, "back", 8);

key!(KEY_A, "a", 65);
key!(KEY_B, "b", 66);
key!(KEY_C, "c", 67);
key!(KEY_D, "d", 68);
key!(KEY_E, "e", 69);
key!(KEY_F, "f", 70);
key!(KEY_G, "g", 71);
key!(KEY_H, "h", 72);
key!(KEY_I, "i", 73);
key!(KEY_J, "j", 74);
key!(KEY_K, "k", 75);
key!(KEY_L, "l", 76);
key!(KEY_M, "m", 77);
key!(KEY_N, "n", 78);
key!(KEY_O, "o", 79);
key!(KEY_P, "p", 80);
key!(KEY_Q, "q", 81);
key!(KEY_R, "r", 82);
key!(KEY_S, "s", 83);
key!(KEY_T, "t", 84);
key!(KEY_U, "u", 85);
key!(KEY_V, "v", 86);
key!(KEY_W, "w", 87);
key!(KEY_X, "x", 88);
key!(KEY_Y, "y", 89);
key!(KEY_Z, "z", 90);

key!(KEY_CTRL, "ctrl", 17);
key!(KEY_ALT, "alt", 18);
key!(KEY_SHIFT, "shift", 16);
key!(KEY_WIN, "win", 91);
key!(KEY_SPACE, "space", 32);
key!(KEY_CAP, "cap", 20);
key!(KEY_TAB, "tab", 9);
key!(KEY_WAVY_LINES, "~", 192);
key!(KEY_ESC, "esc", 27);
key!(KEY_ENTER, "enter", 13);

key!(KEY_UP, "up", 38);
key!(KEY_DOWN, "down", 40);
key!(KEY_LEFT, "left", 37);
key!(KEY_RIGHT, "right", 39);

key!(KEY_OPTION, "option", 93);
key!(KEY_PRINT, "print", 44);
key!(KEY_DELETE, "delete", 46);
key!(KEY_HOME, "home", 36);
key!(KEY_END, "end", 35);
key!(KEY_PGUP, "pgup", 33);
key!(KEY_PGDN, "pgdn", 34);

key!(KEY_F1, "f1", 112);
key!(KEY_F2, "f2", 113);
key!(KEY_F3, "f3", 114);
key!(KEY_F4, "f4", 115);
key!(KEY_F5, "f5", 116);
key!(KEY_F6, "f6", 117);
key!(KEY_F7, "f7", 118);
key!(KEY_F8, "f8", 119);
key!(KEY_F9, "f9", 120);
key!(KEY_F10, "f10", 121);
key!(KEY_F11, "f11", 122);
key!(KEY_F12, "f12", 123);

key!(KEY_OPEN_BRACKET, "[", 219);
key!(KEY_CLOSE_BRACKET, "]", 221);
key!(KEY_BACKSLASH, "\\", 220);
key!(KEY_SEMICOLON, ";", 186);
key!(KEY_SINGLE_QUOTES, "'", 222);
key!(KEY_COMMA, ",", 188);
key!(KEY_DOT, ".", 190);
key!(KEY_SLASH, "/", 191);
