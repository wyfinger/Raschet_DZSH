Attribute VB_Name = "Raschet_DZSH"
'
' Макрос ДЗШ
' ~~~~~~~~~~
' https://github.com/wyfinger/Raschet_DZSH
' Игорь Матвеев, miv@prim.so-ups.ru
' 2013-2022
'

Option Explicit

Private Declare PtrSafe Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Function SendMessageStr Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare PtrSafe Function GetWindow Lib "User32" (ByVal hwnd As Long, ByVal uCmd As Long) As Long
Private Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAcess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare PtrSafe Function GetModuleFileNameEx Lib "PSAPI" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function IsWindowVisible Lib "User32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetParent Lib "User32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function EnumWindows Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare PtrSafe Function CreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare PtrSafe Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare PtrSafe Function DestroyWindow Lib "User32" (ByVal hwnd As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare PtrSafe Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer

Private Const WM_COMMAND = &H111
Private Const WM_PASTE = &H302
Private Const WM_CUT = &H300
Private Const GW_CHILD = &H5
Private Const GW_HWNDNEXT = &H2
Private Const EM_SETSEL = &HB1
Private Const WM_SETTEXT = &HC
Private Const WM_GETTEXT As Integer = &HD
Private Const WM_GETTEXTLENGTH As Integer = &HE
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_SIZEBOX = &H40000
Private Const WS_CAPTION = &HC00000
Private Const SW_NORMAL = 1

Private Const DefaultBufferSize& = 32768
Private CRC_32_Tab(0 To 255) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim GT_Class As String         ' Временная переменная класса для передачи в Enum-функцию
Dim GT_Result As Long

Dim RootNode As String         ' Самый главный узел :)
Dim arrRootBranch()            ' Список ветвей главного узла, заполняется в Get_Sensitivity_Code()

Dim arrBranch()                ' Массив ветвей и его копии
Dim arrBranchCopy()
Dim arrBranchCopy2()
Dim arrNode()                  ' Массив наименований узлов
Dim arrElement()               ' Массив наименований элементов
  
Dim arrPowerNodes()            ' Массив питающих узлов (узел и первая ветвь от RootNode в сторону питающего узла)
                               ' заполняется в Find_Power_Nodes()
Dim arrBaseRejims()            ' Массив базовых режимов для проверки чувствительности,
                               ' заполняется при подготовке приказа в Get_Testing_Code()
                               ' т.к. из протокола расчета эту инфу не получить

' Menu items

Dim TKZ_MENU_ADV_MODE ' = 14     ' Расширенный формат задания для расчета...
Dim TKZ_MENU_OPEN_LOG '= 183    ' Открыть протокол...
Dim TKZ_MENU_CLEAR_CODE '= 201  ' Очистить задание
Dim TKZ_MENU_CLEAR_LOG '= 252   ' Очистить протокол
Dim TKZ_MENU_CALC '= 175        ' Расчет с эквиваленированием


'##########################################################################[ Функции взаимодействия с окнами ]

Private Function Exe_Name_by_Window_Handle(wnd As Long) As String
'
' Получение имени EXE файла по Handle окна
'

  Exe_Name_by_Window_Handle = ""

  Dim prcID As Long  ' Номер процесса
  Dim prc As Long    ' Дескриптор доступа к данным процесса
  Dim wt As String
  wt = Space(1024)

  If GetWindowThreadProcessId(wnd, prcID) <> 0 Then         ' Ищем процесс
    prc = OpenProcess(PROCESS_ALL_ACCESS, False, prcID)     ' Открываем
    On Error GoTo Finally
    If GetModuleFileNameEx(prc, 0, wt, 1024) <> 0 Then      ' Получаем имя EXE
      Exe_Name_by_Window_Handle = Trim$(wt)
    End If
Finally:
    CloseHandle prc
  End If
  On Error GoTo 0

End Function


Private Function Extract_File_Name(FileName As String) As String
'
' Получение имени файла из полного пути и имени
'

  Dim SlashPos As Long

  If FileName <> "" Then
    SlashPos = InStrRev(FileName, "\")
    If SlashPos > 0 Then
      SlashPos = SlashPos + 1
      Extract_File_Name = Mid(FileName, SlashPos, Len(FileName) - SlashPos)
    End If
  End If

End Function


Private Function Extract_File_Path(FileName As String) As String
'
' Получение пути к папке, содержащей файл (без завершающего слеша)
'

  Dim SlashPos As Long

  If FileName <> "" Then
    SlashPos = InStrRev(FileName, "\")
    If SlashPos > 0 Then
      SlashPos = SlashPos + 1
      Extract_File_Name = Mid(FileName, 1, SlashPos - 1)
    End If
  End If

End Function


Private Function Get_Class_Name(ByVal wnd As Long) As String
'
' Обертка над стандартной API функцией
'

  Dim ClassName As String
  Dim ClassLen As Long
  ClassName = Space(256)
  ClassLen = GetClassName(wnd, ClassName, 256)
  Get_Class_Name = Mid(ClassName, 1, ClassLen)

End Function


Public Function Find_Window_Enum_Proc(ByVal wnd As Long, ByVal lParam As Long) As Boolean
'
' Функция обработчик перечисления всех окон в системе для поиска главного окна ТКЗ-2000
'

  Find_Window_Enum_Proc = True

  Dim ExeName As String
  Dim WndClass As String

  If GetParent(wnd) = 0 Then
  
    ExeName = UCase(Extract_File_Name(Exe_Name_by_Window_Handle(wnd)))
    If (ExeName = "TKZ2000.EXE") Then
      WndClass = Get_Class_Name(wnd)
      If WndClass = GT_Class Then
        GT_Result = wnd
        Find_Window_Enum_Proc = False
      End If
    End If

  End If

End Function


Private Function Find_TKZ_Window_Handle(ByVal Class As String) As Long
'
' Поиск главного окна программы ТКЗ-2000
'

  GT_Result = 0
  GT_Class = Class
  Call EnumWindows(AddressOf Find_Window_Enum_Proc, 0)
  Find_TKZ_Window_Handle = GT_Result

End Function


Private Function Find_SubClass_Recurce(hwnd As Long, sClassName As String, Optional iPos As Integer = 1) As Long
'
' Эта функция рекурсивно перебирает все дочерние окна hWnd, сверяя класс окна с sClassName
' Если указан iPos функция возвращает iPos вхождение интересующего класса
'

  Dim window_class As String

  window_class = Get_Class_Name(hwnd)
  If window_class = sClassName Then
    iPos = iPos - 1
    If iPos = 0 Then
      Find_SubClass_Recurce = hwnd
      Exit Function
    End If
  End If

  Dim hext_handle As Long

  hext_handle = GetWindow(hwnd, GW_CHILD)
  If hext_handle > 0 Then Find_SubClass_Recurce = Find_SubClass_Recurce(hext_handle, sClassName, iPos)
  If Find_SubClass_Recurce > 0 Then Exit Function

  hext_handle = GetWindow(hwnd, GW_HWNDNEXT)
  If hext_handle > 0 Then Find_SubClass_Recurce = Find_SubClass_Recurce(hext_handle, sClassName, iPos)

End Function


Private Function Window_Set_Text(hwnd As Long, sText As String)
'
' Установить в поле ввода текст

SendMessageStr hwnd, WM_SETTEXT, 0, sText

End Function


Private Function Window_Get_Text(hwnd As Long)
'
' Забираем текст из поля ввода

Dim TextLen As Long
TextLen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0) + 1

Dim sStr As String
sStr = Space(TextLen)

Call SendMessageStr(hwnd, WM_GETTEXT, TextLen, sStr)
Window_Get_Text = sStr

End Function


Private Function Show_Process_Window(Caption As String) As Long
'
' Отображение диалога процесса
'

  Dim WinRect As RECT
  Dim cs
  Dim X, Y As Long

  GetWindowRect Application.hwnd, WinRect

  X = WinRect.Left + (WinRect.Right - WinRect.Left) / 2 - 100
  Y = WinRect.Top + (WinRect.Bottom - WinRect.Top) / 2 - 50

  Show_Process_Window = CreateWindowEx(WS_EX_TOOLWINDOW, "MDICLIENT", Caption, _
    WS_SIZEBOX Or WS_CAPTION, X, Y, 200, 100, Application.hwnd, 0, Application.hInstance, cs)
  ShowWindow Show_Process_Window, SW_NORMAL

End Function


Private Sub CRCinit()
'
' https://github.com/slash-cyberpunk/VBA-CRC32/blob/master/CRC32.bas
'

  CRC_32_Tab(0) = &H0
  CRC_32_Tab(1) = &H77073096
  CRC_32_Tab(2) = &HEE0E612C
  CRC_32_Tab(3) = &H990951BA
  CRC_32_Tab(4) = &H76DC419
  CRC_32_Tab(5) = &H706AF48F
  CRC_32_Tab(6) = &HE963A535
  CRC_32_Tab(7) = &H9E6495A3
  CRC_32_Tab(8) = &HEDB8832
  CRC_32_Tab(9) = &H79DCB8A4
  CRC_32_Tab(10) = &HE0D5E91E
  CRC_32_Tab(11) = &H97D2D988
  CRC_32_Tab(12) = &H9B64C2B
  CRC_32_Tab(13) = &H7EB17CBD
  CRC_32_Tab(14) = &HE7B82D07
  CRC_32_Tab(15) = &H90BF1D91
  CRC_32_Tab(16) = &H1DB71064
  CRC_32_Tab(17) = &H6AB020F2
  CRC_32_Tab(18) = &HF3B97148
  CRC_32_Tab(19) = &H84BE41DE
  CRC_32_Tab(20) = &H1ADAD47D
  CRC_32_Tab(21) = &H6DDDE4EB
  CRC_32_Tab(22) = &HF4D4B551
  CRC_32_Tab(23) = &H83D385C7
  CRC_32_Tab(24) = &H136C9856
  CRC_32_Tab(25) = &H646BA8C0
  CRC_32_Tab(26) = &HFD62F97A
  CRC_32_Tab(27) = &H8A65C9EC
  CRC_32_Tab(28) = &H14015C4F
  CRC_32_Tab(29) = &H63066CD9
  CRC_32_Tab(30) = &HFA0F3D63
  CRC_32_Tab(31) = &H8D080DF5
  CRC_32_Tab(32) = &H3B6E20C8
  CRC_32_Tab(33) = &H4C69105E
  CRC_32_Tab(34) = &HD56041E4
  CRC_32_Tab(35) = &HA2677172
  CRC_32_Tab(36) = &H3C03E4D1
  CRC_32_Tab(37) = &H4B04D447
  CRC_32_Tab(38) = &HD20D85FD
  CRC_32_Tab(39) = &HA50AB56B
  CRC_32_Tab(40) = &H35B5A8FA
  CRC_32_Tab(41) = &H42B2986C
  CRC_32_Tab(42) = &HDBBBC9D6
  CRC_32_Tab(43) = &HACBCF940
  CRC_32_Tab(44) = &H32D86CE3
  CRC_32_Tab(45) = &H45DF5C75
  CRC_32_Tab(46) = &HDCD60DCF
  CRC_32_Tab(47) = &HABD13D59
  CRC_32_Tab(48) = &H26D930AC
  CRC_32_Tab(49) = &H51DE003A
  CRC_32_Tab(50) = &HC8D75180
  CRC_32_Tab(51) = &HBFD06116
  CRC_32_Tab(52) = &H21B4F4B5
  CRC_32_Tab(53) = &H56B3C423
  CRC_32_Tab(54) = &HCFBA9599
  CRC_32_Tab(55) = &HB8BDA50F
  CRC_32_Tab(56) = &H2802B89E
  CRC_32_Tab(57) = &H5F058808
  CRC_32_Tab(58) = &HC60CD9B2
  CRC_32_Tab(59) = &HB10BE924
  CRC_32_Tab(60) = &H2F6F7C87
  CRC_32_Tab(61) = &H58684C11
  CRC_32_Tab(62) = &HC1611DAB
  CRC_32_Tab(63) = &HB6662D3D
  CRC_32_Tab(64) = &H76DC4190
  CRC_32_Tab(65) = &H1DB7106
  CRC_32_Tab(66) = &H98D220BC
  CRC_32_Tab(67) = &HEFD5102A
  CRC_32_Tab(68) = &H71B18589
  CRC_32_Tab(69) = &H6B6B51F
  CRC_32_Tab(70) = &H9FBFE4A5
  CRC_32_Tab(71) = &HE8B8D433
  CRC_32_Tab(72) = &H7807C9A2
  CRC_32_Tab(73) = &HF00F934
  CRC_32_Tab(74) = &H9609A88E
  CRC_32_Tab(75) = &HE10E9818
  CRC_32_Tab(76) = &H7F6A0DBB
  CRC_32_Tab(77) = &H86D3D2D
  CRC_32_Tab(78) = &H91646C97
  CRC_32_Tab(79) = &HE6635C01
  CRC_32_Tab(80) = &H6B6B51F4
  CRC_32_Tab(81) = &H1C6C6162
  CRC_32_Tab(82) = &H856530D8
  CRC_32_Tab(83) = &HF262004E
  CRC_32_Tab(84) = &H6C0695ED
  CRC_32_Tab(85) = &H1B01A57B
  CRC_32_Tab(86) = &H8208F4C1
  CRC_32_Tab(87) = &HF50FC457
  CRC_32_Tab(88) = &H65B0D9C6
  CRC_32_Tab(89) = &H12B7E950
  CRC_32_Tab(90) = &H8BBEB8EA
  CRC_32_Tab(91) = &HFCB9887C
  CRC_32_Tab(92) = &H62DD1DDF
  CRC_32_Tab(93) = &H15DA2D49
  CRC_32_Tab(94) = &H8CD37CF3
  CRC_32_Tab(95) = &HFBD44C65
  CRC_32_Tab(96) = &H4DB26158
  CRC_32_Tab(97) = &H3AB551CE
  CRC_32_Tab(98) = &HA3BC0074
  CRC_32_Tab(99) = &HD4BB30E2
  CRC_32_Tab(100) = &H4ADFA541
  CRC_32_Tab(101) = &H3DD895D7
  CRC_32_Tab(102) = &HA4D1C46D
  CRC_32_Tab(103) = &HD3D6F4FB
  CRC_32_Tab(104) = &H4369E96A
  CRC_32_Tab(105) = &H346ED9FC
  CRC_32_Tab(106) = &HAD678846
  CRC_32_Tab(107) = &HDA60B8D0
  CRC_32_Tab(108) = &H44042D73
  CRC_32_Tab(109) = &H33031DE5
  CRC_32_Tab(110) = &HAA0A4C5F
  CRC_32_Tab(111) = &HDD0D7CC9
  CRC_32_Tab(112) = &H5005713C
  CRC_32_Tab(113) = &H270241AA
  CRC_32_Tab(114) = &HBE0B1010
  CRC_32_Tab(115) = &HC90C2086
  CRC_32_Tab(116) = &H5768B525
  CRC_32_Tab(117) = &H206F85B3
  CRC_32_Tab(118) = &HB966D409
  CRC_32_Tab(119) = &HCE61E49F
  CRC_32_Tab(120) = &H5EDEF90E
  CRC_32_Tab(121) = &H29D9C998
  CRC_32_Tab(122) = &HB0D09822
  CRC_32_Tab(123) = &HC7D7A8B4
  CRC_32_Tab(124) = &H59B33D17
  CRC_32_Tab(125) = &H2EB40D81
  CRC_32_Tab(126) = &HB7BD5C3B
  CRC_32_Tab(127) = &HC0BA6CAD
  CRC_32_Tab(128) = &HEDB88320
  CRC_32_Tab(129) = &H9ABFB3B6
  CRC_32_Tab(130) = &H3B6E20C
  CRC_32_Tab(131) = &H74B1D29A
  CRC_32_Tab(132) = &HEAD54739
  CRC_32_Tab(133) = &H9DD277AF
  CRC_32_Tab(134) = &H4DB2615
  CRC_32_Tab(135) = &H73DC1683
  CRC_32_Tab(136) = &HE3630B12
  CRC_32_Tab(137) = &H94643B84
  CRC_32_Tab(138) = &HD6D6A3E
  CRC_32_Tab(139) = &H7A6A5AA8
  CRC_32_Tab(140) = &HE40ECF0B
  CRC_32_Tab(141) = &H9309FF9D
  CRC_32_Tab(142) = &HA00AE27
  CRC_32_Tab(143) = &H7D079EB1
  CRC_32_Tab(144) = &HF00F9344
  CRC_32_Tab(145) = &H8708A3D2
  CRC_32_Tab(146) = &H1E01F268
  CRC_32_Tab(147) = &H6906C2FE
  CRC_32_Tab(148) = &HF762575D
  CRC_32_Tab(149) = &H806567CB
  CRC_32_Tab(150) = &H196C3671
  CRC_32_Tab(151) = &H6E6B06E7
  CRC_32_Tab(152) = &HFED41B76
  CRC_32_Tab(153) = &H89D32BE0
  CRC_32_Tab(154) = &H10DA7A5A
  CRC_32_Tab(155) = &H67DD4ACC
  CRC_32_Tab(156) = &HF9B9DF6F
  CRC_32_Tab(157) = &H8EBEEFF9
  CRC_32_Tab(158) = &H17B7BE43
  CRC_32_Tab(159) = &H60B08ED5
  CRC_32_Tab(160) = &HD6D6A3E8
  CRC_32_Tab(161) = &HA1D1937E
  CRC_32_Tab(162) = &H38D8C2C4
  CRC_32_Tab(163) = &H4FDFF252
  CRC_32_Tab(164) = &HD1BB67F1
  CRC_32_Tab(165) = &HA6BC5767
  CRC_32_Tab(166) = &H3FB506DD
  CRC_32_Tab(167) = &H48B2364B
  CRC_32_Tab(168) = &HD80D2BDA
  CRC_32_Tab(169) = &HAF0A1B4C
  CRC_32_Tab(170) = &H36034AF6
  CRC_32_Tab(171) = &H41047A60
  CRC_32_Tab(172) = &HDF60EFC3
  CRC_32_Tab(173) = &HA867DF55
  CRC_32_Tab(174) = &H316E8EEF
  CRC_32_Tab(175) = &H4669BE79
  CRC_32_Tab(176) = &HCB61B38C
  CRC_32_Tab(177) = &HBC66831A
  CRC_32_Tab(178) = &H256FD2A0
  CRC_32_Tab(179) = &H5268E236
  CRC_32_Tab(180) = &HCC0C7795
  CRC_32_Tab(181) = &HBB0B4703
  CRC_32_Tab(182) = &H220216B9
  CRC_32_Tab(183) = &H5505262F
  CRC_32_Tab(184) = &HC5BA3BBE
  CRC_32_Tab(185) = &HB2BD0B28
  CRC_32_Tab(186) = &H2BB45A92
  CRC_32_Tab(187) = &H5CB36A04
  CRC_32_Tab(188) = &HC2D7FFA7
  CRC_32_Tab(189) = &HB5D0CF31
  CRC_32_Tab(190) = &H2CD99E8B
  CRC_32_Tab(191) = &H5BDEAE1D
  CRC_32_Tab(192) = &H9B64C2B0
  CRC_32_Tab(193) = &HEC63F226
  CRC_32_Tab(194) = &H756AA39C
  CRC_32_Tab(195) = &H26D930A
  CRC_32_Tab(196) = &H9C0906A9
  CRC_32_Tab(197) = &HEB0E363F
  CRC_32_Tab(198) = &H72076785
  CRC_32_Tab(199) = &H5005713
  CRC_32_Tab(200) = &H95BF4A82
  CRC_32_Tab(201) = &HE2B87A14
  CRC_32_Tab(202) = &H7BB12BAE
  CRC_32_Tab(203) = &HCB61B38
  CRC_32_Tab(204) = &H92D28E9B
  CRC_32_Tab(205) = &HE5D5BE0D
  CRC_32_Tab(206) = &H7CDCEFB7
  CRC_32_Tab(207) = &HBDBDF21
  CRC_32_Tab(208) = &H86D3D2D4
  CRC_32_Tab(209) = &HF1D4E242
  CRC_32_Tab(210) = &H68DDB3F8
  CRC_32_Tab(211) = &H1FDA836E
  CRC_32_Tab(212) = &H81BE16CD
  CRC_32_Tab(213) = &HF6B9265B
  CRC_32_Tab(214) = &H6FB077E1
  CRC_32_Tab(215) = &H18B74777
  CRC_32_Tab(216) = &H88085AE6
  CRC_32_Tab(217) = &HFF0F6A70
  CRC_32_Tab(218) = &H66063BCA
  CRC_32_Tab(219) = &H11010B5C
  CRC_32_Tab(220) = &H8F659EFF
  CRC_32_Tab(221) = &HF862AE69
  CRC_32_Tab(222) = &H616BFFD3
  CRC_32_Tab(223) = &H166CCF45
  CRC_32_Tab(224) = &HA00AE278
  CRC_32_Tab(225) = &HD70DD2EE
  CRC_32_Tab(226) = &H4E048354
  CRC_32_Tab(227) = &H3903B3C2
  CRC_32_Tab(228) = &HA7672661
  CRC_32_Tab(229) = &HD06016F7
  CRC_32_Tab(230) = &H4969474D
  CRC_32_Tab(231) = &H3E6E77DB
  CRC_32_Tab(232) = &HAED16A4A
  CRC_32_Tab(233) = &HD9D65ADC
  CRC_32_Tab(234) = &H40DF0B66
  CRC_32_Tab(235) = &H37D83BF0
  CRC_32_Tab(236) = &HA9BCAE53
  CRC_32_Tab(237) = &HDEBB9EC5
  CRC_32_Tab(238) = &H47B2CF7F
  CRC_32_Tab(239) = &H30B5FFE9
  CRC_32_Tab(240) = &HBDBDF21C
  CRC_32_Tab(241) = &HCABAC28A
  CRC_32_Tab(242) = &H53B39330
  CRC_32_Tab(243) = &H24B4A3A6
  CRC_32_Tab(244) = &HBAD03605
  CRC_32_Tab(245) = &HCDD70693
  CRC_32_Tab(246) = &H54DE5729
  CRC_32_Tab(247) = &H23D967BF
  CRC_32_Tab(248) = &HB3667A2E
  CRC_32_Tab(249) = &HC4614AB8
  CRC_32_Tab(250) = &H5D681B02
  CRC_32_Tab(251) = &H2A6F2B94
  CRC_32_Tab(252) = &HB40BBE37
  CRC_32_Tab(253) = &HC30C8EA1
  CRC_32_Tab(254) = &H5A05DF1B
  CRC_32_Tab(255) = &H2D02EF8D
End Sub

Private Function Shr(n As Long, m As Integer) As Long

  Dim Q&

  If (m > 31) Then
    Shr = 0
    Exit Function
  End If
  If (n >= 0) Then
    Shr = n \ (2 ^ m)
  Else
    Q& = n And &H7FFFFFFF
    Q& = Q& \ (2 ^ m)
    Shr = Q& Or (2 ^ (31 - m))
  End If
End Function

Private Function Calc(Stri As String) As Long
'
' https://github.com/slash-cyberpunk/VBA-CRC32/blob/master/CRC32.bas
'

  Dim tCRC32&, m&, i&, n&

  CRCinit
  tCRC32& = &HFFFFFFFF
  For i& = 1 To Len(Stri)
    m& = Asc(Mid$(Stri, i&, 1))
    n& = (tCRC32& Xor m&) And &HFF
    tCRC32& = CRC_32_Tab(n&) Xor (Shr(tCRC32&, 8) And &HFFFFFF)
  Next i&
  Calc = -(tCRC32& + 1)
  
End Function

Private Function CalcStr(Stri As String) As String
'
' https://github.com/slash-cyberpunk/VBA-CRC32/blob/master/CRC32.bas
'

  CalcStr = CStr(Hex$(Calc(Stri)))
End Function

Private Function CalcFile(File As String, Optional BufferSize As Long = DefaultBufferSize&) As String
'
' https://github.com/slash-cyberpunk/VBA-CRC32/blob/master/CRC32.bas
'

  Dim FileSize&, FileNumber%, Buffer$, Code$, modCrc&
  
  If Len(Dir$(File)) > 0 And BufferSize > 0 Then
    FileSize& = FileLen(File)
    FileNumber% = FreeFile()
    Open File For Binary Access Read As #FileNumber%
    Buffer$ = String(BufferSize, 0)
    While (FileSize& - Seek(FileNumber%) + 1) >= BufferSize
        Get #FileNumber%, , Buffer$
        Code$ = Code$ & Buffer$
    Wend
    Buffer$ = String(FileSize& - Seek(FileNumber%) + 1, 0)
    Get #FileNumber%, , Buffer$
    Code$ = Code$ & Buffer$
    Close #FileNumber%
    modCrc& = Calc(Code$)
    CalcFile = CStr(Hex$(modCrc&))
  Else
    CalcFile = ""
  End If
  
End Function

'#################################################################################################[Расчет ДЗШ]

Private Sub Initialize(TKZPath As String)
'
' Подготовка - изымаем данные с листа в переменные
'

  Dim wshBranch
  Dim wshNode
  Dim wshElement

  ' 1. Ветви - берем первые пять столбцов страницы 'Таблица ветвей'
  Set wshBranch = ActiveWorkbook.Worksheets("Таблица ветвей")
  arrBranch = wshBranch.Range("A3:K" & wshBranch.UsedRange.Rows.Count).Value2
  arrBranchCopy = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2
  arrBranchCopy2 = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2

  ' 2. Узлы - берем первые три столбца страницы 'Наим.узлов'
  Set wshNode = ActiveWorkbook.Worksheets("Наим.узлов")
  arrNode = wshNode.Range("A3:E" & wshNode.UsedRange.Rows.Count).Value2

  ' 3. Элементы - берем первые два столбца страницы 'Наим.элементов'
  Set wshElement = ActiveWorkbook.Worksheets("Наим.элементов")
  arrElement = wshElement.Range("A3:B" & wshElement.UsedRange.Rows.Count).Value2

  ' Очистка памяти
  Set wshBranch = Nothing
  Set wshNode = Nothing
  Set wshElement = Nothing
  
  ' Номера используемых пунктов меню (зависят от версии ТКЗ)
  Dim TKZCrc
  TKZCrc = CalcFile(TKZPath)
  
  If TKZCrc = "4AE065E6" Or _
     TKZCrc = "FCC29215" Then       ' 5.7 or 5.6
    TKZ_MENU_ADV_MODE = 14          ' Расширенный формат задания для расчета...
    TKZ_MENU_OPEN_LOG = 194         ' Открыть протокол...
    TKZ_MENU_CLEAR_CODE = 212       ' Очистить задание
    TKZ_MENU_CLEAR_LOG = 263        ' Очистить протокол
    TKZ_MENU_CALC = 186             ' Расчет с эквиваленированием
  ElseIf TKZCrc = "57CE9188" Then   ' 5.7 menu fix
    TKZ_MENU_ADV_MODE = 14
    TKZ_MENU_OPEN_LOG = 183
    TKZ_MENU_CLEAR_CODE = 201
    TKZ_MENU_CLEAR_LOG = 252
    TKZ_MENU_CALC = 175
  End If
  
End Sub


Private Function Find_Branch_By_Node(BranchArray, Node)
'
' Ищем все ветви, в которые входит заданный узел
'

  Dim Rez()
  Dim i As Integer
  Dim j As Integer
  j = 0
  For i = LBound(BranchArray) To UBound(BranchArray)
    If (Trim$(BranchArray(i, 3)) = Node) Or (Trim$(BranchArray(i, 4)) = Node) Then
      ReDim Preserve Rez(j)
      Rez(j) = i
      j = j + 1
    End If
  Next
  Find_Branch_By_Node = Rez

End Function


Private Function Find_Branch_By_2Node(BranchArray, Node1 As String, Node2 As String)
'
' Ищем одну ветвь по имеющимся двум узлам, ее образующим, возвращаем номер ветви
'
  Dim i, Rez As Integer
  Rez = -1

  For i = LBound(BranchArray) To UBound(BranchArray)
    If ((Trim$(BranchArray(i, 3)) = Node1) And (Trim$(BranchArray(i, 4)) = Node2)) Or _
       ((Trim$(BranchArray(i, 3)) = Node2) And (Trim$(BranchArray(i, 4)) = Node1)) Then
      Find_Branch_By_2Node = i
      Exit For
    End If
  Next

End Function


Private Function Find_Node(Node As String) As String
'
' Ищем наименование узла по номеру
'

  Dim i As Integer
  For i = LBound(arrNode) To UBound(arrNode)
    If Trim$(arrNode(i, 1)) = Trim$(Node) Then
      Find_Node = Trim$(arrNode(i, 2))
      Exit For
    End If
  Next

End Function


Private Function Find_Node_Index(Node As String) As Integer
'
' Ищем индекс узла по номеру
'

  Find_Node_Index = -1
  Dim i As Integer
  For i = LBound(arrNode) To UBound(arrNode)
    If Trim$(arrNode(i, 1)) = Node Then
      Find_Node_Index = i
      Exit For
    End If
  Next

End Function


Private Function Node_Exists(Node As String) As Boolean
'
' Проверка существования узла
'

  Dim i As Integer
  Node_Exists = False
  For i = LBound(arrNode) To UBound(arrNode)
    If Trim$(arrNode(i, 1)) = Node Then
      Node_Exists = True
      Exit For
    End If
  Next

End Function


Private Function Find_Element(Element As Long) As String
'
' Ищем наименование элемента по номеру
'

  Dim i As Integer
  For i = LBound(arrElement) To UBound(arrElement)
    If Int(arrElement(i, 1)) = Element Then
      Find_Element = Trim$(arrElement(i, 2))
      Exit For
    End If
  Next

End Function


Private Function Find_Element_By_2Node(Branch, Node1 As String, Node2 As String)
'
' Ищем список элементов, которые имеют в своем составе узлы Node1 и Node2
'

  Dim Rez()
  Dim i, j As Long
  j = 0

  For i = LBound(Branch) To UBound(Branch)
    If ((Branch(i, 3) = Node1) And (Branch(i, 4) = Node2)) Or _
       ((Branch(i, 3) = Node2) And (Branch(i, 4) = Node1)) Then
        ReDim Preserve Rez(j)
        Rez(j) = Branch(i, 5)
        j = j + 1
    End If
  Next
  Find_Element_By_2Node = Rez

End Function


Private Function Find_Branch_Index(Node1, Node2)
'
' Ищем индекс ветви по узлам слева и справа, порядок узлов не имеет значения
'

  Dim i As Integer
  Find_Branch_Index = 0
  For i = LBound(arrBranch) To UBound(arrBranch)
    If ((arrBranch(i, 3) = Node1) And (arrBranch(i, 4) = Node2)) Or _
       ((arrBranch(i, 3) = Node2) And (arrBranch(i, 4) = Node1)) Then
       Find_Branch_Index = i
       Exit For
     End If
  Next

End Function


Private Function Array_Exists(ByRef arr()) As Boolean
'
' Проверка инициализированности массива
'

Dim TempVar As Integer
On Error GoTo NotExists
  TempVar = UBound(arr)
  Array_Exists = True
  On Error GoTo 0
  Exit Function
  
NotExists:
  Array_Exists = False
  On Error GoTo 0

End Function


Private Function Array_Find(Source(), Val, Optional col As Integer = -1) As Integer
'
' Проверка содержания в массиве Source значения Val в столбце Col,
' если Col = -1 считаем, что массив одномерный.
' Возвращаем индекс элемента (первый) или -1, если не найдено
'

  Dim i As Integer
  Array_Find = -1

  If Not Array_Exists(Source) Then Exit Function

  For i = LBound(Source) To UBound(Source)
    If col = -1 Then
      If Source(i) = Val Then
        Array_Find = i
      End If
    Else
      If Source(col, i) = Val Then
        Array_Find = i
      End If
    End If
  Next
  
End Function


Private Function Get_Sensitivity_Code() As String
'
' Подготовка приказа для оценки чувствительности ДЗШ
' Ищем все присоединения текущего узла, делаем КЗ на узле в нормальном режиме,
' а в подрежимах отключаем по одному присоединению
'

  Dim R As String

  R = _
    "*         ПРОВЕРКА ЧУВСТВИТЕЛЬНОСТИ ДЗШ, УЗЕЛ " & RootNode & " [" & Find_Node(RootNode) & "]" & vbCrLf & _
    "ВЕЛИЧИНА  IA IB IC" & vbCrLf & _
    "1-ПОЯС    " & RootNode & "      /* " & Find_Node(RootNode) & vbCrLf & _
    "СНСМ      1" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/ABC" & vbCrLf & _
    "СНСМ      2" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/AB" & vbCrLf & _
    "СНСМ      3" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/AB0" & vbCrLf & _
    "СНСМ      4" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/A0" & vbCrLf & _
    "ПОДРЕЖИМ  1    /* ВСЕ ВКЛЮЧЕНО" & vbCrLf

  Dim CrossNode As String
  Dim ElemNo As Long
  Dim BranchType As Long
  Dim ElemName As String
  Dim NewStr As String
  Dim i As Long

  ' Для каждого из присоединений RootNode создаем подрежим, где отключаем присоединения
  arrRootBranch = Find_Branch_By_Node(arrBranch, RootNode)
  For i = LBound(arrRootBranch) To UBound(arrRootBranch)
    If arrBranch(arrRootBranch(i), 3) = RootNode Then
      CrossNode = arrBranch(arrRootBranch(i), 4)
    Else
      CrossNode = arrBranch(arrRootBranch(i), 3)
    End If
    ElemNo = arrBranch(arrRootBranch(i), 5)
    BranchType = arrBranch(arrRootBranch(i), 1)
    ' Issue#8: Незачем отключать отключенные ШСВ
    If BranchType <> 101 Then
      ElemName = Find_Element(ElemNo)
      If (CrossNode = "0") Or (ElemNo = 0) Then               ' отключаем ветвь
        NewStr = "ПОДРЕЖИМ  " & (i + 2) & vbCrLf & _
        "ОТКЛ      0 *" & RootNode & "-" & CrossNode & "      /* НЕЙТРАЛЬ ?"
      Else                                                  ' отключаем элемент
        NewStr = "ПОДРЕЖИМ  " & (i + 2) & vbCrLf & _
        "ЭЛЕМЕНТ   " & ElemNo & "      /* " & ElemName
      End If
      R = R & NewStr & vbCrLf
    End If
  Next

  Get_Sensitivity_Code = R

End Function


Private Function Parse_Current_Line(T, FromPos, ByRef FinishPos)
'
' Функция ищет суммарные токи КЗ в несимметриях в тексте протокола Т начиная с
' FromPos, в параметр FinishPos записывается позиция, на которой остановился парсер
'

  Dim Imin(1 To 4)
  Dim i As Long
  Dim Ia, Ib As Long

  For i = 1 To 4
    FinishPos = InStr(FromPos, T, "СНСМ      " & i)
    FinishPos = InStr(FinishPos, T, "IАсум")
    Ia = Int(Mid(T, FinishPos + 5, 10))
    FinishPos = InStr(FinishPos, T, "IВсум")
    Ib = Int(Mid(T, FinishPos + 5, 10))
    If (Ib < Ia) And (Ib > 0) Then Imin(i) = Ib Else Imin(i) = Ia
  Next
  Parse_Current_Line = Imin

End Function


Private Function Parse_Rezhim_Single(T, FromPos)
'
' Ищем что отключается в подрежиме, здесь обрабатывается только одно отключение
'

  Dim apos, bpos As Long
  Dim Prefix As String
  Dim pa, pb, pc As Long

  apos = InStr(FromPos, T, "Подрежим")
  bpos = InStr(FromPos, T, vbCrLf)
  If (apos > 0) And (bpos > apos) Then
    Prefix = Trim$(Mid(T, apos + 8, bpos - apos - 8))
  End If

  pa = InStr(FromPos, T, "(") + 1
  pb = InStr(FromPos, T, ")")
  pc = InStr(FromPos, T, "СНСМ")
  If (pa > 0) And (pb > 0) And (pc > 0) Then
    If pc > pb Then
      Parse_Rezhim_Single = "-" & Trim$(Mid(T, pa, pb - pa))
    Else
      Parse_Rezhim_Single = "КЗ на " & RootNode & ", ВСЕ ВКЛЮЧЕНО"
    End If
  End If
  Parse_Rezhim_Single = "[" & Prefix & "] " & Parse_Rezhim_Single

End Function
  

Private Sub Analiz_Sensitivity(Protokol As String)
'
' Разбор протокола АРМ ТКЗ и концентрация токов в отдельном листе
'

  ' Добавляем лист для результатов
  Dim objRez
  Dim TempSheetName As String
  Dim NewSheetName As String
  Dim i As Long
  Set objRez = ActiveWorkbook.Worksheets.Add
  objRez.Columns("A:A").ColumnWidth = 35#

  ' Подберем подходящее имя для нового листа
  TempSheetName = RootNode & " (" & Find_Node(RootNode) & ")"
  On Error Resume Next
    For i = 0 To 255          ' Врятли кто-то создаст столько вкладок :)
      If i = 0 Then
        NewSheetName = TempSheetName
      Else
        NewSheetName = TempSheetName & " #" & i
      End If
      If ActiveWorkbook.Worksheets(NewSheetName) Is Nothing Then Exit For
    Next
  On Error GoTo 0
  objRez.Name = NewSheetName

  objRez.Cells(1, 1).Value = "Узел " & RootNode & " (" & Find_Node(RootNode) & ")"
  objRez.Cells(2, 1).Value = "ТКЗ для чувств. пуск. и изб. органов"
  objRez.Cells(2, 2).Value = "КЗ(3)"
  objRez.Cells(2, 3).Value = "КЗ(2)"
  objRez.Cells(2, 4).Value = "КЗ(1+1)"
  objRez.Cells(2, 5).Value = "КЗ(1)"

  ' Пройдемся по подрежимам
  Dim list()
  Dim j As Long
  Dim StartPos As Long
  Dim R As String
  j = 0

  StartPos = InStr(Protokol, "Р Е З У Л Ь Т А Т Ы    Р А С Ч Е Т А")
  StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)

  ' Найдем что отключалось в этом подрежиме, в данном случае это только одна ветвь или элемент
  R = Parse_Rezhim_Single(Protokol, StartPos)

  Do
    ReDim Preserve list(j)
    list(j) = Array(R, Parse_Current_Line(Protokol, StartPos, StartPos))
    j = j + 1
    StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)
    If StartPos > 0 Then R = Parse_Rezhim_Single(Protokol, StartPos)
    DoEvents     ' Для того, чтобы работал выход по  Ctrl+C
  Loop While StartPos > 0

  ' Массив наименований режимов и соответствующих токов заполнен, переносим его на лист
  For i = LBound(list) To UBound(list)
    objRez.Cells(i + 3, 1).Value = list(i)(0)
    For j = 0 To 3
      objRez.Cells(i + 3, j + 2).Value = list(i)(1)(j + 1)
    Next j
  Next i

End Sub


Private Function Delete_Interm_Nodes(Without As String) As Long
'
' Удаление промежуточных узлов, т.е. тех у которых только два присоединения
' Функция не гарантирует полного удаления промежуточных п/станций, ее нужно вызывать
' до тех пор, пока не будет сделано никаких изменений (функция возвращает количество
' удаленных узлов
'

  Dim i, j, n As Long
  Dim Node As String
  Dim NodeBranch()
  Dim NodePosA, NodePosB As String
  Dim ContrNodeA, ContrNodeB As String
  Dim DestType, DestElement As Long

  j = 0
  For i = LBound(arrNode) To UBound(arrNode)
    Node = arrNode(i, 1)
    If Node <> Without Then
      NodeBranch = Find_Branch_By_Node(arrBranchCopy, Node)
      If Array_Exists(NodeBranch) Then
        n = UBound(NodeBranch)
        ' Промежуточные узлы
        If n = 1 Then
          ' В ветвях текущего узла найдем позицию текущего узла, чтобы выкинуть из ветвей текущий узел
          If arrBranchCopy(NodeBranch(0), 3) = Node Then NodePosA = 3
          If arrBranchCopy(NodeBranch(0), 4) = Node Then NodePosA = 4
          If arrBranchCopy(NodeBranch(1), 3) = Node Then NodePosB = 3
          If arrBranchCopy(NodeBranch(1), 4) = Node Then NodePosB = 4

          ' Найдем противоположные узлы (номера узлов)
          If arrBranchCopy(NodeBranch(0), 3) = Node Then
            ContrNodeA = arrBranchCopy(NodeBranch(0), 4)
          Else
            ContrNodeA = arrBranchCopy(NodeBranch(0), 3)
          End If
          If arrBranchCopy(NodeBranch(1), 3) = Node Then
            ContrNodeB = arrBranchCopy(NodeBranch(1), 4)
          Else
            ContrNodeB = arrBranchCopy(NodeBranch(1), 3)
          End If
          If ContrNodeA <> ContrNodeB Then
            ' Если с одной из сторон от промежуточного узла ветвь была 3 или 4 типа - результирующая
            ' ветвь должна тоже быть генератором или трансформатором
            DestType = 0
            ' Копируем тип ветви
            If (arrBranchCopy(NodeBranch(0), 1) > 1) Then DestType = arrBranchCopy(NodeBranch(0), 1)
            If (arrBranchCopy(NodeBranch(1), 1) > 1) Then DestType = arrBranchCopy(NodeBranch(1), 1)
            ' Номер элемента распространяется от Without (это RootNode), чтобы
            ' потом можно было легко определить присоединение от RootNode, идущее к питающему узлу
            DestElement = 0
            If (arrBranchCopy(NodeBranch(0), 3) = Without) Or (arrBranchCopy(NodeBranch(0), 4) = Without) Then DestElement = arrBranchCopy(NodeBranch(0), 5)
            If (arrBranchCopy(NodeBranch(1), 3) = Without) Or (arrBranchCopy(NodeBranch(1), 4) = Without) Then DestElement = arrBranchCopy(NodeBranch(1), 5)

            ' Изменения
            If (DestType = 0) Or (DestType = 3) Then
              arrBranchCopy(NodeBranch(0), 1) = DestType
              arrBranchCopy(NodeBranch(0), NodePosA) = ContrNodeB
              arrBranchCopy(NodeBranch(0), 5) = DestElement
              arrBranchCopy(NodeBranch(1), 3) = 0
              arrBranchCopy(NodeBranch(1), 4) = 0
              arrBranchCopy(NodeBranch(1), 5) = 0
              j = j + 1
            End If
          End If
        End If
        ' Тупики (узлы, которые входят только в одну ветвь)
        If n = 0 Then
          arrBranchCopy(NodeBranch(0), 3) = 0
          arrBranchCopy(NodeBranch(0), 4) = 0
        End If
      End If
    End If
    DoEvents     ' Для того, чтобы работал выход по  Ctrl+C
  Next
  Delete_Interm_Nodes = j

End Function


Private Sub Find_Power_Nodes()
'
' Поиск питающих узлов для RootNode
'

  Dim i, j, n, k, l As Long

  ' Удаляем сразу ветви с 101 типом (отключенный ШСВ)
  For i = LBound(arrBranchCopy) To UBound(arrBranchCopy)
    If arrBranchCopy(i, 1) = 101 Then
      arrBranchCopy(i, 3) = 0
      arrBranchCopy(i, 4) = 0
    End If
    ' Также затираем номера элементов в копиях таблицы ветвей
    arrBranchCopy(i, 5) = 0
    arrBranchCopy2(i, 5) = 0
  Next

  ' Пройдем по присоединениям RootNode и проставим отходящим от него ветвям уникальный номер
  ' элемента. При удалении промежуточных подстанций в Delete_Interm_Nodes()
  ' если один из узлов ветви = RootNode его номер элемента будет распространяться на вновь
  ' образованную ветвь

  For i = LBound(arrRootBranch) To UBound(arrRootBranch)
    arrBranchCopy(arrRootBranch(i), 5) = i + 1
    arrBranchCopy2(arrRootBranch(i), 5) = i + 1
  Next

  Do
    j = 0
    ' Удаляем промежуточные узлы и тупики (узлы только с одной ветвью)
    n = Delete_Interm_Nodes(RootNode)
    j = j + n

    ' Удаляем тупики в виде нейтралей и тр-ров на ноль (ТСН), но не генераторы
    For i = LBound(arrBranchCopy) To UBound(arrBranchCopy)
      If (arrBranchCopy(i, 1) <> 4) And (arrBranchCopy(i, 1) <> 3) Then
        If (arrBranchCopy(i, 3) = 0) And (arrBranchCopy(i, 4) <> 0) Then
          arrBranchCopy(i, 4) = 0
          j = j + 1
        End If
        If (arrBranchCopy(i, 4) = 0) And (arrBranchCopy(i, 3) <> 0) Then
          arrBranchCopy(i, 3) = 0
          j = j + 1
        End If
      End If
    Next
  Loop While j > 0

  ' Ищем противоположные узлы
  Dim list()
  Dim NodeBranch()
  Dim DestNode As String
  Dim R As Boolean

  j = 0
  NodeBranch = Find_Branch_By_Node(arrBranchCopy, RootNode)
  If Array_Exists(NodeBranch) Then
    n = UBound(NodeBranch)
    For i = LBound(NodeBranch) To n
      If arrBranchCopy(NodeBranch(i), 3) = RootNode Then
        DestNode = arrBranchCopy(NodeBranch(i), 4)
      Else
        DestNode = arrBranchCopy(NodeBranch(i), 3)
      End If
      ' Не добавляем дубликаты, которые могут появиться из за шутнирования СВ линиями (кольца)
      If (Array_Find(list, DestNode) = -1) And (DestNode <> "0") Then
        ReDim Preserve list(j)
        list(j) = DestNode
        j = j + 1
      End If
    Next
  End If

  Dim PowerNode As String
  Dim SecondNode As String
  Dim e As Long
  Dim Elem()
  Dim n_node As String
  Dim n_branch As String
  l = 0

  ' Подготовим номера питающих узлов и первую ветвь присоединения до них
  NodeBranch = Find_Branch_By_Node(arrBranchCopy2, RootNode)
  For i = LBound(list) To UBound(list)
    PowerNode = list(i)
    ' Найдем номер(а) элементов, в которые входят RootNode и PowerNode, если к питающему узлу идет не одна цепь
    ' этих элементов может быть несколько
    Elem = Find_Element_By_2Node(arrBranchCopy, RootNode, PowerNode)
    If Array_Exists(Elem) Then
      n = UBound(Elem)
      ' Найдем среди присоединений RootNode присоединение с элементом Elem
      For j = LBound(Elem) To n
        e = Elem(j)
        For k = LBound(NodeBranch) To UBound(NodeBranch)
          If arrBranchCopy2(NodeBranch(k), 5) = e Then
            If arrBranchCopy2(NodeBranch(k), 3) = RootNode Then
              SecondNode = arrBranchCopy2(NodeBranch(k), 4)
            Else
              SecondNode = arrBranchCopy2(NodeBranch(k), 3)
            End If
            'n_node = Trim$(Find_Node(list(i)))                                       ' ???
            'n_branch = Find_Branch_By_2Node(arrBranch, RootNode, SecondNode)        ' ???
            'n_branch = arrBranch(n_branch, 5)                                       ' ???
            'n_branch = Find_Element(n_branch)                                       ' ???
            ' Добавим питающий узел и первую ветвь до него в arrPowerNodes()
            ReDim Preserve arrPowerNodes(l)
            arrPowerNodes(l) = Array(PowerNode, RootNode, SecondNode)
            l = l + 1
          End If
        Next
      Next
    End If
  Next

End Sub


Private Function Get_Testing_Code() As String
'
' Подготовка приказа для оценки чувствительности ДЗШ в режиме опробования.
' Для каждого питающего узла создаем подрежим в котором отключаем все присоединения
' RootNode кроме ведущей к PowerNode ветви и на основе каждого такого режима создаем
' подрежимы в которых отключаем по одному присоединению питающего узла
'

  Dim R As String
  Dim i, j, k As Long

  R = _
    "*         ПРОВЕРКА ЧУВСТВИТЕЛЬНОСТИ ДЗШ ПРИ ОПРОБОВАНИИ, УЗЕЛ " & RootNode & " [" & Find_Node(RootNode) & "]" & vbCrLf & _
    "ВЕЛИЧИНА  IA IB IC" & vbCrLf & _
    "1-ПОЯС    " & RootNode & "      /* " & Find_Node(RootNode) & vbCrLf & _
    "СНСМ      1" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/ABC" & vbCrLf & _
    "СНСМ      2" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/AB" & vbCrLf & _
    "СНСМ      3" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/AB0" & vbCrLf & _
    "СНСМ      4" & vbCrLf & _
    "ЗАМ-ФАЗ   " & RootNode & "/A0" & vbCrLf

  Dim BaseRejim, Podrejim As Long
  Dim PowerNode As String
  Dim NodeA As String
  Dim NodeB As String
  Dim CrossNode As String
  Dim T As Long
  Dim branchNo As Long
  Dim ElemName As String
  Dim NodeBranch()
  Dim Elem As Long
  Dim rElem As Long
  Dim Collision As Boolean

  Podrejim = 1
  ' Проходим по всем присоединениям, указанным в списке (присоединения к питающему узлу)
  For i = LBound(arrPowerNodes) To UBound(arrPowerNodes)
    PowerNode = arrPowerNodes(i)(0)
    NodeA = arrPowerNodes(i)(1) ' RootNode
    NodeB = arrPowerNodes(i)(2) ' номер противоположного узла первой ветви присоединения к питающему узлу
    R = R & vbCrLf
    R = R & "ПОДРЕЖИМ  " & Podrejim & " /* " & PowerNode & " (" & Find_Node(PowerNode) & ")" & vbCrLf
    BaseRejim = Podrejim
    ' Запишем название базового режима, чтобы можно было это вписать в результирующий лист,
    ' из протокола эту информацию не достать
    ReDim Preserve arrBaseRejims(i)
    branchNo = Find_Branch_By_2Node(arrBranch, NodeA, NodeB)
    ElemName = Find_Element(Int(arrBranch(branchNo, 5)))
    arrBaseRejims(i) = Array(BaseRejim, Find_Node(PowerNode) & " (" & ElemName & ")")
      
    ' Пройдемся по всем присоединениям RootNode
    For j = LBound(arrRootBranch) To UBound(arrRootBranch)
      ' Для ветвей, отходящих от RootNode найдем номер противоположного узла
      If arrBranch(arrRootBranch(j), 3) = RootNode Then
        CrossNode = arrBranch(arrRootBranch(j), 4)
      Else
        CrossNode = arrBranch(arrRootBranch(j), 3)
      End If
      ' Отключаем все присоединения, кроме ведущего к питающему узлу
      If CrossNode = NodeB Then
        ' Для справки выводим коммутацию в комментарии
        R = R & "*"
      End If
      R = R & "ОТКЛ      *" & RootNode & "-" & CrossNode & _
      "      /* Элемент " & arrBranch(arrRootBranch(j), 5) & " (" & _
      Find_Element(Int(arrBranch(arrRootBranch(j), 5))) & "), Ветвь (" & _
      Find_Node(Trim$(arrBranch(Int(arrRootBranch(j)), 3))) & " - " & Find_Node(Trim$(arrBranch(Int(arrRootBranch(j)), 4))) & _
      ")" & vbCrLf
    Next
    
    ' Найдем все присоединения питающего узла и отключим каждое в отдельном подрежиме, основанном на BaseRejim
    NodeBranch = Find_Branch_By_Node(arrBranch, PowerNode)
    For j = LBound(NodeBranch) To UBound(NodeBranch)
      T = arrBranch(NodeBranch(j), 1)
      
      ' Issue#2: Проверим, что ветвь, отходящая от питающего узла, которую мы хотим отключить,
      ' не ведет к RootNode (все ветви RootNode кроме одной отключены в базовом режиме)
      Elem = arrBranch(NodeBranch(j), 5)  ' Номер элемента той ветви от PowerNode, которую хотим отключить
      Collision = False
      For k = LBound(arrRootBranch) To UBound(arrRootBranch)
        rElem = arrBranch(arrRootBranch(k), 5)
        If Elem = rElem Then
          Collision = True
          Exit For
        End If
      Next
      ' Issue#8: Незачем отключать отключенные ШСВ
      If (T <> 101) And Not Collision Then
        Podrejim = Podrejim + 1
        R = R & "ПОДРЕЖИМ  " & Podrejim & " " & BaseRejim & vbCrLf
        NodeA = arrBranch(NodeBranch(j), 3)
        NodeB = arrBranch(NodeBranch(j), 4)
               
        If NodeA = PowerNode Then
          CrossNode = NodeB
        Else
          CrossNode = NodeA
        End If
      
        If Elem = 0 Then
          R = R & "ОТКЛ      *" & PowerNode & "-" & CrossNode & _
          " /* " & Find_Node(PowerNode) & " - " & Find_Node(CrossNode) & vbCrLf
        Else
          R = R & "ЭЛЕМЕНТ   " & Elem & _
          " /* " & Find_Element(Elem) & vbCrLf
        End If
      End If
    Next
    
    Podrejim = Podrejim + 1
  Next

Get_Testing_Code = R

End Function


Private Function Is_Sub_Rejim(Protokol, CurPos)
'
' Проверка стоим ли сейчас в подрежиме, основанном на другом режиме
' (ремонт или отключение на питающем узле)
'

  Dim apos, bpos, cpos As Long

  Is_Sub_Rejim = False
  apos = InStr(CurPos, Protokol, "Подрежим  ")

  If apos > 0 Then
    bpos = InStr(apos + 10, Protokol, " ")
    cpos = InStr(apos + 10, Protokol, vbCrLf)
    If cpos > bpos Then Is_Sub_Rejim = True
  End If

End Function


Private Function Find_BaseRejim_Name(RejimNo)
'
' Ищем наименование режима из базовых режимов в arrBaseRejims()
'

  Dim i As Long

  Find_BaseRejim_Name = ""
  For i = LBound(arrBaseRejims) To UBound(arrBaseRejims)
    If arrBaseRejims(i)(0) = Int(RejimNo) Then
      Find_BaseRejim_Name = arrBaseRejims(i)(1)
      Exit For
    End If
  Next

End Function


Private Function Get_Rejim_Name(Protokol, CurPos)
'
' Если стоим на базовом подрежиме - берем его наименование,
' если это субрежим - берем наименование отключаемого элемента
'

  Dim apos, bpos As Long
  Dim Prefix As String

  ' Найдем номер режима, если это подрежим - номер состоит из двух чисел,
  ' если основной - одного
  apos = InStr(CurPos, Protokol, "Подрежим")
  bpos = InStr(apos, Protokol, vbCrLf)
  If (apos > 0) And (bpos > apos) Then
    Prefix = Trim$(Mid(Protokol, apos + 8, bpos - apos - 8))
  End If

  apos = 0
  bpos = 0

  If Is_Sub_Rejim(Protokol, CurPos) Then
    apos = InStr(CurPos, Protokol, "(")
    bpos = InStr(apos, Protokol, ")")
    If (apos > 0) And (bpos > apos) Then
      Get_Rejim_Name = "[" & Prefix & "] +Откл " & Trim$(Mid(Protokol, apos + 1, bpos - apos - 1))
    End If
  Else
    apos = InStr(CurPos + 10, Protokol, vbCrLf)
    Get_Rejim_Name = "[" & Prefix & "] " & Find_BaseRejim_Name(Trim$(Mid(Protokol, CurPos + 10, apos - (CurPos + 10))))
  End If
  
End Function


Private Sub Analiz_Testing(Protokol As String)
'
' Парсинг протокола по опробованию
'

  Dim objWorkbook, objRez
  Dim s, i As Long

  ' Результаты быдем выводить в тот же лист, что и результаты по проверке чувствительности
  ' в минимальном режиме
  Set objWorkbook = ActiveWorkbook
  Set objRez = objWorkbook.ActiveSheet
  s = objRez.UsedRange.Rows.Count + 2

  objRez.Cells(s, 1).Value = "ТКЗ для опробования"
  objRez.Cells(s, 2).Value = "КЗ(3)"
  objRez.Cells(s, 3).Value = "КЗ(2)"
  objRez.Cells(s, 4).Value = "КЗ(1+1)"
  objRez.Cells(s, 5).Value = "КЗ(1)"

  ' Пройдемся по подрежимам
  Dim list()
  Dim j As Long
  Dim StartPos As Long
  Dim RejimName As String
  Dim Line
  j = 0

  StartPos = InStr(Protokol, "Р Е З У Л Ь Т А Т Ы    Р А С Ч Е Т А")
  StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)

  Do
    ReDim Preserve list(j)
    RejimName = Get_Rejim_Name(Protokol, StartPos)
    Line = Parse_Current_Line(Protokol, StartPos, StartPos)
      
    list(j) = Array(RejimName, Line)
    j = j + 1
    StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)
    DoEvents     ' Для того, чтобы работал выход по  Ctrl+C
  Loop While StartPos > 0

  ' Массив наименований режимов и соответствующих токов заполнен, переносим его на лист
  objRez.Cells(s, 1).Value = "ТКЗ для режима опробования"
  For i = LBound(list) To UBound(list)
    objRez.Cells(s + i + 1, 1).Value = list(i)(0)
    For j = 0 To 3
      objRez.Cells(s + i + 1, j + 2).Value = list(i)(1)(j + 1)
    Next
  Next

End Sub

'######################################################################################[Главный метод макроса]

Public Sub Raschet_DZSH()
'
' Main
'

  Dim MainFormHandle As Long
  Dim CommandFormHandle As Long
  Dim CommandRichEdit As Long
  Dim ProtokolHandle As Long
  Dim ProtokolMemo As Long
  Dim CommandsText As String
  Dim ProtokolText As String
  Dim ProcessWnd As Long
  Dim ManualMode As Boolean
   
  ' Нажата ли Ctrl ?
  ManualMode = CBool(GetKeyState(&H11) < 0)

  ' Ищем окно ТКЗ-2000, если его нет выводим сообщение и завершаемся
  MainFormHandle = Find_TKZ_Window_Handle("TForm1")
  If MainFormHandle = 0 Then
    MsgBox "Окно ТКЗ-2000 не найдено, приложение должно быть запущено. Кроме этого должна быть загружена сеть для расчета.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' Инициализация, выбираем данные из листов
  Call Initialize(Exe_Name_by_Window_Handle(MainFormHandle))

  ' Получаем номер узла (вообще-то узел может быть не только числовой)
  Dim Answer
  Answer = InputBox("Номер узла (рассчитываемые шины)?", "RootNode", 0)
  If Trim$(Answer) = "" Then Exit Sub
  RootNode = Trim$(Answer)
  If Not Node_Exists(RootNode) Then
    MsgBox "Узел " & RootNode & " не найден в сети, дальнейшая работа невозможна.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' Отобразим маленький диалог
  ProcessWnd = Show_Process_Window("Процесс идет, ждите...")

  ' Не проверяя в каком режиме работает программа (приказы или диалоговый расчет) выполним пункт меню
  Call SendMessage(MainFormHandle, WM_COMMAND, TKZ_MENU_ADV_MODE, 0&)     ' "Расширенный формат задания для расчета..."

  ' Найдем окно диалога задания приказов
  CommandFormHandle = Find_TKZ_Window_Handle("TFormZD2")

  ' Откроем окно протокола и очистим его
  Call SendMessage(CommandFormHandle, WM_COMMAND, TKZ_MENU_OPEN_LOG, 0&) ' "Открыть протокол..."
  Call SendMessage(CommandFormHandle, WM_COMMAND, TKZ_MENU_CLEAR_CODE, 0&) ' "Очистить задание"
  ProtokolHandle = Find_TKZ_Window_Handle("TForm3")
  ProtokolMemo = Find_SubClass_Recurce(ProtokolHandle, "TMemo")
  Call SendMessage(ProtokolHandle, WM_COMMAND, TKZ_MENU_CLEAR_LOG, 0&)    ' "Очистить протокол"

  ' <<< Проверка чувствительности пусковых/избирательных органов

  ' Готовим приказ для проверки чувствительности и копируем его в окно приказов ТКЗ-2000
  CommandsText = Get_Sensitivity_Code()
  CommandRichEdit = Find_SubClass_Recurce(CommandFormHandle, "TRichEdit")
  Window_Set_Text CommandRichEdit, CommandsText

  If ManualMode Then MsgBox _
  "В окно ввода приказов ТКЗ-2000 вставлен приказ ПРОВЕРКИ ЧУВСТВИТЕЛЬНОСТИ, подготовленный макросом, если Вы хотите изменить его," & _
  "сделайте это." & vbCrLf & _
  "Для продолжения работы макроса нажмите Ok"

  ' Делаем расчет с эквивалентированием - это значительно быстрее
  Call SendMessage(CommandFormHandle, WM_COMMAND, TKZ_MENU_CALC, 0&) ' "Расчет с эквиваленированием"

  ' Заберем результат для анализа
  ProtokolText = Window_Get_Text(ProtokolMemo)

  ' Анализ протокола расчета
  Analiz_Sensitivity ProtokolText

  ' Предложить пользователю сохранить расширенный протокол (вначале добавлен исходный приказ с комментариями)
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText

  Dim setName As String
  Dim filePRT As String
  Dim FrFi As Integer
    
  setName = Mid(ActiveWorkbook.Name, 1, InStrRev(ActiveWorkbook.Name, ".") - 1)
  filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\Чувствительность " & RootNode & " узел " & "(" & setName & ").prt", "Файлы протокола АРМ (*.prt), *.prt")
  If filePRT <> "False" Then
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, ProtokolText
    Close FrFi
  End If

  ' <<< Проверка чувствительности в режиме опробования

  ' Готовим приказ для проверки опробования
  Call Find_Power_Nodes

  CommandsText = Get_Testing_Code()
  
  ' Очистим протокол и окно приказов, скопируем приказ и выполним его
  Call SendMessage(ProtokolHandle, WM_COMMAND, TKZ_MENU_CLEAR_LOG, 0&)    ' "Очистить протокол"
  Call SendMessage(CommandFormHandle, WM_COMMAND, TKZ_MENU_CLEAR_CODE, 0&) ' "Очистить задание"

  Window_Set_Text CommandRichEdit, CommandsText
  
  If ManualMode Then MsgBox _
  "В окно ввода приказов ТКЗ-2000 вставлен приказ ПРОВЕРКИ ЧУВСТВИТЕЛЬНОСТИ ПРИ ОПРОБОВАНИИ, подготовленный макросом, если Вы хотите изменить его," & _
  "сделайте это." & vbCrLf & _
  "Для продолжения работы макроса нажмите Ok"
  
  Call SendMessage(CommandFormHandle, WM_COMMAND, TKZ_MENU_CALC, 0&) ' "Расчет с эквиваленированием"
  
  ' Заберем результат для анализа
  ProtokolText = Window_Get_Text(ProtokolMemo)
  
  ' Анализируем результаты расчета
  Analiz_Testing (ProtokolText)

  ' Уберем диалог процесса
  If ProcessWnd > 0 Then DestroyWindow (ProcessWnd)

  ' Предложить пользователю сохранить расширенный протокол с комментариями
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText
  filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\Опробование " & RootNode & " узел " & "(" & setName & ").prt", "Файлы протокола АРМ (*.prt), *.prt")
  If filePRT <> "False" Then
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, ProtokolText
    Close FrFi
  End If

End Sub
