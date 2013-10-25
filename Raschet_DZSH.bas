Attribute VB_Name = "Raschet_DZSH"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal uCmd As Long) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAcess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "PSAPI.DLL" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const WM_COMMAND = &H111
Private Const WM_PASTE = &H302
Private Const WM_CUT = &H300
Private Const GW_CHILD = &H5
Private Const GW_OWNER = &H4
Private Const GW_HWNDNEXT = &H2
Private Const EM_SETSEL = &HB1
Private Const EM_REPLACESEL = &HC2
Private Const WM_SETTEXT = &HC
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const MAX_PATH As Long = 260

Dim G_TKZ_Handle As Long

Dim rc&, rs&, sc&

Private Function Exe_Name_by_Window_Handle(wnd As Long) As String
'
' Получение имени EXE файла по Handle окна
'
  
  Exe_Name_by_Window_Handle = ""
  
  Dim prcID As Long  ' Номер процесса
  Dim prc As Long    ' Дискриптор доступа к данным процесса
  Dim wt As String
  wt = Space(1024)
  
  If GetWindowThreadProcessId(wnd, prcID) <> 0 Then         ' Ищем процесс
    prc = OpenProcess(PROCESS_ALL_ACCESS, False, prcID)     ' Открываем
    On Error GoTo Finally
    If GetModuleFileNameEx(prc, 0, wt, 1024) <> 0 Then      ' Получаем имя EXE
      Exe_Name_by_Window_Handle = Trim(wt)
    End If
Finally:
    CloseHandle prc
  End If

End Function
    
Private Function ExtractFileName(FileName As String) As String
'
' Получение имени файла из полного пути и имени
'

Dim SlashPos As Long

If FileName <> "" Then
  SlashPos = InStrRev(FileName, "\")
  If SlashPos > 0 Then
    SlashPos = SlashPos + 1
    ExtractFileName = Mid(FileName, SlashPos, Len(FileName) - SlashPos)
  End If
End If

End Function
   
Public Function Find_TKZ_Window_Enum_Proc(ByVal wnd As Long, ByVal lParam As Long) As Boolean
'
' Функция обработчик перечисления всех окон в системе для
'

  Find_TKZ_Window_Enum_Proc = True

  Dim ExeName As String
  
  If IsWindowVisible(wnd) And (GetParent(wnd) = 0) Then
  
    ExeName = UCase(ExtractFileName(Exe_Name_by_Window_Handle(wnd)))
    If ExeName = "TKZ2000.EXE" Then
      G_TKZ_Handle = wnd
      Find_TKZ_Window_Enum_Proc = False
    End If
    
  End If
 
 
End Function

Public Function Find_TKZ_Window_Handle() As Long
'
' Поиск главного окна программы ТКЗ-2000
'

  G_TKZ_Handle = 0
  Call EnumWindows(AddressOf Find_TKZ_Window_Enum_Proc, 0)

  MsgBox G_TKZ_Handle

End Function
