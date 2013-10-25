VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestForm 
   Caption         =   "UserForm1"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   OleObjectBlob   =   "TestForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
SlashPos = InStrRev(FileName, "\")
If SlashPos > 0 Then
  ExtractFileName = Mid(FileName, SlashPos, Len(FileName) - SlashPos)
End

End Function
   
'Public Function Find_TKZ_Window_Enum_Proc(ByVal wnd As Long, ByVal lParam As Long) As Boolean
'
' Функция обработчик перечисления всех окон в системе для
'

'  Find_TKZ_Window_Enum_Proc = True

'  Dim ExeName As String
  
'  If IsWindowVisible(wnd) And (GetParent(wnd) = 0) Then
  
 '   ExeName = UCase(ExtractFileName(Exe_Name_by_Window_Handle(wnd)))
 '   If ExeName = "TKZ2000.EXE" Then
 '     G_TKZ_Handle = wnd
 '     Find_TKZ_Window_Enum_Proc = False
 '   End If
    
 ' End If
 
 
'End Function

'Private Function Find_TKZ_Window_Handle() As Long
'
' Поиск главного окна программы ТКЗ-2000
'

'  G_TKZ_Handle = 0
'  Call EnumWindows(AddressOf Find_TKZ_Window_Enum_Proc, 0)

'  MsgBox G_TKZ_Handle

'End Function


Private Function GetWindowClass(hwnd As Long) As String
'
' Получение имени класса окна по его хандлу
'

Dim class_name_lengh As Integer
Dim class_name As String

class_name = Space(255)
class_name_lengh = GetClassName(hwnd, class_name, 255)
GetWindowClass = Mid(class_name, 1, class_name_lengh)

End Function


Private Function Find_SubClass_Recurce(hwnd As Long, sClassName As String, Optional iPos As Integer = 1) As Long
'
' Эта функция рекурсивно перебирает все дочерние окна hWnd, сверяя класc окна с sClassName
' Если указан iPos функция возвращает iPos вхождение интересующего класса
'

Dim window_class As String

window_class = GetWindowClass(hwnd)
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

' Копируем приказ в буфер
Dim d As New DataObject
d.SetText (sText)
d.PutInClipboard

PostMessage hwnd, EM_SETSEL, 0, -1
PostMessage hwnd, WM_PASTE, 0, 0
  
End Function


Private Function Window_Get_Text(hwnd As Long)


PostMessage hwnd, EM_SETSEL, 0, -1
PostMessage hwnd, WM_CUT, 0, 0

' Копируем приказ в буфер
Dim d As New DataObject
d.GetFromClipboard
Window_Get_Text = d.GetText
  
End Function



Private Sub CommandButton1_Click()

Raschet_DZSH.Find_TKZ_Window_Handle

Exit Sub


Dim TestHandle As Long
TestHandle = InputBox("Окно")
MsgBox Exe_Name_by_Window_Handle(TestHandle)
Exit Sub

rc& = FindWindow("TFormZD2", "")
TextBox1.Text = TextBox1.Text & vbCrLf & rc&

End Sub

Private Sub CommandButton5_Click()

rs& = GetMenuItemID(GetSubMenu(GetMenu(rc&), 1), 2)
Call PostMessage(rc&, WM_COMMAND, rs&, 0&)

End Sub

Private Sub CommandButton6_Click()

Dim rich_edit As Long
rich_edit = Find_SubClass_Recurce(rc&, "TRichEdit")

Window_Set_Text rich_edit, "Проверка"


End Sub
