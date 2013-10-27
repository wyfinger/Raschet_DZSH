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
' ��������� ����� EXE ����� �� Handle ����
'
  
  Exe_Name_by_Window_Handle = ""
  
  Dim prcID As Long  ' ����� ��������
  Dim prc As Long    ' ���������� ������� � ������ ��������
  Dim wt As String
  wt = Space(1024)
  
  If GetWindowThreadProcessId(wnd, prcID) <> 0 Then         ' ���� �������
    prc = OpenProcess(PROCESS_ALL_ACCESS, False, prcID)     ' ���������
    On Error GoTo Finally
    If GetModuleFileNameEx(prc, 0, wt, 1024) <> 0 Then      ' �������� ��� EXE
      Exe_Name_by_Window_Handle = Trim(wt)
    End If
Finally:
    CloseHandle prc
  End If

End Function
    
Private Function ExtractFileName(FileName As String) As String
'
' ��������� ����� ����� �� ������� ���� � �����
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

Private Function Get_Class_Name(ByVal wnd As Long) As String

  Dim ClassName As String
  Dim ClassLen As Long
  ClassName = Space(256)
  ClassLen = GetClassName(wnd, ClassName, 256)
  Get_Class_Name = Mid(ClassName, 1, ClassLen)
  
End Function

Public Function Find_TKZ_Window_Enum_Proc(ByVal wnd As Long, ByVal lParam As Long) As Boolean
'
' ������� ���������� ������������ ���� ���� � ������� ���
'

  Find_TKZ_Window_Enum_Proc = True

  Dim ExeName As String
  Dim WndClass As String
  Dim TForm1_Handle As Long
  
  If IsWindowVisible(wnd) And (GetParent(wnd) = 0) Then
  
    ExeName = UCase(ExtractFileName(Exe_Name_by_Window_Handle(wnd)))
    If (ExeName = "TKZ2000.EXE") Then
      WndClass = Get_Class_Name(wnd)
      If WndClass = "TForm1" Then
        G_TKZ_Handle = wnd
        Find_TKZ_Window_Enum_Proc = False
      End If
    End If
    
  End If
 
End Function

Private Function Find_SubClass_Recurce(hwnd As Long, sClassName As String, Optional iPos As Integer = 1) As Long
'
' ��� ������� ���������� ���������� ��� �������� ���� hWnd, ������ ����c ���� � sClassName
' ���� ������ iPos ������� ���������� iPos ��������� ������������� ������
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

Private Function Find_TKZ_Window_Handle() As Boolean
'
' ����� �������� ���� ��������� ���-2000
'

  G_TKZ_Handle = 0
  Call EnumWindows(AddressOf Find_TKZ_Window_Enum_Proc, 0)
  Find_TKZ_Window_Handle = G_TKZ_Handle <> 0

End Function

'#############################################################################################################

Public Sub Raschet_DZSH()
'
' Main
'

' ���� ���� ���-2000, ���� ��� ��� ������� ��������� � �����������
If Not Find_TKZ_Window_Handle() Then
 MsgBox "���� ���-2000 �� �������, ���������� ������ ���� ��������. ����� ����� ������ ���� ��������� ���� ��� �������.", vbExclamation + vbOKOnly
 Exit Sub
End If

' �� �������� � ����� ������ �������� ��������� (������� ��� ���������� ������) �������� ����� ����

End Sub
