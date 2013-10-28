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
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const WM_COMMAND = &H111
Private Const WM_PASTE = &H302
Private Const WM_COPY = &H301
Private Const WM_CUT = &H300
Private Const GW_CHILD = &H5
Private Const GW_OWNER = &H4
Private Const GW_HWNDNEXT = &H2
Private Const EM_SETSEL = &HB1
Private Const EM_REPLACESEL = &HC2
Private Const WM_SETTEXT = &HC
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const MAX_PATH As Long = 260

Dim GT_Class As String       ' ��������� ���������� ������ ��� �������� � Enum-�������
Dim GT_Result As Long
                                                                           
Dim RootNode As Integer      ' ����� ������� ���� :)
Dim arrRootBranch()          ' ������ ������ �������� ����, ����������� � Get_Sensitivity_Code()

Dim arrBranch()
Dim arrBranchCopy()
Dim arrBranchCopy2()
Dim arrNode()
Dim arrElement()
Dim Initialized
Dim arrTrueBrach()           ' ������ ������������� ����, ����� �������������
Dim arrBaseRejims()          ' �����, �������� ��� ������� �������, ����� ��� �������� ��������� �� �����������
Const vbTab = "   "          ' ���� ��������� ��� ���������� �� ��������� �������� � �����
                                                                           
'##########################################################################[ ������� �������������� � ������ ]

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


Private Function ExtractFilePath(FileName As String) As String
'
' ��������� ���� � �����, ���������� ���� (��� ������������ �����)
'

  Dim SlashPos As Long

  If FileName <> "" Then
    SlashPos = InStrRev(FileName, "\")
    If SlashPos > 0 Then
      SlashPos = SlashPos + 1
      ExtractFileName = Mid(FileName, 1, SlashPos - 1)
    End If
  End If

End Function


Private Function Get_Class_Name(ByVal wnd As Long) As String
'
' ������� ��� ����������� API ��������
'

  Dim ClassName As String
  Dim ClassLen As Long
  ClassName = Space(256)
  ClassLen = GetClassName(wnd, ClassName, 256)
  Get_Class_Name = Mid(ClassName, 1, ClassLen)
  
End Function


Public Function Find_Window_Enum_Proc(ByVal wnd As Long, ByVal lParam As Long) As Boolean
'
' ������� ���������� ������������ ���� ���� � ������� ��� ������ �������� ���� ���-2000
'

  Find_Window_Enum_Proc = True

  Dim ExeName As String
  Dim WndClass As String
  
  If IsWindowVisible(wnd) And (GetParent(wnd) = 0) Then
  
    ExeName = UCase(ExtractFileName(Exe_Name_by_Window_Handle(wnd)))
    If (ExeName = "TKZ2000.EXE") Then
      WndClass = Get_Class_Name(wnd)
      If WndClass = GT_Class Then
        GT_Result = wnd
        Find_Window_Enum_Proc = False
      End If
    End If
    
  End If
 
End Function


Private Function Find_TKZ_Window_Handle(ByVal class As String) As Long
'
' ����� �������� ���� ��������� ���-2000
'

  GT_Result = 0
  GT_Class = class
  Call EnumWindows(AddressOf Find_Window_Enum_Proc, 0)
  Find_TKZ_Window_Handle = GT_Result

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


Private Function Window_Set_Text(hwnd As Long, sText As String)

' �������� ������ � �����
Dim d As New DataObject
d.SetText (sText)
d.PutInClipboard

SendMessage hwnd, EM_SETSEL, 0, -1
SendMessage hwnd, WM_PASTE, 0, 0
  
End Function

Private Function Window_Get_Text(hwnd As Long)


SendMessage hwnd, EM_SETSEL, 0, -1
SendMessage hwnd, WM_CUT, 0, 0

' �������� ������ � �����
Dim d As New DataObject
d.GetFromClipboard
Window_Get_Text = d.GetText
  
End Function

'#################################################################################################[������ ���]

Private Sub Initialize()
'
' ���������� - ������� ������ � ����� � ����������
'

  Dim wshBranch
  Dim wshNode
  Dim wshElement


  ' 1. ����� - ����� ������ ���� �������� �������� '������� ������'
  Set wshBranch = ActiveWorkbook.Worksheets("������� ������")
  arrBranch = wshBranch.Range("A3:K" & wshBranch.UsedRange.Rows.Count).Value2
  arrBranchCopy = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2
  arrBranchCopy2 = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2

  ' 2. ���� - ����� ������ ��� ������� �������� '����.�����'
  Set wshNode = ActiveWorkbook.Worksheets("����.�����")
  arrNode = wshNode.Range("A3:E" & wshNode.UsedRange.Rows.Count).Value2
  
  ' 3. �������� - ����� ������ ��� ������� �������� '����.���������'
  Set wshElement = ActiveWorkbook.Worksheets("����.���������")
  arrElement = wshElement.Range("A3:B" & wshElement.UsedRange.Rows.Count).Value2

  ' ������� ������
  Set wshBranch = Nothing
  Set wshNode = Nothing
  Set wshElement = Nothing

  Initialized = True

End Sub


Private Function Find_Branch_By_Node(BranchArray, Node)
'
' ���� ��� �����, � ������� ������ ������� ����
'

  Dim rez()
  Dim i As Integer
  Dim j As Integer
  j = 0
  For i = 1 To UBound(BranchArray)
    If (Int(BranchArray(i, 3)) = Node) Or (Int(BranchArray(i, 4)) = Node) Then
      ReDim Preserve rez(j)
      rez(j) = i
      j = j + 1
    End If
  Next
  Find_Branch_By_Node = rez

End Function


Private Function Find_Branch_By_2Node(BranchArray, Node1, Node2)
'
' ���� ���� ����� �� ��������� ���� �����, �� ����������, ���������� ����� �����
'

  rez = -1
  Dim i As Integer
  For i = 1 To UBound(BranchArray)
    If ((Int(BranchArray(i, 3)) = Node1) And (Int(BranchArray(i, 4)) = Node2)) Or _
       ((Int(BranchArray(i, 3)) = Node2) And (Int(BranchArray(i, 4)) = Node1)) Then
      Find_Branch_By_2Node = i
      Exit For
    End If
  Next

End Function


Private Function Find_Node(Node)
'
' ���� ������������ ���� �� ������
'

  Dim i As Integer
  For i = 1 To UBound(arrNode)
    If Int(arrNode(i, 1)) = Node Then
      Find_Node = Trim(arrNode(i, 2))
      Exit For
    End If
  Next

End Function


Private Function Find_Node_Index(Node)
'
' ���� ������ ���� �� ������
'

  Dim i As Integer
  For i = 1 To UBound(arrNode)
    If Int(arrNode(i, 1)) = Node Then
      Find_Node_Index = i
      Exit For
    End If
  Next

End Function


Private Function Node_Exists(Node)
'
' �������� ������������� ����
'

  Dim i As Integer
  Node_Exists = False
  For i = 1 To UBound(arrNode)
    If Int(arrNode(i, 1)) = Node Then
      Node_Exists = True
      Exit For
    End If
  Next

End Function


Private Function Find_Element(Element)
'
' ���� ������������ �������� �� ������
'

  Dim i As Integer
  For i = 1 To UBound(arrElement)
    If Int(arrElement(i, 1)) = Element Then
      Find_Element = Trim(arrElement(i, 2))
      Exit For
    End If
  Next

End Function


Private Function Find_Branch_Index(Node1, Node2)
'
' ���� ������ ����� �� ����� ����� � ������, ������� ����� �� ����� ��������
'

  Dim i As Integer
  Find_Branch_Index = 0
  For i = 1 To UBound(arrBranch)
    If ((arrBranch(i, 3) = Node1) And (arrBranch(i, 4) = Node2)) Or _
       ((arrBranch(i, 3) = Node2) And (arrBranch(i, 4) = Node1)) Then
       Find_Branch_Index = i
       Exit For
     End If
  Next

End Function


Private Function Get_Sensitivity_Code() As String
'
' ���������� ������� ��� ������ ���������������� ���
' ���� ��� ������������� �������� ����, ������ �� �� ���� � ���������� ������,
' � � ���������� ��������� �� ������ �������������
'

  Dim R As String

  R = _
"*         �������� ���������������� ���, ���� " & RootNode & " [" & Find_Node(RootNode) & "]" & vbCrLf & _
"��������  IA IB IC" & vbCrLf & _
"1-����    " & RootNode & vbTab & "/* " & Find_Node(RootNode) & vbCrLf & _
"����      1" & vbCrLf & _
"���-���   " & RootNode & "/ABC" & vbCrLf & _
"����      2" & vbCrLf & _
"���-���   " & RootNode & "/AB" & vbCrLf & _
"����      3" & vbCrLf & _
"���-���   " & RootNode & "/AB0" & vbCrLf & _
"����      4" & vbCrLf & _
"���-���   " & RootNode & "/A0" & vbCrLf & _
"��������  1    /* ��� ��������" & vbCrLf

  Dim CrossNode As Long
  Dim ElemNo As Long
  Dim ElemName As String
  Dim NewStr As String
  Dim i As Long

  ' ��� ������� �� ������������� RootNode ������� ��������, ��� ��������� �������������
  arrRootBranch = Find_Branch_By_Node(arrBranch, RootNode)
  For i = 0 To UBound(arrRootBranch)
  
    If arrBranch(arrRootBranch(i), 3) = RootNode Then
      CrossNode = arrBranch(arrRootBranch(i), 4)
    Else
      CrossNode = arrBranch(arrRootBranch(i), 3)
    End If
    ElemNo = arrBranch(arrRootBranch(i), 5)
    ElemName = Find_Element(ElemNo)
    If (CrossNode = 0) Or (ElemNo = 0) Then               ' ��������� �����
      NewStr = "��������  " & (i + 2) & vbCrLf & _
      "����      0 *" & RootNode & "-" & CrossNode & "      /* �������� ?"
    Else                                                  ' ��������� �������
      NewStr = "��������  " & (i + 2) & vbCrLf & _
      "�������   " & ElemNo & "      /* " & ElemName
    End If
  
    R = R & NewStr & vbCrLf
  
  Next

  Get_Sensitivity_Code = R

End Function


Private Function Parse_Current_Line(T, FromPos, ByRef FinishPos)
'
' ������� ���� ��������� ���� �� � ������������ � ������ ��������� � ������� �
' FromPos, � �������� FinishPos ������������ �������, �� ������� ����������� ������
'

  Dim Imin(1 To 4)
  Dim i As Long
  Dim Ia, Ib As Long
  
  For i = 1 To 4
    FinishPos = InStr(FromPos, T, "����      " & i)
    FinishPos = InStr(FinishPos, T, "I����")
    Ia = Int(Mid(T, FinishPos + 5, 10))
    FinishPos = InStr(FinishPos, T, "I����")
    Ib = Int(Mid(T, FinishPos + 5, 10))
    If (Ib < Ia) And (Ib > 0) Then Imin(i) = Ib Else Imin(i) = Ia
  Next
  Parse_Current_Line = Imin

End Function


Private Function Parse_Rezhim_Single(T, FromPos)
'
' ���� ��� ����������� � ���������, ����� �������������� ������ ���� ����������
'

  Dim apos, bpos As Long
  Dim Prefix As String
  Dim pa, pb, pc As Long

  apos = InStr(FromPos, T, "��������")
  bpos = InStr(FromPos, T, vbCrLf)
  If (apos > 0) And (bpos > apos) Then
    Prefix = Trim(Mid(T, apos + 8, bpos - apos - 8))
  End If

  pa = InStr(FromPos, T, "(") + 1
  pb = InStr(FromPos, T, ")")
  pc = InStr(FromPos, T, "����")
  If (pa > 0) And (pb > 0) And (pc > 0) Then
    If pc > pb Then
      Parse_Rezhim_Single = "-" & Trim(Mid(T, pa, pb - pa))
    Else
      Parse_Rezhim_Single = "�� �� " & RootNode & ", ��� ��������"
    End If
  End If
  Parse_Rezhim_Single = "[" & Prefix & "] " & Parse_Rezhim_Single

End Function

                                                                                       
'######################################################################################[������� ����� �������]

Private Sub Analiz_Sensitivity(Protokol As String)
'
' ������ ��������� ��� ��� � ������������ ����� � ��������� �����
'

  ' ��������� ���� ��� �����������
  Dim objRez
  Dim TempSheetName As String
  Dim NewSheetName As String
  Dim i As Long
  Set objRez = ActiveWorkbook.Worksheets.Add
  objRez.Columns("A:A").ColumnWidth = 35#
  
  ' �������� ���������� ��� ��� ������ �����
  TempSheetName = RootNode & " (" & Find_Node(RootNode) & ")"
  For i = 0 To 25
    If i = 0 Then
      NewSheetName = TempSheetName
    Else
      NewSheetName = TempSheetName & " #" & i
    End If
    On Error Resume Next
    If ActiveWorkbook.Worksheets(NewSheetName) Is Nothing Then Exit For
  Next
  On Error GoTo 0
  objRez.Name = NewSheetName
  
  objRez.Cells(1, 1).Value = "���� " & RootNode & " (" & Find_Node(RootNode) & ")"
  objRez.Cells(2, 1).Value = "��� ��� ������. ����. � ���. �������"
  objRez.Cells(2, 2).Value = "�� 1"
  objRez.Cells(2, 3).Value = "�� 2"
  objRez.Cells(2, 4).Value = "�� 1+1"
  objRez.Cells(2, 5).Value = "�� 3"

  ' ��������� �� ����������
  Dim list()
  Dim j As Long
  Dim StartPos As Long
  Dim R As String
  j = 0

  StartPos = InStr(Protokol, "� � � � � � � � � �    � � � � � � �")
  StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)

  ' ������ ��� ����������� � ���� ���������, � ������ ������ ��� ������ ���� ����� ��� �������
  R = Parse_Rezhim_Single(Protokol, StartPos)

  Do
    ReDim Preserve list(j)
    list(j) = Array(R, Parse_Current_Line(Protokol, StartPos, StartPos))
    j = j + 1
    StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)
    If StartPos > 0 Then R = Parse_Rezhim_Single(Protokol, StartPos)
    DoEvents
  Loop While StartPos > 0

  ' ������ ������������ ������� � ��������������� ����� ��������, ��������� ��� �� ����
  For i = 0 To UBound(list)
    objRez.Cells(i + 3, 1).Value = list(i)(0)
    For j = 0 To 3
      objRez.Cells(i + 3, j + 2).Value = list(i)(1)(j + 1)
    Next j
  Next i

End Sub


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

  ' ���� ���� ���-2000, ���� ��� ��� ������� ��������� � �����������
  MainFormHandle = Find_TKZ_Window_Handle("TForm1")
  If MainFormHandle = 0 Then
    MsgBox "���� ���-2000 �� �������, ���������� ������ ���� ��������. ����� ����� ������ ���� ��������� ���� ��� �������.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' �������������, �������� ������ �� ������
  Initialize

  ' �������� ����� ���� (������-�� ���� ����� ���� �� ������ ��������)
  RootNode = Int(InputBox("����� ���� (�������������� ����)?", "RootNode", 0))
  If Not Node_Exists(RootNode) Then
    MsgBox "���� " & RootNode & " �� ������ � ����, ���������� ������ ����������.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' �� �������� � ����� ������ �������� ��������� (������� ��� ���������� ������) �������� ����� ����
  ' "����������� ������ ������� ��� �������..."
  Call SendMessage(MainFormHandle, WM_COMMAND, 12, 0&)

  ' ������ ���� ������� ������� ��������
  CommandFormHandle = Find_TKZ_Window_Handle("TFormZD2")

  ' ������� ���� ��������� � ������� ���
  Call SendMessage(MainFormHandle, WM_COMMAND, 12, 0&)
  Call SendMessage(CommandFormHandle, WM_COMMAND, 187, 0&)
  ProtokolHandle = Find_TKZ_Window_Handle("TForm3")
  ProtokolMemo = Find_SubClass_Recurce(ProtokolHandle, "TMemo")
  Call SendMessage(ProtokolHandle, WM_COMMAND, 247, 0&)

  ' ������� ������ ��� �������� ���������������� � �������� ��� � ���� �������� ���-2000
  CommandsText = Get_Sensitivity_Code()
  CommandRichEdit = Find_SubClass_Recurce(CommandFormHandle, "TRichEdit")
  Window_Set_Text CommandRichEdit, CommandsText

  ' ������ ������ � ������������������� - ��� ����������� �������
  Call SendMessage(CommandFormHandle, WM_COMMAND, 179, 0&)

  ' �������� ������� � ������� ��������� ��� �������
  ProtokolText = Window_Get_Text(ProtokolMemo)

  ' ������ ��������� �������
  Analiz_Sensitivity ProtokolText

  ' ���������� ������������ ��������� ������������ �������� (������� �������� �������� ������ � ������������)
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText

  ' ��������� ������������ ��������� �������� � ����
  Dim filePRT
  Dim FrFi As Integer
  
  filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\���������������� " & RootNode & " ����.prt", "����� ��������� ��� (*.prt), *.prt")
  If filePRT <> "False" Then
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, ProtokolText
    Close FrFi
  End If
    
  ' ������� ������ ��� �������� �����������
  
  
End Sub
