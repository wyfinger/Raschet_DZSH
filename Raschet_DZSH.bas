Attribute VB_Name = "Raschet_DZSH"
'
' ������ ���
' ~~~~~~~~~~
' https://github.com/wyfinger/Raschet_DZSH
' ����� �������, miv@prim.so-ups.ru
' 2013
'

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal uCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAcess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "PSAPI" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Const WM_COMMAND = &H111
Private Const WM_PASTE = &H302
Private Const WM_CUT = &H300
Private Const GW_CHILD = &H5
Private Const GW_HWNDNEXT = &H2
Private Const EM_SETSEL = &HB1
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Const WS_EX_TOOLWINDOW = &H80
Private Const WS_SIZEBOX = &H40000
Private Const WS_CAPTION = &HC00000
Private Const SW_NORMAL = 1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Dim GT_Class As String       ' ��������� ���������� ������ ��� �������� � Enum-�������
Dim GT_Result As Long

Dim RootNode As Long         ' ����� ������� ���� :)
Dim arrRootBranch()          ' ������ ������ �������� ����, ����������� � Get_Sensitivity_Code()

Dim arrBranch()              ' ������ ������ � ��� �����
Dim arrBranchCopy()
Dim arrBranchCopy2()
Dim arrNode()                ' ������ ������������ �����
Dim arrElement()             ' ������ ������������ ���������

Dim arrPowerNodes()          ' ������ �������� ����� (���� � ������ ����� �� RootNode � ������� ��������� ����)
                             ' ����������� � Find_Power_Nodes()
Dim arrBaseRejims()          ' ������ ������� ������� ��� �������� ����������������,
                             ' ����������� ��� ���������� ������� � Get_Testing_Code()
                             ' �.�. �� ��������� ������� ��� ���� �� ��������


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
  On Error GoTo 0

End Function


Private Function Extract_File_Name(FileName As String) As String
'
' ��������� ����� ����� �� ������� ���� � �����
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
' ��������� ���� � �����, ���������� ���� (��� ������������ �����)
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
' ����� �������� ���� ��������� ���-2000
'

  GT_Result = 0
  GT_Class = Class
  Call EnumWindows(AddressOf Find_Window_Enum_Proc, 0)
  Find_TKZ_Window_Handle = GT_Result

End Function


Private Function Find_SubClass_Recurce(hwnd As Long, sClassName As String, Optional iPos As Integer = 1) As Long
'
' ��� ������� ���������� ���������� ��� �������� ���� hWnd, ������ ����� ���� � sClassName
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
'
' ���������� � ���� ����� �����
'

  Dim d As Object
  Set d = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
  d.SetText (sText)
  d.PutInClipboard

  SendMessage hwnd, EM_SETSEL, 0, -1
  SendMessage hwnd, WM_PASTE, 0, 0

  Set d = Nothing

End Function


Private Function Window_Get_Text(hwnd As Long)
'
' �������� ����� �� ���� �����
'

  SendMessage hwnd, EM_SETSEL, 0, -1
  SendMessage hwnd, WM_CUT, 0, 0

  ' �������� ������ � �����
  Dim d As Object
  Set d = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
  d.GetFromClipboard
  Window_Get_Text = d.GetText

  Set d = Nothing

End Function


Private Function Show_Process_Window(Caption As String) As Long
'
' ����������� ������� ��������
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

End Sub


Private Function Find_Branch_By_Node(BranchArray, Node)
'
' ���� ��� �����, � ������� ������ �������� ����
'

  Dim Rez()
  Dim i As Integer
  Dim j As Integer
  j = 0
  For i = LBound(BranchArray) To UBound(BranchArray)
    If (Int(BranchArray(i, 3)) = Node) Or (Int(BranchArray(i, 4)) = Node) Then
      ReDim Preserve Rez(j)
      Rez(j) = i
      j = j + 1
    End If
  Next
  Find_Branch_By_Node = Rez

End Function


Private Function Find_Branch_By_2Node(BranchArray, Node1, Node2)
'
' ���� ���� ����� �� ��������� ���� �����, �� ����������, ���������� ����� �����
'
  Dim i, Rez As Integer
  Rez = -1

  For i = LBound(BranchArray) To UBound(BranchArray)
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
  For i = LBound(arrNode) To UBound(arrNode)
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
  For i = LBound(arrNode) To UBound(arrNode)
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
  For i = LBound(arrNode) To UBound(arrNode)
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
  For i = LBound(arrElement) To UBound(arrElement)
    If Int(arrElement(i, 1)) = Element Then
      Find_Element = Trim(arrElement(i, 2))
      Exit For
    End If
  Next

End Function


Private Function Find_Element_By_2Node(Branch, Node1, Node2)
'
' ���� ������ ���������, ������� ����� � ����� ������� ���� Node1 � Node2
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
' ���� ������ ����� �� ����� ����� � ������, ������� ����� �� ����� ��������
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
' �������� �������������������� �������
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


Private Function Array_Find(Source(), Val, Optional Col As Integer = -1) As Integer
'
' �������� ���������� � ������� Source �������� Val � ������� Col,
' ���� Col = -1 �������, ��� ������ ����������.
' ���������� ������ �������� (������) ��� -1, ���� �� �������
'

  Dim i As Integer
  Array_Find = -1

  If Not Array_Exists(Source) Then Exit Function

  For i = LBound(Source) To UBound(Source)
    If Col = -1 Then
      If Source(i) = Val Then
        Array_Find = i
      End If
    Else
      If Source(Col, i) = Val Then
        Array_Find = i
      End If
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
    "1-����    " & RootNode & "      /* " & Find_Node(RootNode) & vbCrLf & _
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
  Dim BranchType As Long
  Dim ElemName As String
  Dim NewStr As String
  Dim i As Long

  ' ��� ������� �� ������������� RootNode ������� ��������, ��� ��������� �������������
  arrRootBranch = Find_Branch_By_Node(arrBranch, RootNode)
  For i = LBound(arrRootBranch) To UBound(arrRootBranch)
    If arrBranch(arrRootBranch(i), 3) = RootNode Then
      CrossNode = arrBranch(arrRootBranch(i), 4)
    Else
      CrossNode = arrBranch(arrRootBranch(i), 3)
    End If
    ElemNo = arrBranch(arrRootBranch(i), 5)
    BranchType = arrBranch(arrRootBranch(i), 1)
    ' Issue#8: ������� ��������� ����������� ���
    If BranchType <> 101 Then
      ElemName = Find_Element(ElemNo)
      If (CrossNode = 0) Or (ElemNo = 0) Then               ' ��������� �����
        NewStr = "��������  " & (i + 2) & vbCrLf & _
        "����      0 *" & RootNode & "-" & CrossNode & "      /* �������� ?"
      Else                                                  ' ��������� �������
        NewStr = "��������  " & (i + 2) & vbCrLf & _
        "�������   " & ElemNo & "      /* " & ElemName
      End If
      R = R & NewStr & vbCrLf
    End If
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
  On Error Resume Next
    For i = 0 To 255          ' ������ ���-�� ������� ������� ������� :)
      If i = 0 Then
        NewSheetName = TempSheetName
      Else
        NewSheetName = TempSheetName & " #" & i
      End If
      If ActiveWorkbook.Worksheets(NewSheetName) Is Nothing Then Exit For
    Next
  On Error GoTo 0
  objRez.Name = NewSheetName

  objRez.Cells(1, 1).Value = "���� " & RootNode & " (" & Find_Node(RootNode) & ")"
  objRez.Cells(2, 1).Value = "��� ��� ������. ����. � ���. �������"
  objRez.Cells(2, 2).Value = "��(3)"
  objRez.Cells(2, 3).Value = "��(2)"
  objRez.Cells(2, 4).Value = "��(1+1)"
  objRez.Cells(2, 5).Value = "��(1)"

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
    DoEvents     ' ��� ����, ����� ������� ����� ��  Ctrl+C
  Loop While StartPos > 0

  ' ������ ������������ ������� � ��������������� ����� ��������, ��������� ��� �� ����
  For i = LBound(list) To UBound(list)
    objRez.Cells(i + 3, 1).Value = list(i)(0)
    For j = 0 To 3
      objRez.Cells(i + 3, j + 2).Value = list(i)(1)(j + 1)
    Next j
  Next i

End Sub


Private Function Delete_Interm_Nodes(Without As Long) As Long
'
' �������� ������������� �����, �.�. ��� � ������� ������ ��� �������������
' ������� �� ����������� ������� �������� ������������� �/�������, �� ����� ��������
' �� ��� ���, ���� �� ����� ������� ������� ��������� (������� ���������� ����������
' ��������� �����
'

  Dim i, j, n As Long
  Dim Node As Long
  Dim NodeBranch()
  Dim NodePosA, NodePosB As Long
  Dim ContrNodeA, ContrNodeB As Long
  Dim DestType, DestElement As Long

  j = 0
  For i = LBound(arrNode) To UBound(arrNode)
    Node = arrNode(i, 1)
    If Node <> Without Then
      NodeBranch = Find_Branch_By_Node(arrBranchCopy, Node)
      If Array_Exists(NodeBranch) Then
        n = UBound(NodeBranch)
        ' ������������� ����
        If n = 1 Then
          ' � ������ �������� ���� ������ ������� �������� ����, ����� �������� �� ������ ������� ����
          If arrBranchCopy(NodeBranch(0), 3) = Node Then NodePosA = 3
          If arrBranchCopy(NodeBranch(0), 4) = Node Then NodePosA = 4
          If arrBranchCopy(NodeBranch(1), 3) = Node Then NodePosB = 3
          If arrBranchCopy(NodeBranch(1), 4) = Node Then NodePosB = 4

          ' ������ ��������������� ���� (������ �����)
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
            ' ���� � ����� �� ������ �� �������������� ���� ����� ���� 3 ��� 4 ���� - ��������������
            ' ����� ������ ���� ���� ����������� ��� ���������������
            DestType = 0
            ' �������� ��� �����
            If (arrBranchCopy(NodeBranch(0), 1) > 1) Then DestType = arrBranchCopy(NodeBranch(0), 1)
            If (arrBranchCopy(NodeBranch(1), 1) > 1) Then DestType = arrBranchCopy(NodeBranch(1), 1)
            ' ����� �������� ���������������� �� Without (��� RootNode), �����
            ' ����� ����� ���� ����� ���������� ������������� �� RootNode, ������ � ��������� ����
            DestElement = 0
            If (arrBranchCopy(NodeBranch(0), 3) = Without) Or (arrBranchCopy(NodeBranch(0), 4) = Without) Then DestElement = arrBranchCopy(NodeBranch(0), 5)
            If (arrBranchCopy(NodeBranch(1), 3) = Without) Or (arrBranchCopy(NodeBranch(1), 4) = Without) Then DestElement = arrBranchCopy(NodeBranch(1), 5)

            ' ���������
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
        ' ������ (����, ������� ������ ������ � ���� �����)
        If n = 0 Then
          arrBranchCopy(NodeBranch(0), 3) = 0
          arrBranchCopy(NodeBranch(0), 4) = 0
        End If
      End If
    End If
    DoEvents     ' ��� ����, ����� ������� ����� ��  Ctrl+C
  Next
  Delete_Interm_Nodes = j

End Function


Private Sub Find_Power_Nodes()
'
' ����� �������� ����� ��� RootNode
'

  Dim i, j, n, k, l As Long

  ' ������� ����� ����� � 101 ����� (����������� ���)
  For i = LBound(arrBranchCopy) To UBound(arrBranchCopy)
    If arrBranchCopy(i, 1) = 101 Then
      arrBranchCopy(i, 3) = 0
      arrBranchCopy(i, 4) = 0
    End If
    ' ����� �������� ������ ��������� � ������ ������� ������
    arrBranchCopy(i, 5) = 0
    arrBranchCopy2(i, 5) = 0
  Next

  ' ������� �� �������������� RootNode � ��������� ��������� �� ���� ������ ���������� �����
  ' ��������. ��� �������� ������������� ���������� � Delete_Interm_Nodes()
  ' ���� ���� �� ����� ����� = RootNode ��� ����� �������� ����� ���������������� �� �����
  ' ������������ �����

  For i = LBound(arrRootBranch) To UBound(arrRootBranch)
    arrBranchCopy(arrRootBranch(i), 5) = i + 1
    arrBranchCopy2(arrRootBranch(i), 5) = i + 1
  Next

  Do
    j = 0
    ' ������� ������������� ���� � ������ (���� ������ � ����� ������)
    n = Delete_Interm_Nodes(RootNode)
    j = j + n

    ' ������� ������ � ���� ��������� � ��-��� �� ���� (���), �� �� ����������
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

  ' ���� ��������������� ����
  Dim list()
  Dim NodeBranch()
  Dim DestNode As Long
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
      ' �� ��������� ���������, ������� ����� ��������� �� �� ������������ �� ������� (������)
      If (Array_Find(list, DestNode) = -1) And (DestNode <> 0) Then
        ReDim Preserve list(j)
        list(j) = DestNode
        j = j + 1
      End If
    Next
  End If

  Dim PowerNode, e As Long
  Dim SecondNode As Long
  Dim Elem(), n_node, n_branch
  l = 0

  ' ���������� ������ �������� ����� � ������ ����� ������������� �� ���
  NodeBranch = Find_Branch_By_Node(arrBranchCopy2, RootNode)
  For i = LBound(list) To UBound(list)
    PowerNode = list(i)
    ' ������ �����(�) ���������, � ������� ������ RootNode � PowerNode, ���� � ��������� ���� ���� �� ���� ����
    ' ���� ��������� ����� ���� ���������
    Elem = Find_Element_By_2Node(arrBranchCopy, RootNode, PowerNode)
    If Array_Exists(Elem) Then
      n = UBound(Elem)
      ' ������ ����� ������������� RootNode ������������� � ��������� Elem
      For j = LBound(Elem) To n
        e = Elem(j)
        For k = LBound(NodeBranch) To UBound(NodeBranch)
          If arrBranchCopy2(NodeBranch(k), 5) = e Then
            If arrBranchCopy2(NodeBranch(k), 3) = RootNode Then
              SecondNode = arrBranchCopy2(NodeBranch(k), 4)
            Else
              SecondNode = arrBranchCopy2(NodeBranch(k), 3)
            End If
            n_node = Trim(Find_Node(list(i)))                                       ' ???
            n_branch = Find_Branch_By_2Node(arrBranch, RootNode, SecondNode)        ' ???
            n_branch = arrBranch(n_branch, 5)                                       ' ???
            n_branch = Trim(Find_Element(n_branch))                                 ' ???
            ' ������� �������� ���� � ������ ����� �� ���� � arrPowerNodes()
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
' ���������� ������� ��� ������ ���������������� ��� � ������ �����������.
' ��� ������� ��������� ���� ������� �������� � ������� ��������� ��� �������������
' RootNode ����� ������� � PowerNode ����� � �� ������ ������� ������ ������ �������
' ��������� � ������� ��������� �� ������ ������������� ��������� ����
'

  Dim R As String
  Dim i, j, k As Long

  R = _
    "*         �������� ���������������� ��� ��� �����������, ���� " & RootNode & " [" & Find_Node(RootNode) & "]" & vbCrLf & _
    "��������  IA IB IC" & vbCrLf & _
    "1-����    " & RootNode & "      /* " & Find_Node(RootNode) & vbCrLf & _
    "����      1" & vbCrLf & _
    "���-���   " & RootNode & "/ABC" & vbCrLf & _
    "����      2" & vbCrLf & _
    "���-���   " & RootNode & "/AB" & vbCrLf & _
    "����      3" & vbCrLf & _
    "���-���   " & RootNode & "/AB0" & vbCrLf & _
    "����      4" & vbCrLf & _
    "���-���   " & RootNode & "/A0" & vbCrLf

  Dim BaseRejim, Podrejim As Long
  Dim PowerNode, NodeA, NodeB, CrossNode, T As Long
  Dim branchNo As Long
  Dim ElemName As String
  Dim NodeBranch()
  Dim Elem, rElem As Long
  Dim Collision As Boolean

  Podrejim = 1
  ' �������� �� ���� ��������������, ��������� � ������ (������������� � ��������� ����)
  For i = LBound(arrPowerNodes) To UBound(arrPowerNodes)
    PowerNode = arrPowerNodes(i)(0)
    NodeA = arrPowerNodes(i)(1) ' RootNode
    NodeB = arrPowerNodes(i)(2) ' ����� ���������������� ���� ������ ����� ������������� � ��������� ����
    R = R & vbCrLf
    R = R & "��������  " & Podrejim & " /* " & PowerNode & " (" & Find_Node(PowerNode) & ")" & vbCrLf
    BaseRejim = Podrejim
    ' ������� �������� �������� ������, ����� ����� ���� ��� ������� � �������������� ����,
    ' �� ��������� ��� ���������� �� �������
    ReDim Preserve arrBaseRejims(i)
    branchNo = Find_Branch_By_2Node(arrBranch, NodeA, NodeB)
    ElemName = Find_Element(arrBranch(branchNo, 5))
    arrBaseRejims(i) = Array(BaseRejim, Find_Node(PowerNode) & " (" & ElemName & ")")
      
    ' ��������� �� ���� �������������� RootNode
    For j = LBound(arrRootBranch) To UBound(arrRootBranch)
      ' ��� ������, ��������� �� RootNode ������ ����� ���������������� ����
      If arrBranch(arrRootBranch(j), 3) = RootNode Then
        CrossNode = arrBranch(arrRootBranch(j), 4)
      Else
        CrossNode = arrBranch(arrRootBranch(j), 3)
      End If
      ' ��������� ��� �������������, ����� �������� � ��������� ����
      If CrossNode = NodeB Then
        ' ��� ������� ������� ���������� � �����������
        R = R & "*"
      End If
      R = R & "����      *" & RootNode & "-" & CrossNode & _
      "      /* ������� " & arrBranch(arrRootBranch(j), 5) & " (" & _
      Find_Element(arrBranch(arrRootBranch(j), 5)) & "), ����� (" & _
      Find_Node(arrBranch(arrRootBranch(j), 3)) & " - " & Find_Node(arrBranch(arrRootBranch(j), 4)) & _
      ")" & vbCrLf
    Next
    
    ' ������ ��� ������������� ��������� ���� � �������� ������ � ��������� ���������, ���������� �� BaseRejim
    NodeBranch = Find_Branch_By_Node(arrBranch, PowerNode)
    For j = LBound(NodeBranch) To UBound(NodeBranch)
      T = arrBranch(NodeBranch(j), 1)
      
      ' Issue#2: ��������, ��� �����, ��������� �� ��������� ����, ������� �� ����� ���������,
      ' �� ����� � RootNode (��� ����� RootNode ����� ����� ��������� � ������� ������)
      Elem = arrBranch(NodeBranch(j), 5)  ' ����� �������� ��� ����� �� PowerNode, ������� ����� ���������
      Collision = False
      For k = LBound(arrRootBranch) To UBound(arrRootBranch)
        rElem = arrBranch(arrRootBranch(k), 5)
        If Elem = rElem Then
          Collision = True
          Exit For
        End If
      Next
      ' Issue#8: ������� ��������� ����������� ���
      If (T <> 101) And Not Collision Then
        Podrejim = Podrejim + 1
        R = R & "��������  " & Podrejim & " " & BaseRejim & vbCrLf
        NodeA = arrBranch(NodeBranch(j), 3)
        NodeB = arrBranch(NodeBranch(j), 4)
               
        If NodeA = PowerNode Then
          CrossNode = NodeB
        Else
          CrossNode = NodeA
        End If
      
        If Elem = 0 Then
          R = R & "����      *" & PowerNode & "-" & CrossNode & _
          " /* " & Find_Node(PowerNode) & " - " & Find_Node(CrossNode) & vbCrLf
        Else
          R = R & "�������   " & Elem & _
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
' �������� ����� �� ������ � ���������, ���������� �� ������ ������
' (������ ��� ���������� �� �������� ����)
'

  Dim apos, bpos, cpos As Long

  Is_Sub_Rejim = False
  apos = InStr(CurPos, Protokol, "��������  ")

  If apos > 0 Then
    bpos = InStr(apos + 10, Protokol, " ")
    cpos = InStr(apos + 10, Protokol, vbCrLf)
    If cpos > bpos Then Is_Sub_Rejim = True
  End If

End Function


Private Function Find_BaseRejim_Name(RejimNo)
'
' ���� ������������ ������ �� ������� ������� � arrBaseRejims()
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
' ���� ����� �� ������� ��������� - ����� ��� ������������,
' ���� ��� �������� - ����� ������������ ������������ ��������
'

  Dim apos, bpos As Long
  Dim Prefix As String

  ' ������ ����� ������, ���� ��� �������� - ����� ������� �� ���� �����,
  ' ���� �������� - ������
  apos = InStr(CurPos, Protokol, "��������")
  bpos = InStr(apos, Protokol, vbCrLf)
  If (apos > 0) And (bpos > apos) Then
    Prefix = Trim(Mid(Protokol, apos + 8, bpos - apos - 8))
  End If

  apos = 0
  bpos = 0

  If Is_Sub_Rejim(Protokol, CurPos) Then
    apos = InStr(CurPos, Protokol, "(")
    bpos = InStr(apos, Protokol, ")")
    If (apos > 0) And (bpos > apos) Then
      Get_Rejim_Name = "[" & Prefix & "] +���� " & Trim(Mid(Protokol, apos + 1, bpos - apos - 1))
    End If
  Else
    apos = InStr(CurPos + 10, Protokol, vbCrLf)
    Get_Rejim_Name = "[" & Prefix & "] " & Find_BaseRejim_Name(Trim(Mid(Protokol, CurPos + 10, apos - (CurPos + 10))))
  End If
  
End Function


Private Sub Analiz_Testing(Protokol As String)
'
' ������� ��������� �� �����������
'

  Dim objWorkbook, objRez
  Dim s, i As Long

  ' ���������� ����� �������� � ��� �� ����, ��� � ���������� �� �������� ����������������
  ' � ����������� ������
  Set objWorkbook = ActiveWorkbook
  Set objRez = objWorkbook.ActiveSheet
  s = objRez.UsedRange.Rows.Count + 2

  objRez.Cells(s, 1).Value = "��� ��� �����������"
  objRez.Cells(s, 2).Value = "��(3)"
  objRez.Cells(s, 3).Value = "��(2)"
  objRez.Cells(s, 4).Value = "��(1+1)"
  objRez.Cells(s, 5).Value = "��(1)"

  ' ��������� �� ����������
  Dim list()
  Dim j As Long
  Dim StartPos As Long
  Dim RejimName As String
  Dim Line
  j = 0

  StartPos = InStr(Protokol, "� � � � � � � � � �    � � � � � � �")
  StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)

  Do
    ReDim Preserve list(j)
    RejimName = Get_Rejim_Name(Protokol, StartPos)
    Line = Parse_Current_Line(Protokol, StartPos, StartPos)
      
    list(j) = Array(RejimName, Line)
    j = j + 1
    StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)
    DoEvents     ' ��� ����, ����� ������� ����� ��  Ctrl+C
  Loop While StartPos > 0

  ' ������ ������������ ������� � ��������������� ����� ��������, ��������� ��� �� ����
  objRez.Cells(s, 1).Value = "��� ��� ������ �����������"
  For i = LBound(list) To UBound(list)
    objRez.Cells(s + i + 1, 1).Value = list(i)(0)
    For j = 0 To 3
      objRez.Cells(s + i + 1, j + 2).Value = list(i)(1)(j + 1)
    Next
  Next

End Sub


'######################################################################################[������� ����� �������]

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

  ' ������ �� Ctrl ?
  ManualMode = CBool(GetKeyState(&H11) < 0)

  ' ���� ���� ���-2000, ���� ��� ��� ������� ��������� � �����������
  MainFormHandle = Find_TKZ_Window_Handle("TForm1")
  If MainFormHandle = 0 Then
    MsgBox "���� ���-2000 �� �������, ���������� ������ ���� ��������. ����� ����� ������ ���� ��������� ���� ��� �������.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' �������������, �������� ������ �� ������
  Call Initialize

  ' �������� ����� ���� (������-�� ���� ����� ���� �� ������ ��������)
  Dim Answer
  Answer = InputBox("����� ���� (�������������� ����)?", "RootNode", 0)
  If Trim(Answer) = "" Then Exit Sub
  RootNode = Int(Answer)
  If Not Node_Exists(RootNode) Then
    MsgBox "���� " & RootNode & " �� ������ � ����, ���������� ������ ����������.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' ��������� ��������� ������
  ProcessWnd = Show_Process_Window("������� ����, �����...")

  ' �� �������� � ����� ������ �������� ��������� (������� ��� ���������� ������) �������� ����� ����
  Call SendMessage(MainFormHandle, WM_COMMAND, 12, 0&)     ' "����������� ������ ������� ��� �������..."

  ' ������ ���� ������� ������� ��������
  CommandFormHandle = Find_TKZ_Window_Handle("TFormZD2")

  ' ������� ���� ��������� � ������� ���
  Call SendMessage(CommandFormHandle, WM_COMMAND, 187, 0&) ' "������� ��������..."
  Call SendMessage(CommandFormHandle, WM_COMMAND, 200, 0&) ' "�������� �������"
  ProtokolHandle = Find_TKZ_Window_Handle("TForm3")
  ProtokolMemo = Find_SubClass_Recurce(ProtokolHandle, "TMemo")
  Call SendMessage(ProtokolHandle, WM_COMMAND, 247, 0&)    ' "�������� ��������"

  ' <<< �������� ���������������� ��������/������������� �������

  ' ������� ������ ��� �������� ���������������� � �������� ��� � ���� �������� ���-2000
  CommandsText = Get_Sensitivity_Code()
  CommandRichEdit = Find_SubClass_Recurce(CommandFormHandle, "TRichEdit")
  Window_Set_Text CommandRichEdit, CommandsText

  If ManualMode Then MsgBox _
  "� ���� ����� �������� ���-2000 �������� ������ �������� ����������������, �������������� ��������, ���� �� ������ �������� ���," & _
  "�������� ���." & vbCrLf & _
  "��� ����������� ������ ������� ������� Ok"

  ' ������ ������ � ������������������� - ��� ����������� �������
  Call SendMessage(CommandFormHandle, WM_COMMAND, 179, 0&) ' "������ � ������������������"

  ' ������� ��������� ��� �������
  ProtokolText = Window_Get_Text(ProtokolMemo)

  ' ������ ��������� �������
  Analiz_Sensitivity ProtokolText

  ' ���������� ������������ ��������� ����������� �������� (������� �������� �������� ������ � �������������)
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText

  Dim filePRT
  Dim FrFi As Integer

  filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\���������������� " & RootNode & " ����.prt", "����� ��������� ��� (*.prt), *.prt")
  If filePRT <> "False" Then
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, ProtokolText
    Close FrFi
  End If

  ' <<< �������� ���������������� � ������ �����������

  ' ������� ������ ��� �������� �����������
  Call Find_Power_Nodes

  CommandsText = Get_Testing_Code()
  
  ' ������� �������� � ���� ��������, ��������� ������ � �������� ���
  Call SendMessage(ProtokolHandle, WM_COMMAND, 247, 0&)    ' "�������� ��������"
  Call SendMessage(CommandFormHandle, WM_COMMAND, 200, 0&) ' "�������� �������"

  Window_Set_Text CommandRichEdit, CommandsText
  
  If ManualMode Then MsgBox _
  "� ���� ����� �������� ���-2000 �������� ������ �������� ���������������� ��� �����������, �������������� ��������, ���� �� ������ �������� ���," & _
  "�������� ���." & vbCrLf & _
  "��� ����������� ������ ������� ������� Ok"
  
  Call SendMessage(CommandFormHandle, WM_COMMAND, 179, 0&) ' "������ � ������������������"
  
  ' ������� ��������� ��� �������
  ProtokolText = Window_Get_Text(ProtokolMemo)
  
  ' ����������� ���������� �������
  Analiz_Testing (ProtokolText)

  ' ���������� ������������ ��������� ����������� �������� � �������������
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText
  filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\����������� " & RootNode & " ����.prt", "����� ��������� ��� (*.prt), *.prt")
  If filePRT <> "False" Then
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, ProtokolText
    Close FrFi
  End If

  ' ������ ������ ��������
  If ProcessWnd > 0 Then DestroyWindow (ProcessWnd)

End Sub
