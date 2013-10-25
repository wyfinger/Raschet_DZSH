VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDZSH 
   Caption         =   "������ ��� / ���, 2013-08-09"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9825
   OleObjectBlob   =   "frmDZSH.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDZSH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RootNode As Integer    ' ����� ������� ���� :)

Dim arrBranch()
Dim arrBranchCopy()
Dim arrBranchCopy2()
Dim arrNode()
Dim arrElement()
Dim Initialized
Dim arrTrueBrach()  ' ������ ������������� ����, ����� �������������
Dim arrBaseRejims() ' �����, �������� ��� ������� �������, ����� ��� �������� ��������� �� �����������

Const vbTab = "   "   ' ���� ��������� ��� ���������� �� ��������� �������� � �����



Private Sub Initialize()
'
' ���������� - ������� ������ � ����� � ����������
'

' 1. ����� - ����� ������ ���� �������� �������� '������� ������'
Set objWorkbook = ActiveWorkbook
Set wshBranch = ActiveWorkbook.Worksheets("������� ������")
arrBranch = wshBranch.Range("A3:K" & wshBranch.UsedRange.Rows.Count).Value2
arrBranchCopy = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2
arrBranchCopy2 = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2

' 2. ���� - ����� ������ ��� ������� �������� '����.�����'
Set wshNode = objWorkbook.Worksheets("����.�����")
arrNode = wshNode.Range("A3:E" & wshNode.UsedRange.Rows.Count).Value2

' 3. �������� - ����� ������ ��� ������� �������� '����.���������'
Set wshElement = objWorkbook.Worksheets("����.���������")
arrElement = wshElement.Range("A3:B" & wshElement.UsedRange.Rows.Count).Value2

Initialized = True

End Sub


Private Function Find_Branch_By_Node(BranchArray, Node)
'
' ���� ��� �����, � ������� ������ ������� ����
'

Dim rez()
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

Find_Branch_Index = 0
For i = 1 To UBound(arrBranch)
  If ((arrBranch(i, 3) = Node1) And (arrBranch(i, 4) = Node2)) Or _
     ((arrBranch(i, 3) = Node2) And (arrBranch(i, 4) = Node1)) Then
     Find_Branch_Index = i
     Exit For
   End If
Next

End Function

Private Sub cbProcess2_Click()
'
' ���������� ������� �������� ����������������
'

' ����� �������� ����, ��������� �������������
arrLines = Split(frmDZSH.tbBranchList.Text, vbCrLf)
Dim list()
j = 0
For i = 0 To UBound(arrLines)
  ' �������� ������������
  cpos = InStr(arrLines(i), "/*")
  If cpos > 0 Then arrLines(i) = Mid(arrLines(i), 1, cpos - 1)
  On Error GoTo skeep
  ' ��������� ����� ��������� ���� � �����-�������������, ������� � ����
  apos = InStr(arrLines(i), "(") + 1
  bpos = InStr(arrLines(i), "-") + 1
  cpos = InStr(arrLines(i), ")") + 1
  PowerNode = Int(Trim(Mid(arrLines(i), 1, apos - 3)))
  NodeA = Int(Trim(Mid(arrLines(i), apos, bpos - apos - 1)))
  NodeB = Int(Trim(Mid(arrLines(i), bpos, cpos - bpos - 1)))
  ' ������� ��� � ������
  ReDim Preserve list(j)
  list(j) = Array(PowerNode, NodeA, NodeB)
  j = j + 1
skeep:
Next

' �������������� ����� �������
frmDZSH.tbCommandList = _
"��������  IA IB IC" & vbCrLf & _
"1-����    " & RootNode & vbTab & "/* " & Find_Node(RootNode) & vbCrLf & _
"����      1" & vbCrLf & _
"���-���   " & RootNode & "/ABC" & vbCrLf & _
"����      2" & vbCrLf & _
"���-���   " & RootNode & "/AB" & vbCrLf & _
"����      3" & vbCrLf & _
"���-���   " & RootNode & "/AB0" & vbCrLf & _
"����      4" & vbCrLf & _
"���-���   " & RootNode & "/A0" & vbCrLf

Podrejim = 1
k = 0
' �������� �� ���� ��������������, ��������� � ������ (������������� � ��������� ����)
For i = 0 To UBound(list)
   PowerNode = list(i)(0)
   NodeA = list(i)(1) ' RootNode
   NodeB = list(i)(2) ' ����� ���������������� ���� ������ ����� ������������� � ��������� ����
   frmDZSH.tbCommandList = frmDZSH.tbCommandList & vbCrLf
   frmDZSH.tbCommandList = frmDZSH.tbCommandList & "��������  " & Podrejim & " /* " & PowerNode & " [" & Find_Node(PowerNode) & "]" & vbCrLf
   BaseRejim = Podrejim
   ' ������� �������� �������� ������, ����� ����� ���� ��� ������� � �������������� ����,
   ' �� ��������� ��� ���������� �� �������
   ReDim Preserve arrBaseRejims(i)
   branchNo = Find_Branch_By_2Node(arrBranch, NodeA, NodeB)
   elemName = Find_Element(arrBranch(branchNo, 5))
   arrBaseRejims(i) = Array(Podrejim, Find_Node(PowerNode) & " [" & elemName & "]")
      
   ' ��������� �� ���� ����������� �� ��� �������������� (�������� ������������� ����� 1 ����)
   For j = 0 To UBound(arrTrueBrach)
     ' ��������� ��� �������������, ����� ��������
     If arrTrueBrach(j)(2) <> NodeB Then
    '   frmDZSH.tbCommandList = frmDZSH.tbCommandList & "����      *" & arrTrueBrach(j)(1) & "-" & arrTrueBrach(j)(2) & _
    '     " /* " & Find_Node(arrTrueBrach(j)(1)) & " - " & Find_Node(arrTrueBrach(j)(2)) & vbCrLf
       frmDZSH.tbCommandList = frmDZSH.tbCommandList & "����      *" & arrTrueBrach(j)(1) & "-" & arrTrueBrach(j)(2) & _
         " /* " & Find_Element(arrTrueBrach(j)(0)) & vbCrLf
     Else ' ��� �������
       frmDZSH.tbCommandList = frmDZSH.tbCommandList & "* ����      *" & arrTrueBrach(j)(1) & "-" & arrTrueBrach(j)(2) & _
         " /* " & Find_Node(arrTrueBrach(j)(1)) & " - " & Find_Node(arrTrueBrach(j)(2)) & vbCrLf
     End If
   Next
   
   ' ������ ��� ������������� ��������� ���� � �������� ������ � ��������� ���������, ���������� �� BaseRejim
   NodeBranch = Find_Branch_By_Node(arrBranch, PowerNode)
   ' Issue#2: ���������, ����� �� ��������� ���-������ ������. ������ ������������� RootNode
   RootNodeBranch = Find_Branch_By_Node(arrBranch, RootNode)
   For j = 0 To UBound(NodeBranch)
     T = arrBranch(NodeBranch(j), 1)
     If T <> 101 Then
       Podrejim = Podrejim + 1
       frmDZSH.tbCommandList = frmDZSH.tbCommandList & "��������  " & Podrejim & " " & BaseRejim & vbCrLf
       NodeA = arrBranch(NodeBranch(j), 3)
       NodeB = arrBranch(NodeBranch(j), 4)
       
       ' Issue#2: ��������, ��� �����, ��������� �� ��������� ����, �������� �� ����� ���������,
       ' �� ������� � RootNode (��� ����� RootNode ����� ����� ��������� � ������� ������)
       Elem = arrBranch(NodeBranch(j), 5)    ' ����� �������� ��� ����� �� PowerNode, ������� ����� ���������
       Collision = False
       For k = 0 To UBound(RootNodeBranch)
         rElem = arrBranch(RootNodeBranch(k), 5)
         If Elem = rElem Then
           Collision = True
           Exit For
         End If
       Next
       
       If Not Collision Then
       
         If NodeA = PowerNode Then ContrNode = NodeB Else ContrNode = NodeA
         If Elem = 0 Then
           frmDZSH.tbCommandList = frmDZSH.tbCommandList & "����      *" & PowerNode & "-" & ContrNode & _
             " /* " & Find_Node(PowerNode) & " - " & Find_Node(ContrNode) & vbCrLf
         Else
           frmDZSH.tbCommandList = frmDZSH.tbCommandList & "�������   " & Elem & _
             " /* " & Find_Element(Elem) & vbCrLf
         End If
       End If
     End If
   Next
   Podrejim = Podrejim + 1
   
Next

' �������� ������ � �����
Dim d As New DataObject
d.SetText (frmDZSH.tbCommandList.Text)
d.PutInClipboard

cbAnaliz2.Enabled = True
If Not cbMessages.Value Then
  MsgBox "������ ��� ��� ��� �������� ������ ����������� (�������������� �����) ����������� " & _
    "� ���������� � ����� ������, ���������� �������� ���� ��������� ��� ���, �������� ������ � ��������� ������ " & _
    "����� ���� ����������� ���� ����� � ��������� �.6 (������ ����������� �������)."
End If

End Sub

' #################################################################################

Function Find_TKZ_Handle() As Long
'
' ����� �������� ���� ���-2000
'



End Function

' #################################################################################



Private Sub CommandButton1_Click()
'
' ���� ������� �������
'

' �������������, ����� ������ � �����
If Not Initialized Then Call Initialize

If RootNode = 0 Then RootNode = Int(InputBox("����� ���� (�������������� ����)?", "RootNode", 0))

' � �� ������ �� ������������?
If Not Node_Exists(RootNode) Then
  MsgBox "���� " & RootNode & " ����������� � ������� �����", vbOKOnly + vbExclamation
  Exit Sub
End If

' ������ ���� TKZ, ���� ��� ��� - ����������� � ���������� �� ������
TKZ_Handle& = FindWindow("TFormZD2", "")

End Sub

Private Sub UserForm_Initialize()
'
' �������� �������� ��������� (����� � ����������)
'

cbMessages.Value = GetSetting("Raschet_DZSH", "Settings", "ShowMessages", True)

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'
' ���������� �������� �����, ��������� ��������� ����������
'

SaveSetting "Raschet_DZSH", "Settings", "ShowMessages", cbMessages.Value

End Sub


Private Sub cbFindBranch_Click()
'
' ���� ��� ����� �� ��������� ����, ������� ��� ������ ��������, � �������� � tbBranchList
'

' �������������, ����� ������ � �����

If Not Initialized Then Call Initialize
If RootNode = 0 Then RootNode = Int(InputBox("����� ���� (�������������� ����)?", "RootNode", 0))

If Not Node_Exists(RootNode) Then
  MsgBox "���� " & RootNode & " ����������� � ������� �����"
  Exit Sub
End If

' �������������� ������ ������� ������, ��� �������������� (������������� �������������)
BranchList = Find_Branch_By_Node(arrBranch, RootNode)
frmDZSH.tbBranchList.Text = "/* ���� " & RootNode & " (" & Find_Node(RootNode) & ") - �������������: ������� (�����)" & vbCrLf
For i = 0 To UBound(BranchList)
  T = Int(arrBranch(BranchList(i), 1))       ' ��� �����
  A = Int(arrBranch(BranchList(i), 3))       ' ���� ������
  b = Int(arrBranch(BranchList(i), 4))       ' ���� �����
  e = Int(arrBranch(BranchList(i), 5))       ' ����� ��������
  If T <> 101 Then
    If A = RootNode Then �ross = b Else �ross = A  ' ����� ���� � ��������������� ������� �����
    ' ���� ����� �������� 0 � ��������������� ���� - ���� ���� - ����� � �����������, ��� ��� ��������
    If (�ross = 0) And (e = 0) Then            ' !!!!!!
      str2add = e & vbTab & "(" & RootNode & "-" & �ross & ")" & vbTab & "/* " & "��������"
    Else
      str2add = e & vbTab & "(" & RootNode & "-" & �ross & ")" & vbTab & "/* " & Find_Element(e)
    End If
    frmDZSH.tbBranchList.Text = frmDZSH.tbBranchList.Text & str2add & vbCrLf
  End If
Next

cbProcess1.Enabled = True
Label1.Caption = "������������� ����, ������� �����, ������� �� ����� ���� ��������� :"
If Not cbMessages.Value Then
  MsgBox "����������� ������ ������, ����������� � ���������� ����, ������ ���� ������ ����� ��������������� " & _
    "(������� �����, ������� �� ������ �� ����� ���� ���������), ����� ���� ��������� �.2"
End If

End Sub


Private Sub cbProcess1_Click()

' ������ ������ ������, ��������� ����� �������� � ������ � ����� �����
' ����� ��� ���������� ���������
arrLines = Split(frmDZSH.tbBranchList.Text, vbCrLf)
j = 0
For i = 0 To UBound(arrLines)
  ' �������� ������������
  cpos = InStr(arrLines(i), "/*")
  If cpos > 0 Then arrLines(i) = Mid(arrLines(i), 1, cpos - 1)
  pos1 = InStr(arrLines(i), "(")
  pos2 = InStr(arrLines(i), "-")
  pos3 = InStr(arrLines(i), ")")
  If pos1 > 0 Then
    e = Int(Mid(arrLines(i), 1, pos1 - 1))
    A = Int(Mid(arrLines(i), pos1 + 1, pos2 - pos1 - 1))
    b = Int(Mid(arrLines(i), pos2 + 1, pos3 - pos2 - 1))
    ReDim Preserve arrTrueBrach(j)
    arrTrueBrach(j) = Array(e, A, b)
    j = j + 1
  End If
Next

' �������������� ������
frmDZSH.tbCommandList = _
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
"��������  1" & vbTab & "/* ��� ��������" & vbCrLf

' ��� ������� �� �������������, ������� ���������� ������������, ������� ��������
For i = 0 To UBound(arrTrueBrach)
  If arrTrueBrach(i)(1) = RootNode Then �ross = arrTrueBrach(i)(2) Else �ross = arrTrueBrach(i)(1)
  Elem = Find_Element(arrTrueBrach(i)(0))
  If (�ross = 0) Or (Elem = 0) Then               ' ��������� �����
    str2add = "��������  " & (i + 2) & vbCrLf & _
    "����      0 *" & RootNode & "-" & �ross & vbTab & "/* �������� ??"
  Else                                                                 ' ��������� �������
    str2add = "��������  " & (i + 2) & vbCrLf & _
    "�������   " & arrTrueBrach(i)(0) & vbTab & "/* " & Find_Element(arrTrueBrach(i)(0))
  End If
  
  frmDZSH.tbCommandList = frmDZSH.tbCommandList & str2add & vbCrLf
Next

' �������� ������ � �����
Dim d As New DataObject
d.SetText (frmDZSH.tbCommandList.Text)
d.PutInClipboard

cbAnaliz1.Enabled = True
If Not cbMessages.Value Then
  MsgBox "������ ��� ��� ��� �������� ���������������� �������� � ������������� ������� ��� ����������� " & _
    "� ���������� � ����� ������, ���������� �������� ���� ��������� ��� ���, �������� ������ � ��������� ������ " & _
    "����� ���� ����������� ���� ����� � ��������� �.3 (������ ����������� �������)."
End If

End Sub


Private Function Parse_Current_Line(T, FromPos, ByRef FinishPos)
'
' ������� ���� ��������� ���� �� � ������������ � ������ ��������� � ������� �
' FromPos, � �������� FinishPos ������������ �������, �� ������� ����������� ������
'

Dim Imin(1 To 4)
For i = 1 To 4
  FinishPos = InStr(FromPos, T, "����      " & i)
  FinishPos = InStr(FinishPos, T, "I����")
  ia = Int(Mid(T, FinishPos + 5, 10))
  FinishPos = InStr(FinishPos, T, "I����")
  IB = Int(Mid(T, FinishPos + 5, 10))
  If (IB < ia) And (IB > 0) Then Imin(i) = IB Else Imin(i) = ia
Next
Parse_Current_Line = Imin

End Function


Private Function Parse_Rezhim_Single(T, FromPos)
'
' ���� ��� ����������� � ���������, ����� �������������� ������ ���� ����������
'

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


Private Sub cbAnaliz1_Click()
'
' ������ ��������� ��� ��� � ������������ ����� � ��������� �����
'

' ��������� ���� ��� �����������
Set objWorkbook = ActiveWorkbook
Set objRez = objWorkbook.Worksheets.Add
objRez.Columns("A:A").ColumnWidth = 35#
'objRez.Name = RootNode & " " & Find_Node(RootNode)

' ����� ����� ��������� �� ������ ������
Dim d As New DataObject
d.GetFromClipboard
Protokol = d.GetText

' ��������� �� ����������
StartPos = InStr(Protokol, "� � � � � � � � � �    � � � � � � �")
Dim list()
j = 0
StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)

' ������ ��� ����������� � ���� ���������, � ������ ������ ��� ������ ���� ����� ��� �������
r = Parse_Rezhim_Single(Protokol, StartPos)

Do
  ReDim Preserve list(j)
  list(j) = Array(r, Parse_Current_Line(Protokol, StartPos, StartPos))
  j = j + 1
  StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)
  If StartPos > 0 Then r = Parse_Rezhim_Single(Protokol, StartPos)
  DoEvents
Loop While StartPos > 0

' ������ ������������ ������� � ��������������� ����� ��������, ��������� ��� �� ����
objRez.Cells(1, 1).Value = "���� " & RootNode & " (" & Find_Node(RootNode) & ") - ��� ��� ���������������� ����. � ���. �������"
For i = 0 To UBound(list)
  objRez.Cells(i + 2, 1).Value = list(i)(0)
  For j = 0 To 3
    objRez.Cells(i + 2, j + 2).Value = list(i)(1)(j + 1)
  Next j
Next i

Protokol = tbCommandList.Text & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & Protokol

' ��������� ������������ ��������� �������� � ����
On Error GoTo ErrorHandler
  filePRT = Application.GetSaveAsFilename("���������������� " & RootNode & " ����.prt", "����� ��������� ��� (*.prt), *.prt")
  If filePRT <> "False" Then
    Dim FrFi As Integer
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, Protokol
    Close FrFi
  End If

cbFindNode.Enabled = True
If Not cbMessages.Value Then
  MsgBox "������� ����� �� ��� �������� ����������������� ����������� ������ ������������, " & _
    "���������� � ������ �������� �����"
End If

Exit Sub
ErrorHandler:
   MsgBox ("� �������� ���������� ����� ��������� �������� ������, �������� �������� �� ��������!")

End Sub


Private Function Find_BaseRejim_Name(RejimNo)
'
' ���� ������������ ������ �� ������� �������, ���������� �� ���������� ����
'

Find_BaseRejim_Name = ""
For i = 0 To UBound(arrBaseRejims)
  If arrBaseRejims(i)(0) = Int(RejimNo) Then
    Find_BaseRejim_Name = arrBaseRejims(i)(1)
    Exit For
  End If
Next

End Function


Private Function Is_Sub_Rejim(Protokol, CurPos)
'
' �������� �������� ����� �� ������ � ���������, ���������� �� ������ ������ (������ �� �������� ����)
'

Is_Sub_Rejim = False
apos = InStr(CurPos, Protokol, "��������  ")
If apos = 0 Then GoTo skeep
bpos = InStr(apos + 10, Protokol, " ")
cpos = InStr(apos + 10, Protokol, vbCrLf)
If cpos > bpos Then Is_Sub_Rejim = True

skeep:
End Function


Private Function Get_Rejim_Name(Protokol, CurPos)
'
' ���� ����� �� ������� ��������� - ����� ��� ������������,
' ���� ��� �������� - ����� ������������ ������������ ��������
'

' ������ ����� ������, ���� ��� �������� - ����� ������� �� ���� �����, ���� �������� - ������

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
    Get_Rejim_Name = "+���� " & Trim(Mid(Protokol, apos + 1, bpos - apos - 1))
  End If
Else
  apos = InStr(CurPos + 10, Protokol, vbCrLf)
  Get_Rejim_Name = Find_BaseRejim_Name(Trim(Mid(Protokol, CurPos + 10, apos - (CurPos + 10))))
End If
Get_Rejim_Name = "[" & Prefix & "] " & Get_Rejim_Name
  
End Function


Private Sub cbAnaliz2_Click()
'
' ������� ��������� �� �����������
'

' ���������� ����� �������� � ��� �� ����, ��� � ���������� �� �������� ���������������� � ����������� ������
Set objWorkbook = ActiveWorkbook
Set objRez = objWorkbook.ActiveSheet
s = objRez.UsedRange.Rows.Count + 1

' ����� ����� ��������� �� ������ ������
Dim d As New DataObject
d.GetFromClipboard
Protokol = d.GetText

' ��������� �� ����������
StartPos = InStr(Protokol, "� � � � � � � � � �    � � � � � � �")

Dim list()
j = 0
StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)

Do
  ReDim Preserve list(j)
  RejimName = Get_Rejim_Name(Protokol, StartPos)
  Line = Parse_Current_Line(Protokol, StartPos, StartPos)
      
  list(j) = Array(RejimName, Line)
  j = j + 1
  StartPos = InStr(StartPos, Protokol, "��������  " & j + 1)
  DoEvents
Loop While StartPos > 0

' ������ ������������ ������� � ��������������� ����� ��������, ��������� ��� �� ����
s = s + 1
objRez.Cells(s, 1).Value = "��� ��� ������ �����������"
For i = 0 To UBound(list)
  objRez.Cells(s + i + 1, 1).Value = list(i)(0)
  For j = 0 To 3
    objRez.Cells(s + i + 1, j + 2).Value = list(i)(1)(j + 1)
  Next j
Next i

Protokol = tbCommandList.Text & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & Protokol

' ��������� ������������ ��������� �������� � ����
On Error GoTo ErrorHandler
  filePRT = Application.GetSaveAsFilename("����������� " & RootNode & " ����.prt", "����� ��������� ��� (*.prt), *.prt")
  If filePRT <> "False" Then
    Dim FrFi As Integer
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, Protokol
    Close FrFi
  End If

Exit Sub
ErrorHandler:
   MsgBox ("� �������� ���������� ����� ��������� �������� ������, �������� �������� �� ��������!")

End Sub


Private Function Delete_Interm_Nodes(Without)
'
' �������� ������������� �����, �.�. ��� � ������� ������ ��� ��������������
' ������� �� ����������� ������� �������� ������������� �/�������, �� ����� ��������
' �� ��� ���, ���� �� ����� ������� ������� ��������� (������� ���������� ����������
' ��������� �����
'

j = 0
For i = 1 To UBound(arrNode)
  Node = arrNode(i, 1)
  If Node <> Without Then
    NodeBranch = Find_Branch_By_Node(arrBranchCopy, Node)
    On Error Resume Next
    n = -1
    n = UBound(NodeBranch)
    If err = 0 Then
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
          ' ����� �������� ��������������� �� Without (��� RootNode), �����
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
  DoEvents
Next
Delete_Interm_Nodes = j

End Function


Private Function Find_Element_By_2Node(Branch, Node1, Node2)
'
' ���� ������ ���������, ������� ����� � ������ ������� ���� Node1 � Node2
'

j = 0
Dim rez()
For i = 1 To UBound(Branch)
  If ((Branch(i, 3) = Node1) And (Branch(i, 4) = Node2)) Or _
     ((Branch(i, 3) = Node2) And (Branch(i, 4) = Node1)) Then
      ReDim Preserve rez(j)
      rez(j) = Branch(i, 5)
      j = j + 1
  End If
Next
Find_Element_By_2Node = rez

End Function


Private Sub cbFindNode_Click()
'
' ������ ������� ��������� � ���, ��� ���������� ����������� �������� ����
' ���� ��� �� ������ ������������� �������, �� ������� ��� �� ������� :)
'

If Not Initialized Then Call Initialize
If RootNode = 0 Then RootNode = Int(InputBox("����� ���� (�������������� ����)?", "RootNode"))

' ������� ����� ����� � 101 ����� (����������� ���)
For i = 1 To UBound(arrBranchCopy)
  If arrBranchCopy(i, 1) = 101 Then
    arrBranchCopy(i, 3) = 0
    arrBranchCopy(i, 4) = 0
  End If
  arrBranchCopy(i, 5) = 0
  arrBranchCopy2(i, 5) = 0
Next

' ������� �� �������������� RootNode � ��������� ��������� �� ���� ������ ���������� �����
' ��������. ��� �������� ������������� ���������� � Delete_Interm_Nodes()
' ���� ���� �� ����� ����� = RootNode ��� ����� �������� ����� ��������������� �� �����
' ������������ �����
NodeBranch = Find_Branch_By_Node(arrBranchCopy, RootNode)
For i = 0 To UBound(NodeBranch)
  arrBranchCopy(NodeBranch(i), 5) = i + 1
  arrBranchCopy2(NodeBranch(i), 5) = i + 1
Next

Do
  j = 0
  ' ������� ������������� ���� � ������ (���� ������ � ����� ������)
   n = Delete_Interm_Nodes(RootNode)
   j = j + n

  ' ������� ������ � ���� ��������� � ��-��� �� ���� (���) �� �� ����������
  For i = 1 To UBound(arrBranchCopy)
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
j = 0
NodeBranch = Find_Branch_By_Node(arrBranchCopy, RootNode)
On Error Resume Next
n = UBound(NodeBranch)
If err = 0 Then
  For i = 0 To n
    If arrBranchCopy(NodeBranch(i), 3) = RootNode Then
      DestNode = arrBranchCopy(NodeBranch(i), 4)
    Else
      DestNode = arrBranchCopy(NodeBranch(i), 3)
    End If
    ' �� ��������� ���������, ������� ����� ��������� �� �� ������������ �� ������� (������)
    r = False
    On Error Resume Next
    kn = UBound(list)
    If err = 0 Then
      For k = 0 To kn
        If list(k) = DestNode Then
          r = True
          Exit For
        End If
      Next k
    Else
      r = False
    End If
    
    If (Not r) And (DestNode <> 0) Then
      ReDim Preserve list(j)
      list(j) = DestNode
      j = j + 1
    End If
  Next i
End If

' ������� ������ �������� ����� � ������ ����� ������������� �� ���
tbBranchList.Text = ""
NodeBranch = Find_Branch_By_Node(arrBranchCopy2, RootNode)
For i = 0 To UBound(list)
  PNode = list(i)
  ' ������ �����(�) ���������, � ������� ������ RootNode � PNode, ���� � ��������� ���� ���� �� ���� ����
  ' ���� ��������� ����� ���� ���������
  Elem = Find_Element_By_2Node(arrBranchCopy, RootNode, PNode)
  On Error Resume Next
  n = UBound(Elem)
  If err = 0 Then
    ' ������ ����� ������������� RootNode ������������� � ��������� Elem
    For j = 0 To n
      e = Elem(j)
      For k = LBound(NodeBranch) To UBound(NodeBranch)
        If arrBranchCopy2(NodeBranch(k), 5) = e Then
          If arrBranchCopy2(NodeBranch(k), 3) = RootNode Then
            SecondNode = arrBranchCopy2(NodeBranch(k), 4)
          Else
            SecondNode = arrBranchCopy2(NodeBranch(k), 3)
          End If
          n_node = Trim(Find_Node(list(i)))
          n_branch = Find_Branch_By_2Node(arrBranch, RootNode, SecondNode)
          n_branch = arrBranch(n_branch, 5)
          n_branch = Trim(Find_Element(n_branch))
          tbBranchList.Text = tbBranchList.Text & PNode & vbTab & "(" & RootNode & "-" & SecondNode & ")" & vbTab & "/* " & n_node & " [" & n_branch & "]" & vbCrLf
        End If
      Next
    Next
  End If
Next
'
'  SecondNode = Path(UBound(Path))
'  tbBranchList.Text = tbBranchList.Text & list(i) & " (" & RootNode & "-" & SecondNode & ") /*" & Find_Node(list(i)) & vbCrLf
'Next

' ������� ��������������� ���������
cbProcess2.Enabled = True
Label1.Caption = "����������� �������� ����: �����_���� (���1-���2) - ����� � ��������� ����"
If Not cbMessages.Value Then
  MsgBox "������������� ������� �������� ����, ���������, ������� ���/����, " & _
    "����� ��������� ����� � �������� �����"
End If

End Sub
