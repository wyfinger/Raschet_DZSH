VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDZSH 
   Caption         =   "Расчет ДЗШ / МИВ, 2013-08-09"
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

Dim RootNode As Integer    ' самый главный узел :)

Dim arrBranch()
Dim arrBranchCopy()
Dim arrBranchCopy2()
Dim arrNode()
Dim arrElement()
Dim Initialized
Dim arrTrueBrach()  ' список присоединений узла, кроме неотключаемых
Dim arrBaseRejims() ' НОМЕР, Название для базовых режимов, нужно при парсинге протокола по опробованию

Const vbTab = "   "   ' Этот дебильный АРМ затыкается на некоторых приказах с табом



Private Sub Initialize()
'
' Подготовка - изымаем данные с листа в переменные
'

' 1. Ветви - берем первые пять столбцов страницы 'Таблица ветвей'
Set objWorkbook = ActiveWorkbook
Set wshBranch = ActiveWorkbook.Worksheets("Таблица ветвей")
arrBranch = wshBranch.Range("A3:K" & wshBranch.UsedRange.Rows.Count).Value2
arrBranchCopy = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2
arrBranchCopy2 = wshBranch.Range("A3:E" & wshBranch.UsedRange.Rows.Count).Value2

' 2. Узлы - берем первые три столбца страницы 'Наим.узлов'
Set wshNode = objWorkbook.Worksheets("Наим.узлов")
arrNode = wshNode.Range("A3:E" & wshNode.UsedRange.Rows.Count).Value2

' 3. Элементы - берем первые два столбца страницы 'Наим.элементов'
Set wshElement = objWorkbook.Worksheets("Наим.элементов")
arrElement = wshElement.Range("A3:B" & wshElement.UsedRange.Rows.Count).Value2

Initialized = True

End Sub


Private Function Find_Branch_By_Node(BranchArray, Node)
'
' Ищем все ветви, в которые входит заданый узел
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
' Ищем одну ветвь по имеющимся двум узлам, ее образующим, возвращаем номер ветви
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
' Ищем наименование узла по номеру
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
' Ищем индекс узла по номеру
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
' Проверка существования узла
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
' Ищем наименование элемента по номеру
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
' Ищем индекс ветви по узлам слева и справа, порядок узлов не имеет значения
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
' Подготовка приказа проверки чувствительности
'

' Берем питающие узлы, указанные пользователем
arrLines = Split(frmDZSH.tbBranchList.Text, vbCrLf)
Dim list()
j = 0
For i = 0 To UBound(arrLines)
  ' Удаление комментариев
  cpos = InStr(arrLines(i), "/*")
  If cpos > 0 Then arrLines(i) = Mid(arrLines(i), 1, cpos - 1)
  On Error GoTo skeep
  ' Вычленяем номер питающего узла и ветвь-присоединение, ведущее к нему
  apos = InStr(arrLines(i), "(") + 1
  bpos = InStr(arrLines(i), "-") + 1
  cpos = InStr(arrLines(i), ")") + 1
  PowerNode = Int(Trim(Mid(arrLines(i), 1, apos - 3)))
  NodeA = Int(Trim(Mid(arrLines(i), apos, bpos - apos - 1)))
  NodeB = Int(Trim(Mid(arrLines(i), bpos, cpos - bpos - 1)))
  ' Заносим все в список
  ReDim Preserve list(j)
  list(j) = Array(PowerNode, NodeA, NodeB)
  j = j + 1
skeep:
Next

' Подготавливаем шапку приказа
frmDZSH.tbCommandList = _
"ВЕЛИЧИНА  IA IB IC" & vbCrLf & _
"1-ПОЯС    " & RootNode & vbTab & "/* " & Find_Node(RootNode) & vbCrLf & _
"СНСМ      1" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/ABC" & vbCrLf & _
"СНСМ      2" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/AB" & vbCrLf & _
"СНСМ      3" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/AB0" & vbCrLf & _
"СНСМ      4" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/A0" & vbCrLf

Podrejim = 1
k = 0
' Проходим по всем присоединениям, указанным в списке (присоединения к питающему узлу)
For i = 0 To UBound(list)
   PowerNode = list(i)(0)
   NodeA = list(i)(1) ' RootNode
   NodeB = list(i)(2) ' номер противоположного узла первой ветви присоединения к питающему узлу
   frmDZSH.tbCommandList = frmDZSH.tbCommandList & vbCrLf
   frmDZSH.tbCommandList = frmDZSH.tbCommandList & "ПОДРЕЖИМ  " & Podrejim & " /* " & PowerNode & " [" & Find_Node(PowerNode) & "]" & vbCrLf
   BaseRejim = Podrejim
   ' Запишем название базового режима, чтобы можно было это вписать в результирующий лист,
   ' из протокола эту информацию не достать
   ReDim Preserve arrBaseRejims(i)
   branchNo = Find_Branch_By_2Node(arrBranch, NodeA, NodeB)
   elemName = Find_Element(arrBranch(branchNo, 5))
   arrBaseRejims(i) = Array(Podrejim, Find_Node(PowerNode) & " [" & elemName & "]")
      
   ' Пройдемся по всем отключаемым от ДЗШ присоединениям (задается пользователем после 1 шага)
   For j = 0 To UBound(arrTrueBrach)
     ' Отключаем все присоединения, кроме текущего
     If arrTrueBrach(j)(2) <> NodeB Then
    '   frmDZSH.tbCommandList = frmDZSH.tbCommandList & "ОТКЛ      *" & arrTrueBrach(j)(1) & "-" & arrTrueBrach(j)(2) & _
    '     " /* " & Find_Node(arrTrueBrach(j)(1)) & " - " & Find_Node(arrTrueBrach(j)(2)) & vbCrLf
       frmDZSH.tbCommandList = frmDZSH.tbCommandList & "ОТКЛ      *" & arrTrueBrach(j)(1) & "-" & arrTrueBrach(j)(2) & _
         " /* " & Find_Element(arrTrueBrach(j)(0)) & vbCrLf
     Else ' для справки
       frmDZSH.tbCommandList = frmDZSH.tbCommandList & "* ОТКЛ      *" & arrTrueBrach(j)(1) & "-" & arrTrueBrach(j)(2) & _
         " /* " & Find_Node(arrTrueBrach(j)(1)) & " - " & Find_Node(arrTrueBrach(j)(2)) & vbCrLf
     End If
   Next
   
   ' Найдем все присоединения питающего узла и отключим каждое в отдельном подрежиме, основанном на BaseRejim
   NodeBranch = Find_Branch_By_Node(arrBranch, PowerNode)
   ' Issue#2: Проверяем, чтобы не отключить что-нибудь дважды. Найдем присоединения RootNode
   RootNodeBranch = Find_Branch_By_Node(arrBranch, RootNode)
   For j = 0 To UBound(NodeBranch)
     T = arrBranch(NodeBranch(j), 1)
     If T <> 101 Then
       Podrejim = Podrejim + 1
       frmDZSH.tbCommandList = frmDZSH.tbCommandList & "ПОДРЕЖИМ  " & Podrejim & " " & BaseRejim & vbCrLf
       NodeA = arrBranch(NodeBranch(j), 3)
       NodeB = arrBranch(NodeBranch(j), 4)
       
       ' Issue#2: Проверим, что ветвь, отходящая от питающего узла, котороую мы хотим отключить,
       ' не связана с RootNode (все ветви RootNode кроме одной отключены в базовом режиме)
       Elem = arrBranch(NodeBranch(j), 5)    ' Номер элемента той ветви от PowerNode, которую хотим отключить
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
           frmDZSH.tbCommandList = frmDZSH.tbCommandList & "ОТКЛ      *" & PowerNode & "-" & ContrNode & _
             " /* " & Find_Node(PowerNode) & " - " & Find_Node(ContrNode) & vbCrLf
         Else
           frmDZSH.tbCommandList = frmDZSH.tbCommandList & "ЭЛЕМЕНТ   " & Elem & _
             " /* " & Find_Element(Elem) & vbCrLf
         End If
       End If
     End If
   Next
   Podrejim = Podrejim + 1
   
Next

' Копируем приказ в буфер
Dim d As New DataObject
d.SetText (frmDZSH.tbCommandList.Text)
d.PutInClipboard

cbAnaliz2.Enabled = True
If Not cbMessages.Value Then
  MsgBox "Приказ для ТКЗ для проверки режима опробования (чувствительный огран) подготовлен " & _
    "и скопирован в буфер обмена, необходимо очистить окно протокола АРМ ТКЗ, вставить приказ и выполнить расчет " & _
    "после чего скопировать весь отчет и выполнить п.6 (анализ результатов расчета)."
End If

End Sub

' #################################################################################

Function Find_TKZ_Handle() As Long
'
' Поиск главного окна ТКЗ-2000
'



End Function

' #################################################################################



Private Sub CommandButton1_Click()
'
' Весь рабочий процесс
'

' Инициализация, берем данные с листа
If Not Initialized Then Call Initialize

If RootNode = 0 Then RootNode = Int(InputBox("Номер узла (рассчитываемые шины)?", "RootNode", 0))

' А не ошибся ли пользователь?
If Not Node_Exists(RootNode) Then
  MsgBox "Узел " & RootNode & " отсутствует в таблице узлов", vbOKOnly + vbExclamation
  Exit Sub
End If

' Найдем окно TKZ, если его нет - завершаемся с сообщением об ошибке
TKZ_Handle& = FindWindow("TFormZD2", "")

End Sub

Private Sub UserForm_Initialize()
'
' Загрузка настроек программы (галка о сообщениях)
'

cbMessages.Value = GetSetting("Raschet_DZSH", "Settings", "ShowMessages", True)

End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'
' Обработчик закрытия формы, сохраняем настройки приложения
'

SaveSetting "Raschet_DZSH", "Settings", "ShowMessages", cbMessages.Value

End Sub


Private Sub cbFindBranch_Click()
'
' Ищем все ветки от заданного узла, находим для каждой название, и помещаем в tbBranchList
'

' Инициализация, берем данные с листа

If Not Initialized Then Call Initialize
If RootNode = 0 Then RootNode = Int(InputBox("Номер узла (рассчитываемые шины)?", "RootNode", 0))

If Not Node_Exists(RootNode) Then
  MsgBox "Узел " & RootNode & " отсутствует в таблице узлов"
  Exit Sub
End If

' Подготавливаем список смежных ветвей, для редактирования (неотключаемые присоединения)
BranchList = Find_Branch_By_Node(arrBranch, RootNode)
frmDZSH.tbBranchList.Text = "/* УЗЕЛ " & RootNode & " (" & Find_Node(RootNode) & ") - ПРИСОЕДИНЕНИЯ: ЭЛЕМЕНТ (ВЕТВЬ)" & vbCrLf
For i = 0 To UBound(BranchList)
  T = Int(arrBranch(BranchList(i), 1))       ' Тип ветви
  A = Int(arrBranch(BranchList(i), 3))       ' Узел начала
  b = Int(arrBranch(BranchList(i), 4))       ' Узел конца
  e = Int(arrBranch(BranchList(i), 5))       ' Номер элемента
  If T <> 101 Then
    If A = RootNode Then Сross = b Else Сross = A  ' Номер узла с противоположной стороны ветви
    ' Если номер элемента 0 и противоположный узел - тоже ноль - пишем в комментарии, что это нейтраль
    If (Сross = 0) And (e = 0) Then            ' !!!!!!
      str2add = e & vbTab & "(" & RootNode & "-" & Сross & ")" & vbTab & "/* " & "НЕЙТРАЛЬ"
    Else
      str2add = e & vbTab & "(" & RootNode & "-" & Сross & ")" & vbTab & "/* " & Find_Element(e)
    End If
    frmDZSH.tbBranchList.Text = frmDZSH.tbBranchList.Text & str2add & vbCrLf
  End If
Next

cbProcess1.Enabled = True
Label1.Caption = "Присоединения узла, удалить ветви, которые не могут быть отключены :"
If Not cbMessages.Value Then
  MsgBox "Подготовлен список ветвей, примыкающих к выбранному узлу, теперь этот список нужно отредактировать " & _
    "(удалить ветви, которые по режиму не могут быть отключены), после чего выполнить п.2"
End If

End Sub


Private Sub cbProcess1_Click()

' Читаем каждую строку, вычленяем номер элемента и начало и конец ветви
' нужно для отключения нейтралей
arrLines = Split(frmDZSH.tbBranchList.Text, vbCrLf)
j = 0
For i = 0 To UBound(arrLines)
  ' Удаление комментариев
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

' Подготавливаем приказ
frmDZSH.tbCommandList = _
"ВЕЛИЧИНА  IA IB IC" & vbCrLf & _
"1-ПОЯС    " & RootNode & vbTab & "/* " & Find_Node(RootNode) & vbCrLf & _
"СНСМ      1" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/ABC" & vbCrLf & _
"СНСМ      2" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/AB" & vbCrLf & _
"СНСМ      3" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/AB0" & vbCrLf & _
"СНСМ      4" & vbCrLf & _
"ЗАМ-ФАЗ   " & RootNode & "/A0" & vbCrLf & _
"ПОДРЕЖИМ  1" & vbTab & "/* ВСЕ ВКЛЮЧЕНО" & vbCrLf

' Для каждого из присоединений, которые перечислил пользователь, создаем подрежим
For i = 0 To UBound(arrTrueBrach)
  If arrTrueBrach(i)(1) = RootNode Then Сross = arrTrueBrach(i)(2) Else Сross = arrTrueBrach(i)(1)
  Elem = Find_Element(arrTrueBrach(i)(0))
  If (Сross = 0) Or (Elem = 0) Then               ' отключаем ветвь
    str2add = "ПОДРЕЖИМ  " & (i + 2) & vbCrLf & _
    "ОТКЛ      0 *" & RootNode & "-" & Сross & vbTab & "/* НЕЙТРАЛЬ ??"
  Else                                                                 ' отключаем элемент
    str2add = "ПОДРЕЖИМ  " & (i + 2) & vbCrLf & _
    "ЭЛЕМЕНТ   " & arrTrueBrach(i)(0) & vbTab & "/* " & Find_Element(arrTrueBrach(i)(0))
  End If
  
  frmDZSH.tbCommandList = frmDZSH.tbCommandList & str2add & vbCrLf
Next

' Копируем приказ в буфер
Dim d As New DataObject
d.SetText (frmDZSH.tbCommandList.Text)
d.PutInClipboard

cbAnaliz1.Enabled = True
If Not cbMessages.Value Then
  MsgBox "Приказ для ТКЗ для проверки чувствительности пусковых и избирательных органов ДЗШ подготовлен " & _
    "и скопирован в буфер обмена, необходимо очистить окно протокола АРМ ТКЗ, вставить приказ и выполнить расчет " & _
    "после чего скопировать весь отчет и выполнить п.3 (анализ результатов расчета)."
End If

End Sub


Private Function Parse_Current_Line(T, FromPos, ByRef FinishPos)
'
' Функция ищет суммарные токи КЗ в несимметриях в тексте протокола Т начиная с
' FromPos, в параметр FinishPos записывается позиция, на которой остановился парсер
'

Dim Imin(1 To 4)
For i = 1 To 4
  FinishPos = InStr(FromPos, T, "СНСМ      " & i)
  FinishPos = InStr(FinishPos, T, "IАсум")
  ia = Int(Mid(T, FinishPos + 5, 10))
  FinishPos = InStr(FinishPos, T, "IВсум")
  IB = Int(Mid(T, FinishPos + 5, 10))
  If (IB < ia) And (IB > 0) Then Imin(i) = IB Else Imin(i) = ia
Next
Parse_Current_Line = Imin

End Function


Private Function Parse_Rezhim_Single(T, FromPos)
'
' Ищем что отключается в подрежиме, здесь обрабатывается только одно отключение
'

apos = InStr(FromPos, T, "Подрежим")
bpos = InStr(FromPos, T, vbCrLf)
If (apos > 0) And (bpos > apos) Then
  Prefix = Trim(Mid(T, apos + 8, bpos - apos - 8))
End If

pa = InStr(FromPos, T, "(") + 1
pb = InStr(FromPos, T, ")")
pc = InStr(FromPos, T, "СНСМ")
If (pa > 0) And (pb > 0) And (pc > 0) Then
  If pc > pb Then
    Parse_Rezhim_Single = "-" & Trim(Mid(T, pa, pb - pa))
  Else
    Parse_Rezhim_Single = "КЗ на " & RootNode & ", ВСЕ ВКЛЮЧЕНО"
  End If
End If
Parse_Rezhim_Single = "[" & Prefix & "] " & Parse_Rezhim_Single

End Function


Private Sub cbAnaliz1_Click()
'
' Разбор протокола АРМ ТКЗ и концентрация токов в отдельном листе
'

' Добавляем лист для результатов
Set objWorkbook = ActiveWorkbook
Set objRez = objWorkbook.Worksheets.Add
objRez.Columns("A:A").ColumnWidth = 35#
'objRez.Name = RootNode & " " & Find_Node(RootNode)

' Берем текст протокола из буфера обмена
Dim d As New DataObject
d.GetFromClipboard
Protokol = d.GetText

' Пройдемся по подрежимам
StartPos = InStr(Protokol, "Р Е З У Л Ь Т А Т Ы    Р А С Ч Е Т А")
Dim list()
j = 0
StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)

' Найдем что отключалось в этом подрежиме, в данном случае это только одна ветвь или элемент
r = Parse_Rezhim_Single(Protokol, StartPos)

Do
  ReDim Preserve list(j)
  list(j) = Array(r, Parse_Current_Line(Protokol, StartPos, StartPos))
  j = j + 1
  StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)
  If StartPos > 0 Then r = Parse_Rezhim_Single(Protokol, StartPos)
  DoEvents
Loop While StartPos > 0

' Массив наименований режимов и соответствующих токов заполнен, переносим его на лист
objRez.Cells(1, 1).Value = "Узел " & RootNode & " (" & Find_Node(RootNode) & ") - ТКЗ для чувствительности пуск. и изб. органов"
For i = 0 To UBound(list)
  objRez.Cells(i + 2, 1).Value = list(i)(0)
  For j = 0 To 3
    objRez.Cells(i + 2, j + 2).Value = list(i)(1)(j + 1)
  Next j
Next i

Protokol = tbCommandList.Text & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & Protokol

' Предложим пользователю сохранить протокол в файл
On Error GoTo ErrorHandler
  filePRT = Application.GetSaveAsFilename("Чувствительность " & RootNode & " узел.prt", "Файлы протокола АРМ (*.prt), *.prt")
  If filePRT <> "False" Then
    Dim FrFi As Integer
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, Protokol
    Close FrFi
  End If

cbFindNode.Enabled = True
If Not cbMessages.Value Then
  MsgBox "Таблица токов КЗ для проверки чувствительностив минимальном режиме подготовлена, " & _
    "переходите к поиску питающих узлов"
End If

Exit Sub
ErrorHandler:
   MsgBox ("В процессе сохранения файла протокола возникла ошибка, возможно протокол не сохранен!")

End Sub


Private Function Find_BaseRejim_Name(RejimNo)
'
' Ищем наименование режима из базовых режимов, занесенных на предыдущем шаге
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
' Проверка является стоим ли сейчас в подрежиме, основанном на другом режиме (ремонт на питающем узле)
'

Is_Sub_Rejim = False
apos = InStr(CurPos, Protokol, "Подрежим  ")
If apos = 0 Then GoTo skeep
bpos = InStr(apos + 10, Protokol, " ")
cpos = InStr(apos + 10, Protokol, vbCrLf)
If cpos > bpos Then Is_Sub_Rejim = True

skeep:
End Function


Private Function Get_Rejim_Name(Protokol, CurPos)
'
' Если стоим на базовом подрежиме - берем его наименование,
' если это субрежим - берем наименование отключаемого элемента
'

' Найдем номер режима, если это подрежим - номер состоит из двух чисел, если основной - одного

apos = InStr(CurPos, Protokol, "Подрежим")
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
    Get_Rejim_Name = "+Откл " & Trim(Mid(Protokol, apos + 1, bpos - apos - 1))
  End If
Else
  apos = InStr(CurPos + 10, Protokol, vbCrLf)
  Get_Rejim_Name = Find_BaseRejim_Name(Trim(Mid(Protokol, CurPos + 10, apos - (CurPos + 10))))
End If
Get_Rejim_Name = "[" & Prefix & "] " & Get_Rejim_Name
  
End Function


Private Sub cbAnaliz2_Click()
'
' Парсинг протокола по опробованию
'

' Результаты быдем выводить в тот же лист, что и результаты по проверке чувствительности в минимальном режиме
Set objWorkbook = ActiveWorkbook
Set objRez = objWorkbook.ActiveSheet
s = objRez.UsedRange.Rows.Count + 1

' Берем текст протокола из буфера обмена
Dim d As New DataObject
d.GetFromClipboard
Protokol = d.GetText

' Пройдемся по подрежимам
StartPos = InStr(Protokol, "Р Е З У Л Ь Т А Т Ы    Р А С Ч Е Т А")

Dim list()
j = 0
StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)

Do
  ReDim Preserve list(j)
  RejimName = Get_Rejim_Name(Protokol, StartPos)
  Line = Parse_Current_Line(Protokol, StartPos, StartPos)
      
  list(j) = Array(RejimName, Line)
  j = j + 1
  StartPos = InStr(StartPos, Protokol, "Подрежим  " & j + 1)
  DoEvents
Loop While StartPos > 0

' Массив наименований режимов и соответствующих токов заполнен, переносим его на лист
s = s + 1
objRez.Cells(s, 1).Value = "ТКЗ для режима опробования"
For i = 0 To UBound(list)
  objRez.Cells(s + i + 1, 1).Value = list(i)(0)
  For j = 0 To 3
    objRez.Cells(s + i + 1, j + 2).Value = list(i)(1)(j + 1)
  Next j
Next i

Protokol = tbCommandList.Text & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & Protokol

' Предложим пользователю сохранить протокол в файл
On Error GoTo ErrorHandler
  filePRT = Application.GetSaveAsFilename("Опробование " & RootNode & " узел.prt", "Файлы протокола АРМ (*.prt), *.prt")
  If filePRT <> "False" Then
    Dim FrFi As Integer
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, Protokol
    Close FrFi
  End If

Exit Sub
ErrorHandler:
   MsgBox ("В процессе сохранения файла протокола возникла ошибка, возможно протокол не сохранен!")

End Sub


Private Function Delete_Interm_Nodes(Without)
'
' Удаление промежуточных узлов, т.е. тех у которых только два присоединиения
' Функция не гарантирует полного удаления промежуточных п/станций, ее нужно вызывать
' до тех пор, пока не быдет сделано никаких изменений (функция возвращает количество
' удаленных узлов
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
          ' Номер элемента распостраняется от Without (это RootNode), чтобы
          ' потом можно ыбло легко определить присоединение от RootNode, идущее к питающему узлу
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
  DoEvents
Next
Delete_Interm_Nodes = j

End Function


Private Function Find_Element_By_2Node(Branch, Node1, Node2)
'
' Ищем список элементов, которые имеют в совоем составе узлы Node1 и Node2
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
' Просто выводим сообщение о том, что необходимо перечислить питающие узлы
' Идеи как их искать автоматически имеются, но оставим это на будущее :)
'

If Not Initialized Then Call Initialize
If RootNode = 0 Then RootNode = Int(InputBox("Номер узла (рассчитываемые шины)?", "RootNode"))

' Удаляем сразу ветви с 101 типом (отключенный ШСВ)
For i = 1 To UBound(arrBranchCopy)
  If arrBranchCopy(i, 1) = 101 Then
    arrBranchCopy(i, 3) = 0
    arrBranchCopy(i, 4) = 0
  End If
  arrBranchCopy(i, 5) = 0
  arrBranchCopy2(i, 5) = 0
Next

' Пройдем по присоединениям RootNode и проставим отходящим от него ветвям уникальный номер
' элемента. При удалении промежуточных подстанций в Delete_Interm_Nodes()
' если один из узлов ветви = RootNode его номер элемента будет распостраняться на вновь
' образованную ветвь
NodeBranch = Find_Branch_By_Node(arrBranchCopy, RootNode)
For i = 0 To UBound(NodeBranch)
  arrBranchCopy(NodeBranch(i), 5) = i + 1
  arrBranchCopy2(NodeBranch(i), 5) = i + 1
Next

Do
  j = 0
  ' Удаляем промежуточные узлы и тупики (узлы только с одной ветвью)
   n = Delete_Interm_Nodes(RootNode)
   j = j + n

  ' Удаляем тупики в виде нейтралей и тр-ров на ноль (ТСН) но не генераторы
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

' Ищем противоположные узлы
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
    ' Не добавляем дубликаты, которые могут появиться из за шутнирования СВ линиями (кольца)
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

' Выводим номера питающих узлов и первую ветвь присоединения до них
tbBranchList.Text = ""
NodeBranch = Find_Branch_By_Node(arrBranchCopy2, RootNode)
For i = 0 To UBound(list)
  PNode = list(i)
  ' Найдем номер(а) элементов, в которые входят RootNode и PNode, если к питающему узлу удет не одна цепь
  ' этих элементов может быть несколько
  Elem = Find_Element_By_2Node(arrBranchCopy, RootNode, PNode)
  On Error Resume Next
  n = UBound(Elem)
  If err = 0 Then
    ' Найдем среди присоединений RootNode присоединение с элементом Elem
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

' Выводим диагностическое сообщение
cbProcess2.Enabled = True
Label1.Caption = "Перечислить питающие узлы: НОМЕР_УЗЛА (НОМ1-НОМ2) - ветвь к питающему узлу"
If Not cbMessages.Value Then
  MsgBox "Автоматически найдены питающие узлы, проверьте, удалите ТСН/РТСН, " & _
    "особо проверьте ветви к питающим узлам"
End If

End Sub
