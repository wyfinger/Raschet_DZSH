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

Dim GT_Class As String       ' Временная переменная класса для передачи в Enum-функцию
Dim GT_Result As Long
                                                                           
Dim RootNode As Long         ' Самый главный узел :)
Dim arrRootBranch()          ' Список ветвей главного узла, заполняется в Get_Sensitivity_Code()
Dim arrBranch()


Dim arrBranchCopy()
Dim arrBranchCopy2()
Dim arrNode()
Dim arrElement()
Dim arrTrueBrach()           ' Список присоединений узла, кроме неотключаемых
Dim arrBaseRejims()          ' НОМЕР, Название для базовых режимов, нужно при парсинге протокола по опробованию
Const vbTab = "   "          ' Этот дебильный АРМ затыкается на некоторых приказах с табом
                                                                           
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
      Exe_Name_by_Window_Handle = Trim(wt)
    End If
Finally:
    CloseHandle prc
  End If

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
  
  If IsWindowVisible(wnd) And (GetParent(wnd) = 0) Then
  
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


Private Function Find_TKZ_Window_Handle(ByVal class As String) As Long
'
' Поиск главного окна программы ТКЗ-2000
'

  GT_Result = 0
  GT_Class = class
  Call EnumWindows(AddressOf Find_Window_Enum_Proc, 0)
  Find_TKZ_Window_Handle = GT_Result

End Function


Private Function Find_SubClass_Recurce(hwnd As Long, sClassName As String, Optional iPos As Integer = 1) As Long
'
' Эта функция рекурсивно перебирает все дочерние окна hWnd, сверяя клас окна с sClassName
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

' Копируем приказ в буфер
Dim d As New DataObject
d.SetText (sText)
d.PutInClipboard

SendMessage hwnd, EM_SETSEL, 0, -1
SendMessage hwnd, WM_PASTE, 0, 0
  
End Function

Private Function Window_Get_Text(hwnd As Long)


SendMessage hwnd, EM_SETSEL, 0, -1
SendMessage hwnd, WM_CUT, 0, 0

' Копируем приказ в буфер
Dim d As New DataObject
d.GetFromClipboard
Window_Get_Text = d.GetText
  
End Function

'#################################################################################################[Расчет ДЗШ]

Private Sub Initialize()
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

End Sub


Private Function Find_Branch_By_Node(BranchArray, Node)
'
' Ищем все ветви, в которые входит заданный узел
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
' Ищем одну ветвь по имеющимся двум узлам, ее образующим, возвращаем номер ветви
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
' Ищем наименование узла по номеру
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
' Ищем индекс узла по номеру
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
' Проверка существования узла
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
' Ищем наименование элемента по номеру
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
' Ищем индекс ветви по узлам слева и справа, порядок узлов не имеет значения
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
' Подготовка приказа для оценки чувствительности ДЗШ
' Ищем все присоединения текущего узла, делаем КЗ на узле в нормальном режиме,
' а в подрежимах отключаем по одному присоединению
'

  Dim R As String

  R = _
"*         ПРОВЕРКА ЧУВСТВИТЕЛЬНОСТИ ДЗШ, УЗЕЛ " & RootNode & " [" & Find_Node(RootNode) & "]" & vbCrLf & _
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
"ПОДРЕЖИМ  1    /* ВСЕ ВКЛЮЧЕНО" & vbCrLf

  Dim CrossNode As Long
  Dim ElemNo As Long
  Dim ElemName As String
  Dim NewStr As String
  Dim i As Long

  ' Для каждого из присоединений RootNode создаем подрежим, где отключаем присоединения
  arrRootBranch = Find_Branch_By_Node(arrBranch, RootNode)
  For i = 0 To UBound(arrRootBranch)
  
    If arrBranch(arrRootBranch(i), 3) = RootNode Then
      CrossNode = arrBranch(arrRootBranch(i), 4)
    Else
      CrossNode = arrBranch(arrRootBranch(i), 3)
    End If
    ElemNo = arrBranch(arrRootBranch(i), 5)
    ElemName = Find_Element(ElemNo)
    If (CrossNode = 0) Or (ElemNo = 0) Then               ' отключаем ветвь
      NewStr = "ПОДРЕЖИМ  " & (i + 2) & vbCrLf & _
      "ОТКЛ      0 *" & RootNode & "-" & CrossNode & "      /* НЕЙТРАЛЬ ?"
    Else                                                  ' отключаем элемент
      NewStr = "ПОДРЕЖИМ  " & (i + 2) & vbCrLf & _
      "ЭЛЕМЕНТ   " & ElemNo & "      /* " & ElemName
    End If
  
    R = R & NewStr & vbCrLf
  
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
  For i = 0 To UBound(list)
    objRez.Cells(i + 3, 1).Value = list(i)(0)
    For j = 0 To 3
      objRez.Cells(i + 3, j + 2).Value = list(i)(1)(j + 1)
    Next j
  Next i

End Sub


Private Function Delete_Interm_Nodes(Without As Long) As Long
'
' Удаление промежуточных узлов, т.е. тех у которых только два присоединиения
' Функция не гарантирует полного удаления промежуточных п/станций, ее нужно вызывать
' до тех пор, пока не будет сделано никаких изменений (функция возвращает количество
' удаленных узлов
'

  Dim i, j, n As Long
  Dim Node As Long
  Dim NodeBranch()
  Dim NodePosA, NodePosB As Long
  Dim ContrNodeA, ContrNodeB As Long
  Dim DestType, DestElement As Long

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

Dim i, j, n As Long

' Удаляем сразу ветви с 101 типом (отключенный ШСВ)
For i = 1 To UBound(arrBranchCopy)
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

'DEL NodeBranch = Find_Branch_By_Node(arrBranchCopy, RootNode)
For i = 0 To UBound(arrRootBranch)
  arrBranchCopy(arrRootBranch(i), 5) = i + 1
  arrBranchCopy2(arrRootBranch(i), 5) = i + 1
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
    R = False
    On Error Resume Next
    kn = UBound(list)
    If err = 0 Then
      For k = 0 To kn
        If list(k) = DestNode Then
          R = True
          Exit For
        End If
      Next k
    Else
      R = False
    End If

    If (Not R) And (DestNode <> 0) Then
      ReDim Preserve list(j)
      list(j) = DestNode
      j = j + 1
    End If
  Next i
End If

End Sub

' Выводим номера питающих узлов и первую ветвь присоединения до них
'tbBranchList.Text = ""
'NodeBranch = Find_Branch_By_Node(arrBranchCopy2, RootNode)
'For i = 0 To UBound(list)
'  PNode = list(i)
'  ' Найдем номер(а) элементов, в которые входят RootNode и PNode, если к питающему узлу удет не одна цепь
'  ' этих элементов может быть несколько
'  Elem = Find_Element_By_2Node(arrBranchCopy, RootNode, PNode)
'  On Error Resume Next
'  n = UBound(Elem)
'  If err = 0 Then
'    ' Найдем среди присоединений RootNode присоединение с элементом Elem
'    For j = 0 To n
'      e = Elem(j)
'      For k = LBound(NodeBranch) To UBound(NodeBranch)
'        If arrBranchCopy2(NodeBranch(k), 5) = e Then
'          If arrBranchCopy2(NodeBranch(k), 3) = RootNode Then
'            SecondNode = arrBranchCopy2(NodeBranch(k), 4)
'          Else
'            SecondNode = arrBranchCopy2(NodeBranch(k), 3)
'          End If
'          n_node = Trim(Find_Node(list(i)))
'          n_branch = Find_Branch_By_2Node(arrBranch, RootNode, SecondNode)
'          n_branch = arrBranch(n_branch, 5)
'          n_branch = Trim(Find_Element(n_branch))
'          tbBranchList.Text = tbBranchList.Text & PNode & vbTab & "(" & RootNode & "-" & SecondNode & ")" & vbTab & "/* " & n_node & " [" & n_branch & "]" & vbCrLf
'        End If
'      Next
'    Next
'  End If
'Next
'
'  SecondNode = Path(UBound(Path))
'  tbBranchList.Text = tbBranchList.Text & list(i) & " (" & RootNode & "-" & SecondNode & ") /*" & Find_Node(list(i)) & vbCrLf
'Next

' Выводим диагностическое сообщение
'cbProcess2.Enabled = True
'Label1.Caption = "Перечислить питающие узлы: НОМЕР_УЗЛА (НОМ1-НОМ2) - ветвь к питающему узлу"
'If Not cbMessages.Value Then
'  MsgBox "Автоматически найдены питающие узлы, проверьте, удалите ТСН/РТСН, " & _
'    "особо проверьте ветви к питающим узлам"
'End If

'End Sub


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

  ' Ищем окно ТКЗ-2000, если его нет выводим сообщение и завершаемся
  MainFormHandle = Find_TKZ_Window_Handle("TForm1")
  If MainFormHandle = 0 Then
    MsgBox "Окно ТКЗ-2000 не найдено, приложение должно быть запущено. Кроме этого должна быть загружена сеть для расчета.", vbExclamation + vbOKOnly
    Exit Sub
  End If
  
  ' Инициализация, выбираем данные из листов
  Initialize
  
  ' Получаем номер узла (вообще-то узел может быть не только числовой)
  RootNode = Int(InputBox("Номер узла (рассчитываемые шины)?", "RootNode", 0))
  If Not Node_Exists(RootNode) Then
    MsgBox "Узел " & RootNode & " не найден в сети, дальнейшая работа невозможна.", vbExclamation + vbOKOnly
    Exit Sub
  End If
  
  ' Не проверяя в каком режиме работает программа (приказы или диалоговый расчет) выполним пункт меню
  ' "Расширенный формат задания для расчета..."
  Call SendMessage(MainFormHandle, WM_COMMAND, 12, 0&)

  ' Найдем окно диалога задания приказов
  CommandFormHandle = Find_TKZ_Window_Handle("TFormZD2")

  ' Откроем окно протокола и очистим его
  Call SendMessage(MainFormHandle, WM_COMMAND, 12, 0&)
  Call SendMessage(CommandFormHandle, WM_COMMAND, 187, 0&)
  ProtokolHandle = Find_TKZ_Window_Handle("TForm3")
  ProtokolMemo = Find_SubClass_Recurce(ProtokolHandle, "TMemo")
  Call SendMessage(ProtokolHandle, WM_COMMAND, 247, 0&)

  ' Готовим приказ для проверки чувствительности и копируем его в окно приказов ТКЗ-2000
  CommandsText = Get_Sensitivity_Code()
  CommandRichEdit = Find_SubClass_Recurce(CommandFormHandle, "TRichEdit")
  Window_Set_Text CommandRichEdit, CommandsText

  ' Делаем расчет с эквивалентированием - это значительно быстрее
  Call SendMessage(CommandFormHandle, WM_COMMAND, 179, 0&)

  ' Подождем секунду и заберем результат для анализа
  ProtokolText = Window_Get_Text(ProtokolMemo)

  ' Анализ протокола расчета
  Analiz_Sensitivity ProtokolText

  ' Предложить пользователю сохранить расширенный протокол (вначале добавлен исходный приказ с комментариями)
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText

  ' Предложим пользователю сохранить протокол в файл
  Dim filePRT
  Dim FrFi As Integer
  
  'filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\Чувствительность " & RootNode & " узел.prt", "Файлы протокола АРМ (*.prt), *.prt")
  'If filePRT <> "False" Then
  '  FrFi = FreeFile
  '  Open filePRT For Output As FrFi
  '  Print #FrFi, ProtokolText
  '  Close FrFi
  'End If
    
  ' Готовим приказ для проверки опробования
  Call Find_Power_Nodes
  
  
  
End Sub
