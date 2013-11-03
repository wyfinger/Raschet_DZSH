Attribute VB_Name = "Raschet_DZSH"
'
' Макрос ДЗШ
' ~~~~~~~~~~
' https://github.com/wyfinger/Raschet_DZSH
' Игорь Матвеев, miv@prim.so-ups.ru
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

Dim GT_Class As String       ' Временная переменная класса для передачи в Enum-функцию
Dim GT_Result As Long

Dim RootNode As Long         ' Самый главный узел :)
Dim arrRootBranch()          ' Список ветвей главного узла, заполняется в Get_Sensitivity_Code()

Dim arrBranch()              ' Массив ветвей и его копии
Dim arrBranchCopy()
Dim arrBranchCopy2()
Dim arrNode()                ' Массив наименований узлов
Dim arrElement()             ' Массив наименований элементов

Dim arrPowerNodes()          ' Массив питающих узлов (узел и первая ветвь от RootNode в сторону питающего узла)
                             ' заполняется в Find_Power_Nodes()
Dim arrBaseRejims()          ' Массив базовых режимов для проверки чувствительности,
                             ' заполняется при подготовке приказа в Get_Testing_Code()
                             ' т.к. из протокола расчета эту инфу не получить


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
' Забираем текст из поля ввода
'

  SendMessage hwnd, EM_SETSEL, 0, -1
  SendMessage hwnd, WM_CUT, 0, 0

  ' Копируем приказ в буфер
  Dim d As Object
  Set d = GetObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
  d.GetFromClipboard
  Window_Get_Text = d.GetText

  Set d = Nothing

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
' Ищем одну ветвь по имеющимся двум узлам, ее образующим, возвращаем номер ветви
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
' Ищем наименование узла по номеру
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
' Ищем индекс узла по номеру
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
' Проверка существования узла
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
' Ищем наименование элемента по номеру
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


Private Function Array_Find(Source(), Val, Optional Col As Integer = -1) As Integer
'
' Проверка содержания в массиве Source значения Val в столбце Col,
' если Col = -1 считаем, что массив одномерный.
' Возвращаем индекс элемента (первый) или -1, если не найдено
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

  Dim CrossNode As Long
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
      If (CrossNode = 0) Or (ElemNo = 0) Then               ' отключаем ветвь
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


Private Function Delete_Interm_Nodes(Without As Long) As Long
'
' Удаление промежуточных узлов, т.е. тех у которых только два присоединения
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
      ' Не добавляем дубликаты, которые могут появиться из за шутнирования СВ линиями (кольца)
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
            n_node = Trim(Find_Node(list(i)))                                       ' ???
            n_branch = Find_Branch_By_2Node(arrBranch, RootNode, SecondNode)        ' ???
            n_branch = arrBranch(n_branch, 5)                                       ' ???
            n_branch = Trim(Find_Element(n_branch))                                 ' ???
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
  Dim PowerNode, NodeA, NodeB, CrossNode, T As Long
  Dim branchNo As Long
  Dim ElemName As String
  Dim NodeBranch()
  Dim Elem, rElem As Long
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
    ElemName = Find_Element(arrBranch(branchNo, 5))
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
      Find_Element(arrBranch(arrRootBranch(j), 5)) & "), Ветвь (" & _
      Find_Node(arrBranch(arrRootBranch(j), 3)) & " - " & Find_Node(arrBranch(arrRootBranch(j), 4)) & _
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
    Prefix = Trim(Mid(Protokol, apos + 8, bpos - apos - 8))
  End If

  apos = 0
  bpos = 0

  If Is_Sub_Rejim(Protokol, CurPos) Then
    apos = InStr(CurPos, Protokol, "(")
    bpos = InStr(apos, Protokol, ")")
    If (apos > 0) And (bpos > apos) Then
      Get_Rejim_Name = "[" & Prefix & "] +Откл " & Trim(Mid(Protokol, apos + 1, bpos - apos - 1))
    End If
  Else
    apos = InStr(CurPos + 10, Protokol, vbCrLf)
    Get_Rejim_Name = "[" & Prefix & "] " & Find_BaseRejim_Name(Trim(Mid(Protokol, CurPos + 10, apos - (CurPos + 10))))
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

  ' Ищем окно ТКЗ-2000, если его нет выводим сообщение и завершаемся
  MainFormHandle = Find_TKZ_Window_Handle("TForm1")
  If MainFormHandle = 0 Then
    MsgBox "Окно ТКЗ-2000 не найдено, приложение должно быть запущено. Кроме этого должна быть загружена сеть для расчета.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' Инициализация, выбираем данные из листов
  Call Initialize

  ' Получаем номер узла (вообще-то узел может быть не только числовой)
  Dim Answer
  Answer = InputBox("Номер узла (рассчитываемые шины)?", "RootNode", 0)
  If Trim(Answer) = "" Then Exit Sub
  RootNode = Int(Answer)
  If Not Node_Exists(RootNode) Then
    MsgBox "Узел " & RootNode & " не найден в сети, дальнейшая работа невозможна.", vbExclamation + vbOKOnly
    Exit Sub
  End If

  ' Отобразим маленький диалог
  ProcessWnd = Show_Process_Window("Процесс идет, ждите...")

  ' Не проверяя в каком режиме работает программа (приказы или диалоговый расчет) выполним пункт меню
  Call SendMessage(MainFormHandle, WM_COMMAND, 12, 0&)     ' "Расширенный формат задания для расчета..."

  ' Найдем окно диалога задания приказов
  CommandFormHandle = Find_TKZ_Window_Handle("TFormZD2")

  ' Откроем окно протокола и очистим его
  Call SendMessage(CommandFormHandle, WM_COMMAND, 187, 0&) ' "Открыть протокол..."
  Call SendMessage(CommandFormHandle, WM_COMMAND, 200, 0&) ' "Очистить задание"
  ProtokolHandle = Find_TKZ_Window_Handle("TForm3")
  ProtokolMemo = Find_SubClass_Recurce(ProtokolHandle, "TMemo")
  Call SendMessage(ProtokolHandle, WM_COMMAND, 247, 0&)    ' "Очистить протокол"

  ' <<< Проверка чувствительности пусковых/избирательных органов

  ' Готовим приказ для проверки чувствительности и копируем его в окно приказов ТКЗ-2000
  CommandsText = Get_Sensitivity_Code()
  CommandRichEdit = Find_SubClass_Recurce(CommandFormHandle, "TRichEdit")
  Window_Set_Text CommandRichEdit, CommandsText

  ' Делаем расчет с эквивалентированием - это значительно быстрее
  Call SendMessage(CommandFormHandle, WM_COMMAND, 179, 0&) ' "Расчет с эквиваленированием"

  ' Заберем результат для анализа
  ProtokolText = Window_Get_Text(ProtokolMemo)

  ' Анализ протокола расчета
  Analiz_Sensitivity ProtokolText

  ' Предложить пользователю сохранить расширенный протокол (вначале добавлен исходный приказ с комментариями)
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText

  Dim filePRT
  Dim FrFi As Integer

  filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\Чувствительность " & RootNode & " узел.prt", "Файлы протокола АРМ (*.prt), *.prt")
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
  Call SendMessage(ProtokolHandle, WM_COMMAND, 247, 0&)    ' "Очистить протокол"
  Call SendMessage(CommandFormHandle, WM_COMMAND, 200, 0&) ' "Очистить задание"

  Window_Set_Text CommandRichEdit, CommandsText
  Call SendMessage(CommandFormHandle, WM_COMMAND, 179, 0&) ' "Расчет с эквиваленированием"
  
  ' Заберем результат для анализа
  ProtokolText = Window_Get_Text(ProtokolMemo)
  
  ' Анализируем результаты расчета
  Analiz_Testing (ProtokolText)

  ' Предложить пользователю сохранить расширенный протокол с комментариями
  ProtokolText = CommandsText & vbCrLf & "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~" & vbCrLf & ProtokolText
  filePRT = Application.GetSaveAsFilename(ActiveWorkbook.Path & "\Опробование " & RootNode & " узел.prt", "Файлы протокола АРМ (*.prt), *.prt")
  If filePRT <> "False" Then
    FrFi = FreeFile
    Open filePRT For Output As FrFi
    Print #FrFi, ProtokolText
    Close FrFi
  End If

  ' Уберем диалог процесса
  If ProcessWnd > 0 Then DestroyWindow (ProcessWnd)

End Sub
