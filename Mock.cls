VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' https://www.youtube.com/channel/UCQMbRhaPEFa1NoZLhRzQzSA/videos?view=0&shelf_id=0&sort=dd
' https://inexsu.wordpress.com

' мок
' предоставляет временные переменные:
' книгу, лист, диапазон, массив, строку, символ, числов

Private Wb_Act_Name As String
Private Ws_Act_Name As String

Private Wb_p As Workbook
Private Wb_p_Name As String

Private Ws_p As Worksheet
Private Ws_p_Name As String

Public bDebug As Boolean

'
Private Sub Class_Initialize()    ' procedure Checked by test опосредованно

    With Application

        Dim AppScrUp As Boolean
        AppScrUp = .ScreenUpdating

        If AppScrUp Then .ScreenUpdating = False

        If Окна_Видимые_Есть = False Then Set Wb_p = Workbooks.Add
        ' чтобы вернуться
        Wb_Act_Name = ActiveWorkbook.Name
        Ws_Act_Name = ActiveSheet.Name

        Set Wb_p = Workbooks.Add

        With Wb_p

            Wb_p_Name = .Name

            Set Ws_p = .Worksheets(1)

        End With

        Ws_p_Name = Ws_p.Name

        If ОткрытаЛиКнига(Wb_Act_Name) Then

            Workbooks(Wb_Act_Name).Activate

            ActiveWorkbook.Worksheets(Ws_Act_Name).Select

        End If

        If .ScreenUpdating <> AppScrUp Then _
           .ScreenUpdating = AppScrUp
    End With
End Sub

Function G_Workbook(Optional ByVal msg As String) _
         As Workbook
    ' procedure Checked by test
    '
    Class_Initialize

    Set G_Workbook = Wb_p

End Function

Private Sub Class_Terminate()        ' procedure Checked by test опосредованно

    If ОткрытаЛиКнига(Wb_p_Name) Then _
       Wb_p.Close False

End Sub

Function wb(Optional ByVal msg As String) _
         As Workbook
    ' Тестом покрыта

    Если_Книгу_Мок_Закрыли_Открыть

    Set wb = Wb_p

End Function

Function ws(Optional ByVal msg As String) _
         As Worksheet
    ' Тестом покрыта

    Class_Initialize

    Set ws = Ws_p

End Function

Function Если_Книгу_Мок_Закрыли_Открыть(Optional ByVal msg As String)
    ' Тестом покрыта
    ' если книгу случайно закрыли

    If ОткрытаЛиКнига(Wb_p_Name) = False Then

        Class_Initialize

    End If
End Function

Function ОткрытаЛиКнига(ByVal wbName As String) As Boolean
    ' открыта ли книга
    ' вызов If Not ОткрытаЛиКнига(ИмяФайла) Then Workbooks.Open ИмяФайла

    wbName = Right$(wbName, Len(wbName) - InStrRev(wbName, "\"))    ' отсечь путь

    Static wbBook As Workbook

    For Each wbBook In Workbooks

        If wbBook.Name = wbName Then

            ОткрытаЛиКнига = True

            Exit For
        End If
    Next
End Function

' '==== Генераторы
Function G_a2(Optional ByVal lRow As Long, _
              Optional ByVal lCol As Long, _
              Optional ByVal row_Len As Long) _
              As Variant()
    ' Тестом покрыта
    ' генератор массива, скромно, быстро, без вопросов
    ' индекс начинаю с 1, для вставки в лист

    ' variables initialize

    If lRow = 0 Then lRow = Rand_Between(1, 9)

    If lCol = 0 Then lCol = Rand_Between(1, 9)

    If row_Len = 0 Then row_Len = Rand_Between(1, 9)

    ' declare array

    Static a2() As Variant

    ReDim a2(1 To lRow, 1 To lCol)

    ' array fill

    Dim y As Long, x As Long

    For y = LBound(a2) To UBound(a2)

        For x = LBound(a2, 2) To UBound(a2, 2)

            a2(y, x) = G_S

        Next
    Next

    G_a2 = a2

End Function

Function G_a2_0(Optional ByVal lRow As Long, _
                Optional ByVal lCol As Long, _
                Optional ByVal row_Len As Long) _
                As Variant()
    ' procedure Checked by test опосредованно
    ' генератор массива, скромно, быстро, без вопросов

    ' variables initialize

    If lRow = 0 Then lRow = Rand_Between(1, 9)

    If lCol = 0 Then lCol = Rand_Between(1, 9)

    If row_Len = 0 Then row_Len = Rand_Between(1, 9)

    ' declare array

    Static a2() As Variant

    ReDim a2(0 To lRow, 0 To lCol)

    ' array fill

    Dim y As Long, x As Long

    For y = LBound(a2) To UBound(a2)

        For x = LBound(a2, 2) To UBound(a2, 2)

            a2(y, x) = G_S

        Next
    Next

    G_a2_0 = a2

End Function

Function G_a1(Optional ByVal row_Min As Long, _
              Optional ByVal row_Max As Long, _
              Optional ByVal row_Len As Long) _
              As Variant()
    ' Тестом покрыта
    ' генератор массива одномерного, скромно, быстро, без вопросов

    ' variables initialize

    If row_Min = 0 Then row_Min = Rand_Between(0, 9)

    If row_Max = 0 Then row_Max = Rand_Between(row_Min, 9)

    If row_Len = 0 Then row_Len = Rand_Between(1, 9)

    If row_Min = -1 Then row_Min = 0

    If row_Max = -1 Then row_Max = 0

    ' declare array

    Dim a1 As Variant
    ReDim a1(row_Min To row_Max)

    Dim y As Long
    For y = LBound(a1) To UBound(a1)

        a1(y) = G_S(row_Len)

    Next

    G_a1 = a1

End Function

Function G_a1_Latin(Optional ByVal row_Min As Long, _
                    Optional ByVal row_Max As Long, _
                    Optional ByVal row_Len As Long) _
                    As Variant()
    ' Тестом покрыта
    ' генератор массива, скромно, быстро, без вопросов

    ' variables initialize

    If row_Min = 0 Then row_Min = Rand_Between(0, 9)

    If row_Max = 0 Then row_Max = Rand_Between(row_Min, 9)

    If row_Len = 0 Then row_Len = Rand_Between(1, 9)

    ' declare array

    Static a1 As Variant

    ReDim a1(row_Min To row_Max)

    ' array fill

    Dim y As Long

    For y = LBound(a1) To UBound(a1)

        a1(y) = G_S_Latin(row_Len)

    Next

    G_a1_Latin = a1

End Function

Function G_a2_Latin(Optional ByVal lRow As Long, _
                    Optional ByVal lCol As Long, _
                    Optional ByVal row_Len As Long) _
                    As Variant()
    ' Тестом покрыта
    ' генератор массива, скромно, быстро, без вопросов

    ' variables initialize

    If lRow = 0 Then lRow = Rand_Between(1, 9)

    If lCol = 0 Then lCol = Rand_Between(1, 9)

    If row_Len = 0 Then row_Len = Rand_Between(1, 9)

    ' declare array

    Static a2() As Variant
    ReDim a2(1 To lRow, 1 To lCol)

    ' array fill

    Dim y As Long, x As Long

    For y = LBound(a2) To UBound(a2)

        For x = LBound(a2, 2) To UBound(a2, 2)

            a2(y, x) = G_S_Latin

        Next
    Next

    G_a2_Latin = a2

End Function

Function G_a1_Double(Optional ByVal row_Min As Long, _
                     Optional ByVal row_Max As Long, _
                     Optional ByVal row_Len As Long) _
                     As Variant()
    ' procedure Checked by test опосредованно
    ' генератор массива, скромно, быстро, без вопросов

    ' variables initialize

    If row_Min = 0 Then row_Min = Rand_Between(0, 9)

    If row_Max = 0 Then row_Max = Rand_Between(row_Min, 9)

    If row_Len = 0 Then row_Len = Rand_Between(1, 9)

    ' declare array

    Static a1 As Variant

    ReDim a1(row_Min To row_Max)

    ' array fill

    Dim y As Long

    For y = LBound(a1) To UBound(a1)

        a1(y) = G_Double

    Next

    G_a1_Double = a1

End Function

Function G_Date(Optional ByVal msg As String) _
         As Date
    ' procedure Checked by test опосредованно

    G_Date = Date + Rand_Between(-64, 64)

End Function

Function G_Date_beTween( _
         dMin As Date, _
         dMax As Date) _
         As Date
    ' procedure NOT checked by test
    '

    G_Date_beTween = CDate( _
                     Rand_Between( _
                     CLng(dMin), CLng(dMax)))

End Function

Function G_String(Optional ByVal iLen As Long) _
         As String
    ' procedure NOT checked by test
    '
    G_String = G_S(iLen)

End Function

Function G_S(Optional ByVal iLen As Long) _
         As String
    ' Тестом покрыта
    ' генератор случайных строк

    If iLen = 0 Then iLen = Rand_Between(1, 9)

    Dim str As String

    Dim y As Long

    For y = 1 To iLen

        str = str & Symb_Rand

    Next

    G_S = str

End Function

Function G_Shape(Optional ByVal msg As String) _
         As Shape
    ' procedure Checked by test опосредованно

    Static sh As Shape

    Set sh = ws.Shapes.AddShape(msoShapeRectangle, _
                                144, 144, 72, 72)

    With sh
        .Name = "Red Square"
        .Fill.ForeColor.RGB = RGB(255, 0, 0)
        .Line.DashStyle = msoLineDashDot
    End With

    Set G_Shape = sh
End Function

Function G_Variant(Optional ByVal iLen As Long) _
         As Variant
    ' Тестом покрыта
    ' генератор случайных строк

    If iLen = 0 Then iLen = Rand_Between(1, 9)

    Dim str As Variant

    Dim y As Long

    For y = 1 To iLen

        str = str & Symb_Rand

    Next

    G_Variant = str

End Function
Function Symb_Rand(Optional ByVal msg As String) _
         As String
    ' procedure Checked by test

    Symb_Rand = Chr$(Rand_Between(0, 255))

End Function

Function Rand_Between(ByVal iMin As Long, _
                      ByVal iMax As Long) As Double
    ' procedure Checked by test опосредованно

    ' Для сокращения кода, чтобы не читать _
      Application.WorksheetFunction.RandBetween(iMin, _
      iMax)

    ' Возвращает случайное число целое,
    ' в диапазоне между двумя заданными числами.

    ' Randomize

    With Application.WorksheetFunction

        Rand_Between = . _
                       RandBetween(iMin, iMax)

    End With
End Function

Function Option_Base(Optional ByVal msg As String) As Long
    ' procedure Checked by test

    Static arr_1d As Variant

    arr_1d = Array(vbNullString)

    Option_Base = LBound(arr_1d)

End Function

Function isAx(d As Variant) _
         As Boolean  ' Тестом покрыта
    ' раньше называл isDx
    ' isarray массив ли, проверить массив по взрослому, массивли

    Err.Clear

    On Error Resume Next

    If LBound(d) > -1 Then

        If Err.Number = 0 Then

            isAx = True

        End If
    End If
End Function

Function G_Worksheet(Optional ws As Worksheet) _
         As Worksheet
    ' Тестом Покрыта
    ' накидать в ячейки значения

    Class_Initialize

    Set G_Worksheet = Ws_p

End Function

Function G_Range(Optional ByVal msg As String) _
         As Range
    ' Тестом Покрыта

    Static row_Up As Long, row_dowN As Long, _
           col_Left As Long, col_Right As Long

    row_Up = Rand_Between(1, 30 / 2)
    row_dowN = Rand_Between(row_Up, 30)

    col_Left = Rand_Between(1, 10 / 2)
    col_Right = Rand_Between(col_Left, 10)

    Dim r As Range

    With ws

        Set r = .Range(.Cells(row_Up, col_Left), _
                       .Cells(row_dowN, col_Right))

    End With

    Dim a2() As Variant

    With r

        a2 = .Value

        a2 = G_a2( _
             UBound(a2), UBound(a2, 2))

        .Value = a2

    End With

    Set G_Range = r

End Function

Function a2_Replace_Left_X( _
         a2() As Variant, _
         lCol As Long, _
         sFind As String, _
         sRepl As String)
    ' Тестом НЕ покрыто
    ' in the array replace at the beginning of the line
    ' в массиве заменить в начале строки

    Dim lRow As Long

    For lRow = LBound(a2) To UBound(a2)

        Replace _
                a2(lRow, lCol), _
                sFind, sRepl, _
                1, , vbBinaryCompare

    Next
End Function

Function G_Double(Optional ByVal msg As String) _
         As Double
    ' procedure Checked by test опосредованно

    G_Double = Rand_Between(0, 65432) + _
               Rand_Between(0, 99) / 100

End Function


Function G_Long(Optional ByVal msg As String) _
         As Long
    ' Тестом Покрыта

    G_Long = WorksheetFunction.RandBetween(1, 9)

End Function

Function G_Digit(Optional ByVal msg As String) _
         As Long
    ' Тестом Покрыта

    G_Digit = WorksheetFunction.RandBetween(0, 9)

End Function

Public Sub a2msg(a2() As Variant)
    ' procedure Checked by test опосредованно

    ' показать массив в MsgBox

    If isAx(a2) = False Then Exit Sub

    Dim str As String
    Dim y As Long, x As Long, a2_Cell_Len As Long

    a2_Cell_Len = Массив_Ячейка_Длина_Макс(a2)

    Static row_Max As Long, col_Max As Long
    row_Max = 30: col_Max = 40

    If UBound(a2) < row_Max Then row_Max = UBound(a2)
    If UBound(a2) < col_Max Then col_Max = UBound(a2, 2)

    For y = LBound(a2) To UBound(a2)

        For x = LBound(a2, 2) To UBound(a2, 2)

            str = str & a2(y, x) & " / "

        Next

        str = Left(str, Len(str) - 3)

        str = str & vbCrLf

        If y = 50 Then Exit For

    Next

    If bDebug = False Then

        MsgBox str, vbOKOnly, "a2msg. Возмжно уместилось НЕ всё ... :-("

    End If
End Sub

Function Массив_Ячейка_Длина_Макс(a2() As Variant) _
         As Long  ' Тестом покрыта

    ' максимальная длина ячейки в массиве

    If isAx(a2) = False Then

        Exit Function

    End If

    Dim y As Long, x As Long, len_Cell As Long

    For y = LBound(a2) To UBound(a2)

        For x = LBound(a2, 2) To UBound(a2, 2)

            If Len(a2(y, x)) > len_Cell Then

                len_Cell = Len(a2(y, x))

            End If
        Next
    Next

    Массив_Ячейка_Длина_Макс = len_Cell

End Function


Function isWs(ByVal ws As Worksheet) _
         As Boolean
    ' procedure Checked by test опосредованно

    ' если книга объявлена, но не проинициализирована

    On Error GoTo isWs_Error

    If Not ws.Cells(1, 1) Is Nothing Then

        isWs = True

    End If

    Exit Function

isWs_Error:

    isWs = False

End Function

Function Workbooks_Worksheet_Name(ByVal ws_Name As String) _
         As Worksheet
    ' procedure Checked by test опосредованно
    ' найти и вернуть лист среди всех открытых книг

    If ws_Name = vbNullString Then Exit Function    '=>

    Dim wb As Workbook, ws As Worksheet

    For Each wb In Application.Workbooks

        For Each ws In wb.Worksheets

            If ws.Name = ws_Name Then

                Set Workbooks_Worksheet_Name = ws

                Exit Function

            End If
        Next
    Next
End Function

Function ws_Exist(ByVal wb As Workbook, ByVal ws_Name As String) _
         As Boolean
    ' procedure Checked by test опосредованно
    ' существует ли лист в указанной книге

    If ws_Name = vbNullString Then Exit Function    '=>

    Static ws As Worksheet

    For Each ws In wb.Worksheets

        If ws.Name = ws_Name Then

            ws_Exist = True

            Exit Function

        End If
    Next
End Function

Function Лист_Слово_Случайно(ByVal str As String, _
                             Optional ByVal howMany As Long, _
                             Optional ws_ As Worksheet) _
                             As Worksheet
    ' procedure Checked by test
    ' вернуть лист со случайно расставленным str

    If ws_ Is Nothing Then Set ws_ = ws

    If howMany = 0 Then howMany = Rand_Between(1, 9)

    Static x As Long

    With ws_

        For x = 1 To howMany

            .Cells(Rand_Between(1, 30), _
                   Rand_Between(1, 11)).Value = str

        Next
    End With

    Set Лист_Слово_Случайно = ws_

End Function

Function G_Boolean(Optional ByVal msg As String) _
         As Boolean
    ' procedure Checked by test опосредованно

    G_Boolean = Rand_Between(0, 1)

End Function

Function Окна_Видимые_Есть(Optional ByVal msg As String) _
         As Boolean

    ' procedure Checked by test

    ' АртеФакт

    Dim win As Window

    For Each win In Application.Windows

        If win.Visible Then

            Окна_Видимые_Есть = True

            Exit Function

        End If
    Next
End Function

Function G_ChartObject(Optional ByVal msg As String) _
         As ChartObject

    ' procedure Checked by test опосредованно
    Static ws As Worksheet
    Set ws = wb.ActiveSheet

    DoEvents

    Set G_ChartObject = ws.ChartObjects.Add

End Function

Function G_Object(Optional ByVal msg As String) _
         As Object

    ' procedure Checked by test опосредованно
    Static obj As Object
    DoEvents

    Set G_Object = obj

End Function

Function G_S_Latin( _
         Optional iLen As Long) _
         As String

    ' procedure Checked by test опосредованно
    ' генератор строк латинских случайных

    If iLen = 0 Then iLen = Rand_Between(1, 9)

    Dim str As String
    str = vbNullString

    Dim y As Long

    With Mock

        For y = 1 To iLen

            str = str & Symb_Latin

        Next

    End With

    G_S_Latin = str

End Function

Function Symb_Latin(Optional msg As String) _
         As String
    ' procedure Checked by test

    If Rand_Between(0, 1) = 0 Then

        Symb_Latin = Chr$(Rand_Between(65, 90))

    Else

        Symb_Latin = Chr$(Rand_Between(97, 122))

    End If
End Function


Function Диалог_Каталог(Optional ByVal Путь As String) _
         As String
    ' Тестом НЕ покрыто
    Dim oFD As FileDialog
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)

    With oFD        'используем короткое обращение к объекту

        .Title = "Каталог"
        .ButtonName = "Выбрать"
        .Filters.Clear                'очищаем установленные ранее типы файлов
        .InitialFileName = Путь        '"назначаем первую папку отображения
        .InitialView = msoFileDialogViewLargeIcons        'вид диалогового окна(доступно 9 вариантов)

        If .Show = 0 Then _
           Exit Function             'показывает диалог

        Диалог_Каталог = .SelectedItems(1)        'считываем путь к папке

    End With
End Function

Function A2_Empty( _
         a2() As Variant) _
         As Variant()
    ' procedure Checked by test
    '
    Dim a2_New() As Variant

    ReDim a2_New(LBound(a2) To UBound(a2), _
                 LBound(a2, 2) To UBound(a2, 2))

    A2_Empty = a2_New

End Function

Function G_Password_Type( _
         sType As String, _
         iLen As Long) _
         As String
    ' procedure Checked by test
    ' пароль строка символов из диапазона ASCII
    Dim iMin As Long, iMax As Long

    Select Case sType
    Case "00-"
        iMin = 0
        iMax = Asc("0") - 1
    Case "09"
        iMin = Asc("0")
        iMax = Asc("9")
    Case "9+A-"
        iMin = Asc("9") + 1
        iMax = Asc("A") - 1
    Case "AZ"
        iMin = Asc("A")
        iMax = Asc("Z")
    Case "Z+a-"
        iMin = Asc("Z") + 1
        iMax = Asc("a") - 1
    Case "az"
        iMin = Asc("a")
        iMax = Asc("z")
    Case "z+А-"
        iMin = Asc("z") + 1
        iMax = Asc("А") - 1
    Case "АЯ"
        iMin = Asc("А")
        iMax = Asc("Я")
    Case "ая"
        iMin = Asc("а")
        iMax = Asc("я")

    End Select

    G_Password_Type = G_Password_Range( _
                      iMin, iMax, iLen)

End Function

Function G_Password_Range( _
         iMin As Long, _
         iMax As Long, _
         iLen As Long) _
         As String
    ' procedure Checked by test
    ' генератор пароля строки из диапазона ASCII

    Dim s As String

    Dim x As Long

    For x = 1 To iLen

        s = s & _
            G_Symb_Rand_ASCII_Range( _
            iMin, iMax)

    Next

    G_Password_Range = s

End Function

Function G_Symb_Rand_ASCII_Range( _
         iMin As Long, _
         iMax As Long) _
         As String
    ' procedure Checked by test
    ' символ случайный из диапазона ASCII

    G_Symb_Rand_ASCII_Range = Chr( _
                              Rand_Between( _
                              iMin, iMax))

End Function
