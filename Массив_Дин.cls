VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Массив_Дин"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' динамический d2 двумерный массив изменяемый
' и по строкам и по столбцам

' в модуле нужно разместить:
' Массив_Дин As New Массив_Дин
'Type Таблица_Строки
'    Столбец() As Variant
'End Type

' Использование:
' Массив_Дин.Создать Строк, Столбцов  // создать пустой заданного размера
' Массив_Дин.в_Ячейку Строку, Столбец, Значение // поместить в ячейку значение и если строка или столбец больше,  чем имеется динамически расширить
' d2 = Массив_Дин.Статик // вернуть обычный массив
' Массив_Дин.Столбцы_Колво_Изменить
' Массив_Дин.Строки_Колво_Изменить
' str = Массив_Дин.Ячейка(Строка, Столбец) //  вернуть содержимое ячейки

' массив одномерный массивов одномерных
Private Таблица_Строки() As Таблица_Строки    ' основа массива динамического


Function Создать( _
        Строк As Long, _
        Столбцов As Long)
    ' procedure Checked by test опосредованно

    ' индекс с 1 для массивов, работающих с листами

    ReDim Таблица_Строки(1 To Строк)

    Dim y As Long

    For y = LBound(Таблица_Строки) To UBound(Таблица_Строки)

        ReDim Таблица_Строки(y).Столбец(1 To Столбцов)

    Next
End Function

Function в_Ячейку(Строка As Long, _
        Столбец As Long, _
        Значение As Variant)                    ' Тестом покрыта
    ' динамически увеличивает "массив"
    ' при необходимости

    If Строка > UBound(Таблица_Строки) Then _
            Строки_Колво_Изменить Строка

    If Столбец > UBound(Таблица_Строки(1).Столбец) Then _
            Столбцы_Колво_Изменить Столбец

    Таблица_Строки(Строка).Столбец(Столбец) = Значение

End Function

Function Ячейка(Строка As Long, _
        Столбец As Long) _
        As Variant                   ' procedure Checked by test опосредованно
    ' вернуть содержимое ячейки "таблицы массива"
    ' get/геттер

    Ячейка = Таблица_Строки(Строка).Столбец(Столбец)

End Function

Function Строки_Колво_Изменить(Строк As Long)

    ReDim Preserve Таблица_Строки(1 To Строк)

    Dim y As Long

    For y = LBound(Таблица_Строки) To Строк

        ReDim Preserve Таблица_Строки(y).Столбец(1 To Столбцов)

    Next
End Function

Function Столбцы_Колво_Изменить(Столбец As Long)

    Dim y As Long

    For y = LBound(Таблица_Строки) To UBound(Таблица_Строки)

        ReDim Preserve Таблица_Строки(y).Столбец(1 To Столбец)

    Next
End Function

Function Статик(Optional msg As String) _
        As Variant()   ' Тестом  покрыта
    ' вернуть обычный d2 массив
    ' если неохота привыкать к использованию:
    ' Таблица_Строки(y).Столбец(x)

    Dim d2() As Variant
    ReDim d2(1 To UBound(Таблица_Строки), _
            1 To UBound(Таблица_Строки(1).Столбец))

    Dim y As Long, x As Long

    For y = LBound(d2) To UBound(d2)
        For x = LBound(d2, 2) To UBound(d2, 2)
            d2(y, x) = Таблица_Строки(y).Столбец(x)
        Next
    Next

    If IsArray(d2) Then _
            Статик = d2

End Function

Function Заполнить(Optional msg As String)
    '  тестом Покрыто опосредованно
    ' для отладки
    Dim y As Long, x As Long

    For y = LBound(Таблица_Строки) To UBound(Таблица_Строки)
        For x = LBound(Таблица_Строки(y).Столбец) To UBound(Таблица_Строки(y).Столбец)
            Таблица_Строки(y).Столбец(x) = y & x
        Next
    Next
End Function

Function Усечь(Строк As Long, Столбцов As Long) _
        As Variant   ' procedure Checked by test опосредованно

    Строки_Колво_Изменить Строк
    Столбцы_Колво_Изменить Столбцов

End Function

Function Строк(Optional msg As String) _
        As Long   ' procedure Checked by test опосредованно
    ' вернуть колво строк

    Строк = UBound(Таблица_Строки)
End Function

Function Столбцов(Optional msg As String) _
        As Long   ' procedure Checked by test опосредованно
    ' вернуть колво ст

    Столбцов = UBound(Таблица_Строки( _
            LBound(Таблица_Строки)).Столбец)
End Function

Function G_Массив_Дин(lRows As Long, _
        lCols As Long) _
        As Массив_Дин
    ' procedure Checked by test опосредованно
    '
    Dim a2_dyn As New Массив_Дин

    a2_dyn.Создать lRows, lRows

    Массив_Дин_Заполнить a2_dyn

    Set G_Массив_Дин = a2_dyn

End Function

Function G_Массив_Дин_Пустой(lRows As Long, _
        lCols As Long) _
        As Массив_Дин
    ' Тестом  покрыто
    '
    Dim a2_Dyn_Empty As New Массив_Дин

    a2_Dyn_Empty.Создать lRows, lRows

    Set G_Массив_Дин_Пустой = a2_Dyn_Empty

End Function

Function Массив_Дин_Заполнить(a2_Дин As Массив_Дин)
    ' procedure Checked by test опосредованно

    With a2_Дин

        Dim y As Long, x As Long

        For y = 1 To .Строк

            For x = 1 To .Столбцов

                .в_Ячейку y, x, Mock.G_S

            Next
        Next
    End With
End Function


