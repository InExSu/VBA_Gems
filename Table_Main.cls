VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' таблицу создать
' для модуля я_Size (BookSize)

Private a2() As Variant

Function Create( _
         Optional s As String)
' Тестом НЕ покрыто
'
    ReDim a2(1 To A2_Rows + 1, 1 To UBound(a1_Headers) + 1)

End Function

Function A2_Headers_Fill(a2() As Variant)
' Тестом НЕ покрыто
'  декоратор для массив одномерный в строку массива двумерного

    Dim lCol As Long

    '    For lCol = LBound(a1) To UBound(a1)
    '       ' a2(1,lCol)
    '    Next
End Function

Function A1_2_a1_Row( _
         a1 As Variant, _
         a2() As Variant, _
         lRow As Long, _
         lCol As Long) _
         As String
' Тестом НЕ покрыто
' массив одномерный в строку массива двумерного
End Function

Function A2_Rows() _
         As Long
' procedure Checked by test
'
    A2_Rows = App_WorkSheets_Count * _
              UBound(a1_ParaMeter) + 1

End Function

Function App_WorkSheets_Count() _
         As Long
' procedure Checked by test
' количество листов во всех открытых книгах
    Dim iCount As Long

    Dim wb As Workbook
    For Each wb In Application.Workbooks

        iCount = iCount + wb.Worksheets.Count

    Next
    App_WorkSheets_Count = iCount
End Function

Function a1_Headers() _
         As Variant
' procedure Checked by test
' здесь определяется количество столбцов таблицы

    a1_Headers = Array( _
                 "Книга", "Лист", "ПараМетр", "Колво на Листе", "Колво в Книге")

End Function

Function a1_ParaMeter() _
         As Variant
' procedure Checked by test
'
    a1_ParaMeter = Array( _
                   "Лист_Размер", "Лист_Скрытый", "Строки висячие, ячеек", "Столбцов правых, ячеек", "Стили НЕСтандарт", "Диап Имя Ошибки", "Объектов", "Примечаний")

End Function
