VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private End_Year As Long
Private End_Month As Long
Private End_Day As Long

Public Sub Workbook_Open()

    ' Date_End = DateSerial(2019, 7, 31) ' дата последнего дня работы надстройки

    Пробил_Час

    If НЕ_InExSu Then Exit Sub    '=>

    Application.OnKey "^+{R}", "RefStyle_Сhange"

    Копия_Дата_Время ThisWorkbook

End Sub

Sub Xlam_UnView_Make_()

    ' И обратно дороги нет
    ' в Workbook_Open сделай Date_End

    '    Копия_Дата_Время ThisWorkbook, vbNullString
    '
    '    Make_Xlam_UnView = True
    '
    '    ThisWorkbook.Save

End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    If Make_Xlam_UnView = False Then Exit Sub

    Application.DisplayAlerts = False

    With ThisWorkbook
        If Not .MultiUserEditing Then
            .SaveAs filename:=.FullName, accessMode:=xlShared
        End If
    End With

    Application.DisplayAlerts = True

End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)

    If Make_Xlam_UnView = False Then Exit Sub

    ThisWorkbook.IsAddin = True

End Sub

Function Пробил_Час( _
         Optional tDate_End As Date) _
         As String
    ' procedure Checked by test опосредованно
    '
    If tDate_End = Empty Then

        tDate_End = Date_End

        If tDate_End = Empty Then

            Exit Function
        
        End If
    End If

    If Now > tDate_End Then

        uf_About.Show

        End    ' ->

    Else

        Пробил_Час = "Z"

    End If
End Function

Public Sub Копия_Дата_Время(wb As Workbook, _
                            Optional подКаталог As String = "Версии", _
                            Optional Force As Boolean = False)

    If Окна_Видимые_Есть = False Then

        'Workbooks.Add

    End If
    ' Создает резервную копию файла

    wb.SaveCopyAs _
            Имя_Дата_Время(wb.Name, _
                           Каталог( _
                           Путь(wb.Path, подКаталог)))
End Sub


Function НЕ_InExSu(Optional msg As String)
    ' procedure Checked by test опосредованно

    With Application

        If .UserName = "InExSu" Or .UserName = "\\Mihail" Then

            НЕ_InExSu = False

        End If
    End With
End Function

Function Имя_Дата_Время(wbName As String, _
                        Путь As String) _
                        As String
    Имя_Дата_Время = Путь & _
                     Имя_без_Расширения(wbName) & " " & _
                     Format$(Date, "yyyy-mm-dd") & " " & Format$(Time, "hh-mm-ss") & _
                     "." & _
                     Расширение_без_Имени(wbName)
End Function

Private Function Каталог(Путь As String) As String

    If Каталог_Есть_FSO(Путь) = False Then

        MkDir Путь

    End If

    Каталог = Путь

End Function

Function Каталог_Есть_FSO(msg As String) _
         As Boolean
    ' procedure Checked by test опосредованно

    With CreateObject("Scripting.FileSystemObject")

        If .FolderExists(msg) Then

            Каталог_Есть_FSO = True

        End If
    End With
End Function

Private Function Путь(sPath As String, _
                      подКаталог As String) _
                      As String

    If подКаталог = vbNullString Then _
       подКаталог = "Версии"

    sPath = Если_Путь_Запрошенный_Пуст_Выход(sPath)

    With Application

        Путь = sPath & _
               .PathSeparator & _
               подКаталог & _
               .PathSeparator

    End With
End Function

Function Если_Путь_Запрошенный_Пуст_Выход( _
         Optional sPath As String) _
         As String
    ' Тестом НЕ покрыто
    ' АртеФакт

    If sPath <> vbNullString Then

        Если_Путь_Запрошенный_Пуст_Выход = sPath

        Exit Function

    End If

    sPath = Диалог_Каталог

    If sPath = vbNullString Then

        End

    End If

    Если_Путь_Запрошенный_Пуст_Выход = Сепаратор_Правый_Отсечь(sPath)

End Function

Function Сепаратор_Правый_Отсечь(sPath As String) _
         As String
    ' procedure Checked by test
    ' если путь заканчивается на \, отсечь его
    ' АртеФакт

    If Right$(sPath, 1) = Application.PathSeparator Then

        sPath = Left$(sPath, Len(sPath) - 1)

    End If

    Сепаратор_Правый_Отсечь = sPath

End Function

Function Расширение_без_Имени(str As String) _
         As String
    'Возвращает имя файла без расширения
    ' procedure Checked by test

    str = CreateObject("Scripting.FileSystemObject"). _
          GetExtensionName(str)

    Dim iPos As Long
    iPos = Len(str) - InStrRev(str, ".")

    Расширение_без_Имени = Right$(str, iPos)

End Function

Function Имя_без_Расширения(str As String) _
         As String
    ' Возвращает расширение без имени файла
    ' procedure Checked by test

    If InStr(str, ".") < 1 Then
        Имя_без_Расширения = str
        Exit Function
    End If

    str = CreateObject("Scripting.FileSystemObject"). _
          GetFileName(str)      ' отсекает путь к файлу

    Dim iPos As Long
    iPos = InStr(str, ".") - 1

    Имя_без_Расширения = Left$(str, iPos)

End Function

Function Окна_Видимые_Есть(Optional msg As String) _
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

Function Диалог_Каталог(Optional Путь As String) _
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
           Exit Function         'показывает диалог

        Диалог_Каталог = .SelectedItems(1)        'считываем путь к папке

    End With
End Function

Sub del2()
    ' procedure NOT checked by test
    '

End Sub



