VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Каталог_Цикл"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' https://www.youtube.com/channel/UCQMbRhaPEFD1NoZLhRzQzSA/videos?view=0&shelf_id=0&sort=dd
' https://inexsu.wordpress.com

' при вызове
' выдаёт следующее имя файла из подкаталога
' как в Ruby .each
' Чтобы не было Permission Denied:
' VBE-Tools-Options включи General - Break on Unhandled Errors

' в модуле
'  Каталог_Цикл As New Каталог_Цикл

Private FolderPath_local As String

Private Coll_FileNameFull As Collection

Private coll_Index As Long    ' см Class_Initialize

Sub Каталог_Цикл_Тест()
    ' procedure Checked by test опосредованно
    '
    '  Каталог_Цикл.Class_Initialize ' обновить счётчик

    Dim sMsg As String

    With Каталог_Цикл

        Do

            sMsg = .Файл_Путь_Полный("c:\TCPU\", "*.*", 4)

Debug.Print sMsg

        Loop While sMsg <> vbNullString

        .Class_Initialize

    End With
End Sub

Sub Class_Initialize()

    coll_Index = 1

End Sub

Function Файл_Путь_Полный(folderPath As String, _
        Optional mask As String = vbNullString, _
        Optional searchDeep As Long = 999, _
        Optional StatusBar As String = vbNullString) _
        As String
    ' procedure Checked by test опосредованно
    ' вернуть позицию в коллекции

    coll_FileNameFull_Init _
            folderPath, mask, searchDeep, StatusBar

    If в_Границах_Коллекции(Coll_FileNameFull, coll_Index) Then

        Файл_Путь_Полный = Coll_FileNameFull(coll_Index)

    End If

    coll_Index = coll_Index + 1

End Function

Private Function coll_FileNameFull_Init(folderPath As String, _
        Optional mask As String = vbNullString, _
        Optional searchDeep As Long = 999, _
        Optional StatusBar As String = vbNullString)
    ' procedure Checked by test опосредованно
    '
    If Coll_FileNameFull Is Nothing Then

        Set Coll_FileNameFull = FilenamesCollection( _
                folderPath, mask, searchDeep)

        FolderPath_local = folderPath

    End If

    If FolderPath_local <> folderPath Then

        Set Coll_FileNameFull = FilenamesCollection( _
                folderPath, mask, searchDeep, StatusBar)

        FolderPath_local = folderPath

    End If
End Function

Function FilenamesCollection(folderPath As String, _
        Optional mask As String = vbNullString, _
        Optional searchDeep As Long = 999, _
        Optional StatusBar As String = vbNullString) _
        As Collection
    ' для прохода по файлам каталога и подкаталогов, папок и подпапок
    ' Получает в качестве параметра путь к папке FolderPath,
    ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
    ' и глубину поиска SearchDeep в подпапках (если SearchDeep=1, то подпапки не просматриваются).
    ' Возвращает коллекцию, содержащую полные пути найденных файлов
    ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)

    ' можно использовать самостоятельно

    Set FilenamesCollection = New Collection    ' создаём пустую коллекцию

    Static fso As Object

    If fso Is Nothing Then

        Set fso = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject

    End If

    GetAllFileNamesUsingFSO _
            folderPath, mask, fso, FilenamesCollection, searchDeep    ' поиск

End Function

Private Function GetAllFileNamesUsingFSO( _
        folderPath As String, _
        mask As String, ByRef fso, _
        ByRef FileNamesColl As Collection, _
        searchDeep As Long, _
        Optional StatusBar As String = vbNullString)
    ' перебирает все файлы и подпапки в папке FolderPath, используя объект FSO
    ' перебор папок осуществляется в том случае, если SearchDeep > 1
    ' добавляет пути найденных файлов в коллекцию FileNamesColl

    Dim curfold As Object

    Set curfold = fso.GetFolder(folderPath)

    If Not curfold Is Nothing Then    ' если удалось получить доступ к папке

        ' раскомментируйте эту строку для вывода пути к просматриваемой
        ' в текущий момент папке в строку состояния Excel
        ' Application.StatusBar = "Поиск в папке: " & FolderPath

        Dim fil As Object

        For Each fil In curfold.Files        ' перебираем все файлы в папке FolderPath

            With fil

                If .Name Like "*" & mask Then

                    If Coll_IN_v_q(FileNamesColl, .Path) = False Then

                        FileNamesColl.Add .Path

                    End If
                End If
            End With
        Next

        ' при длительных файловых операциях можно использовать StatusBar

        searchDeep = searchDeep - 1    ' уменьшаем глубину поиска в подпапках

        Dim sfol As Object

        If searchDeep Then    ' если надо искать глубже

            For Each sfol In curfold.SubFolders    ' перебираем все подпапки в папке FolderPath

                GetAllFileNamesUsingFSO _
                        sfol.Path, mask, fso, FileNamesColl, searchDeep

            Next
        End If
    End If
End Function

Function Coll_IN_v_q(Coll As Collection, msg As String) _
        As Boolean
    ' procedure Checked by test опосредованно
    ' есть ли элемент в коллекции ?
    ' АртеФакт

    Dim var As Variant

    For Each var In Coll

        If var = msg Then

            Coll_IN_v_q = True

            Exit For

        End If
    Next
End Function

Function в_Границах_Коллекции(Coll As Collection, _
        indx As Long) _
        As Boolean
    ' procedure Checked by test опосредованно
    '
    If indx > 0 And indx <= Coll.Count Then

        в_Границах_Коллекции = True

    End If
End Function
