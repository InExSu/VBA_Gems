VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "�������_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' https://www.youtube.com/channel/UCQMbRhaPEFD1NoZLhRzQzSA/videos?view=0&shelf_id=0&sort=dd
' https://inexsu.wordpress.com

' ��� ������
' ����� ��������� ��� ����� �� �����������
' ��� � Ruby .each
' ����� �� ���� Permission Denied:
' VBE-Tools-Options ������ General - Break on Unhandled Errors

' � ������
'  �������_���� As New �������_����

Private FolderPath_local As String

Private Coll_FileNameFull As Collection

Private coll_Index As Long    ' �� Class_Initialize

Sub �������_����_����()
    ' procedure Checked by test �������������
    '
    '  �������_����.Class_Initialize ' �������� �������

    Dim sMsg As String

    With �������_����

        Do

            sMsg = .����_����_������("c:\TCPU\", "*.*", 4)

Debug.Print sMsg

        Loop While sMsg <> vbNullString

        .Class_Initialize

    End With
End Sub

Sub Class_Initialize()

    coll_Index = 1

End Sub

Function ����_����_������(folderPath As String, _
        Optional mask As String = vbNullString, _
        Optional searchDeep As Long = 999, _
        Optional StatusBar As String = vbNullString) _
        As String
    ' procedure Checked by test �������������
    ' ������� ������� � ���������

    coll_FileNameFull_Init _
            folderPath, mask, searchDeep, StatusBar

    If �_��������_���������(Coll_FileNameFull, coll_Index) Then

        ����_����_������ = Coll_FileNameFull(coll_Index)

    End If

    coll_Index = coll_Index + 1

End Function

Private Function coll_FileNameFull_Init(folderPath As String, _
        Optional mask As String = vbNullString, _
        Optional searchDeep As Long = 999, _
        Optional StatusBar As String = vbNullString)
    ' procedure Checked by test �������������
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
    ' ��� ������� �� ������ �������� � ������������, ����� � ��������
    ' �������� � �������� ��������� ���� � ����� FolderPath,
    ' ����� ����� ������� ������ Mask (����� �������� ������ ����� � ����� ������/�����������)
    ' � ������� ������ SearchDeep � ��������� (���� SearchDeep=1, �� �������� �� ���������������).
    ' ���������� ���������, ���������� ������ ���� ��������� ������
    ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)

    ' ����� ������������ ��������������

    Set FilenamesCollection = New Collection    ' ������ ������ ���������

    Static fso As Object

    If fso Is Nothing Then

        Set fso = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject

    End If

    GetAllFileNamesUsingFSO _
            folderPath, mask, fso, FilenamesCollection, searchDeep    ' �����

End Function

Private Function GetAllFileNamesUsingFSO( _
        folderPath As String, _
        mask As String, ByRef fso, _
        ByRef FileNamesColl As Collection, _
        searchDeep As Long, _
        Optional StatusBar As String = vbNullString)
    ' ���������� ��� ����� � �������� � ����� FolderPath, ��������� ������ FSO
    ' ������� ����� �������������� � ��� ������, ���� SearchDeep > 1
    ' ��������� ���� ��������� ������ � ��������� FileNamesColl

    Dim curfold As Object

    Set curfold = fso.GetFolder(folderPath)

    If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����

        ' ���������������� ��� ������ ��� ������ ���� � ���������������
        ' � ������� ������ ����� � ������ ��������� Excel
        ' Application.StatusBar = "����� � �����: " & FolderPath

        Dim fil As Object

        For Each fil In curfold.Files        ' ���������� ��� ����� � ����� FolderPath

            With fil

                If .Name Like "*" & mask Then

                    If Coll_IN_v_q(FileNamesColl, .Path) = False Then

                        FileNamesColl.Add .Path

                    End If
                End If
            End With
        Next

        ' ��� ���������� �������� ��������� ����� ������������ StatusBar

        searchDeep = searchDeep - 1    ' ��������� ������� ������ � ���������

        Dim sfol As Object

        If searchDeep Then    ' ���� ���� ������ ������

            For Each sfol In curfold.SubFolders    ' ���������� ��� �������� � ����� FolderPath

                GetAllFileNamesUsingFSO _
                        sfol.Path, mask, fso, FileNamesColl, searchDeep

            Next
        End If
    End If
End Function

Function Coll_IN_v_q(Coll As Collection, msg As String) _
        As Boolean
    ' procedure Checked by test �������������
    ' ���� �� ������� � ��������� ?
    ' ��������

    Dim var As Variant

    For Each var In Coll

        If var = msg Then

            Coll_IN_v_q = True

            Exit For

        End If
    Next
End Function

Function �_��������_���������(Coll As Collection, _
        indx As Long) _
        As Boolean
    ' procedure Checked by test �������������
    '
    If indx > 0 And indx <= Coll.Count Then

        �_��������_��������� = True

    End If
End Function
