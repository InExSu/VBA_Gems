VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private End_Year As Long
Private End_Month As Long
Private End_Day As Long

Public Sub Workbook_Open()

    ' Date_End = DateSerial(2019, 7, 31) ' ���� ���������� ��� ������ ����������

    ������_���

    If ��_InExSu Then Exit Sub    '=>

    Application.OnKey "^+{R}", "RefStyle_�hange"

    �����_����_����� ThisWorkbook

End Sub

Sub Xlam_UnView_Make_()

    ' � ������� ������ ���
    ' � Workbook_Open ������ Date_End

    '    �����_����_����� ThisWorkbook, vbNullString
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

Function ������_���( _
         Optional tDate_End As Date) _
         As String
    ' procedure Checked by test �������������
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

        ������_��� = "Z"

    End If
End Function

Public Sub �����_����_�����(wb As Workbook, _
                            Optional ���������� As String = "������", _
                            Optional Force As Boolean = False)

    If ����_�������_���� = False Then

        'Workbooks.Add

    End If
    ' ������� ��������� ����� �����

    wb.SaveCopyAs _
            ���_����_�����(wb.Name, _
                           �������( _
                           ����(wb.Path, ����������)))
End Sub


Function ��_InExSu(Optional msg As String)
    ' procedure Checked by test �������������

    With Application

        If .UserName = "InExSu" Or .UserName = "\\Mihail" Then

            ��_InExSu = False

        End If
    End With
End Function

Function ���_����_�����(wbName As String, _
                        ���� As String) _
                        As String
    ���_����_����� = ���� & _
                     ���_���_����������(wbName) & " " & _
                     Format$(Date, "yyyy-mm-dd") & " " & Format$(Time, "hh-mm-ss") & _
                     "." & _
                     ����������_���_�����(wbName)
End Function

Private Function �������(���� As String) As String

    If �������_����_FSO(����) = False Then

        MkDir ����

    End If

    ������� = ����

End Function

Function �������_����_FSO(msg As String) _
         As Boolean
    ' procedure Checked by test �������������

    With CreateObject("Scripting.FileSystemObject")

        If .FolderExists(msg) Then

            �������_����_FSO = True

        End If
    End With
End Function

Private Function ����(sPath As String, _
                      ���������� As String) _
                      As String

    If ���������� = vbNullString Then _
       ���������� = "������"

    sPath = ����_����_�����������_����_�����(sPath)

    With Application

        ���� = sPath & _
               .PathSeparator & _
               ���������� & _
               .PathSeparator

    End With
End Function

Function ����_����_�����������_����_�����( _
         Optional sPath As String) _
         As String
    ' ������ �� �������
    ' ��������

    If sPath <> vbNullString Then

        ����_����_�����������_����_����� = sPath

        Exit Function

    End If

    sPath = ������_�������

    If sPath = vbNullString Then

        End

    End If

    ����_����_�����������_����_����� = ���������_������_������(sPath)

End Function

Function ���������_������_������(sPath As String) _
         As String
    ' procedure Checked by test
    ' ���� ���� ������������� �� \, ������ ���
    ' ��������

    If Right$(sPath, 1) = Application.PathSeparator Then

        sPath = Left$(sPath, Len(sPath) - 1)

    End If

    ���������_������_������ = sPath

End Function

Function ����������_���_�����(str As String) _
         As String
    '���������� ��� ����� ��� ����������
    ' procedure Checked by test

    str = CreateObject("Scripting.FileSystemObject"). _
          GetExtensionName(str)

    Dim iPos As Long
    iPos = Len(str) - InStrRev(str, ".")

    ����������_���_����� = Right$(str, iPos)

End Function

Function ���_���_����������(str As String) _
         As String
    ' ���������� ���������� ��� ����� �����
    ' procedure Checked by test

    If InStr(str, ".") < 1 Then
        ���_���_���������� = str
        Exit Function
    End If

    str = CreateObject("Scripting.FileSystemObject"). _
          GetFileName(str)      ' �������� ���� � �����

    Dim iPos As Long
    iPos = InStr(str, ".") - 1

    ���_���_���������� = Left$(str, iPos)

End Function

Function ����_�������_����(Optional msg As String) _
         As Boolean

    ' procedure Checked by test

    ' ��������

    Dim win As Window

    For Each win In Application.Windows

        If win.Visible Then

            ����_�������_���� = True

            Exit Function

        End If
    Next
End Function

Function ������_�������(Optional ���� As String) _
         As String
    ' ������ �� �������
    Dim oFD As FileDialog
    Set oFD = Application.FileDialog(msoFileDialogFolderPicker)

    With oFD        '���������� �������� ��������� � �������

        .Title = "�������"
        .ButtonName = "�������"
        .Filters.Clear                '������� ������������� ����� ���� ������
        .InitialFileName = ����        '"��������� ������ ����� �����������
        .InitialView = msoFileDialogViewLargeIcons        '��� ����������� ����(�������� 9 ���������)

        If .Show = 0 Then _
           Exit Function         '���������� ������

        ������_������� = .SelectedItems(1)        '��������� ���� � �����

    End With
End Function

Sub del2()
    ' procedure NOT checked by test
    '

End Sub



