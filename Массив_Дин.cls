VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "������_���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ������������ d2 ��������� ������ ����������
' � �� ������� � �� ��������

' � ������ ����� ����������:
' ������_��� As New ������_���
'Type �������_������
'    �������() As Variant
'End Type

' �������������:
' ������_���.������� �����, ��������  // ������� ������ ��������� �������
' ������_���.�_������ ������, �������, �������� // ��������� � ������ �������� � ���� ������ ��� ������� ������,  ��� ������� ����������� ���������
' d2 = ������_���.������ // ������� ������� ������
' ������_���.�������_�����_��������
' ������_���.������_�����_��������
' str = ������_���.������(������, �������) //  ������� ���������� ������

' ������ ���������� �������� ����������
Private �������_������() As �������_������    ' ������ ������� �������������


Function �������( _
        ����� As Long, _
        �������� As Long)
    ' procedure Checked by test �������������

    ' ������ � 1 ��� ��������, ���������� � �������

    ReDim �������_������(1 To �����)

    Dim y As Long

    For y = LBound(�������_������) To UBound(�������_������)

        ReDim �������_������(y).�������(1 To ��������)

    Next
End Function

Function �_������(������ As Long, _
        ������� As Long, _
        �������� As Variant)                    ' ������ �������
    ' ����������� ����������� "������"
    ' ��� �������������

    If ������ > UBound(�������_������) Then _
            ������_�����_�������� ������

    If ������� > UBound(�������_������(1).�������) Then _
            �������_�����_�������� �������

    �������_������(������).�������(�������) = ��������

End Function

Function ������(������ As Long, _
        ������� As Long) _
        As Variant                   ' procedure Checked by test �������������
    ' ������� ���������� ������ "������� �������"
    ' get/������

    ������ = �������_������(������).�������(�������)

End Function

Function ������_�����_��������(����� As Long)

    ReDim Preserve �������_������(1 To �����)

    Dim y As Long

    For y = LBound(�������_������) To �����

        ReDim Preserve �������_������(y).�������(1 To ��������)

    Next
End Function

Function �������_�����_��������(������� As Long)

    Dim y As Long

    For y = LBound(�������_������) To UBound(�������_������)

        ReDim Preserve �������_������(y).�������(1 To �������)

    Next
End Function

Function ������(Optional msg As String) _
        As Variant()   ' ������  �������
    ' ������� ������� d2 ������
    ' ���� ������� ��������� � �������������:
    ' �������_������(y).�������(x)

    Dim d2() As Variant
    ReDim d2(1 To UBound(�������_������), _
            1 To UBound(�������_������(1).�������))

    Dim y As Long, x As Long

    For y = LBound(d2) To UBound(d2)
        For x = LBound(d2, 2) To UBound(d2, 2)
            d2(y, x) = �������_������(y).�������(x)
        Next
    Next

    If IsArray(d2) Then _
            ������ = d2

End Function

Function ���������(Optional msg As String)
    '  ������ ������� �������������
    ' ��� �������
    Dim y As Long, x As Long

    For y = LBound(�������_������) To UBound(�������_������)
        For x = LBound(�������_������(y).�������) To UBound(�������_������(y).�������)
            �������_������(y).�������(x) = y & x
        Next
    Next
End Function

Function �����(����� As Long, �������� As Long) _
        As Variant   ' procedure Checked by test �������������

    ������_�����_�������� �����
    �������_�����_�������� ��������

End Function

Function �����(Optional msg As String) _
        As Long   ' procedure Checked by test �������������
    ' ������� ����� �����

    ����� = UBound(�������_������)
End Function

Function ��������(Optional msg As String) _
        As Long   ' procedure Checked by test �������������
    ' ������� ����� ��

    �������� = UBound(�������_������( _
            LBound(�������_������)).�������)
End Function

Function G_������_���(lRows As Long, _
        lCols As Long) _
        As ������_���
    ' procedure Checked by test �������������
    '
    Dim a2_dyn As New ������_���

    a2_dyn.������� lRows, lRows

    ������_���_��������� a2_dyn

    Set G_������_��� = a2_dyn

End Function

Function G_������_���_������(lRows As Long, _
        lCols As Long) _
        As ������_���
    ' ������  �������
    '
    Dim a2_Dyn_Empty As New ������_���

    a2_Dyn_Empty.������� lRows, lRows

    Set G_������_���_������ = a2_Dyn_Empty

End Function

Function ������_���_���������(a2_��� As ������_���)
    ' procedure Checked by test �������������

    With a2_���

        Dim y As Long, x As Long

        For y = 1 To .�����

            For x = 1 To .��������

                .�_������ y, x, Mock.G_S

            Next
        Next
    End With
End Function


