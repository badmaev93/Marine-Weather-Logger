Attribute VB_Name = "ModuleMain"
' ===================================
' �������� ������ ��� ������ � ��������
' ===================================
Option Explicit

' ���������
Private Const PASSWORD As String = "3timitimi3" ' ���������������� �������� �������

'  ��������� ������ ��� ���������� ������
Public Sub AddRecord()
    '  ������������� ����� ��� ����� ������
    With New UserForm1
        .Tag = "New"
        .Show vbModeless
    End With
End Sub

Public Sub EditRecord()
    On Error GoTo ErrorHandler
    
    ' ������������ ���� ������ ����� �������� ������ ������
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    ws.Activate
    
    ' ��������� ������ �������� ���� �� ������������
    Dim rowNum As String
    rowNum = InputBox("������� ����� ������ ��� ��������������:", "�������������� ������")
    
    ' �������� ������������ �����
    If rowNum = "" Then Exit Sub ' ������������ �������
    
    If Not IsNumeric(rowNum) Then
        MsgBox "������� ���������� ����� ������!", vbExclamation
        Exit Sub
    End If
    
    ' �������������� � ������� ������ � �������� ������������� �����
    Dim rowNumLong As Long
    rowNumLong = CLng(rowNum)
    
    If rowNumLong < 2 Or rowNumLong > ws.Cells(ws.Rows.Count, 1).End(xlUp).Row Then
        MsgBox "������ � ����� ������� �� ����������!", vbExclamation
        Exit Sub
    End If
    
    ' ��������� ������ ��� �����������
    ws.Rows(rowNumLong).Select
    
    ' �������� ����� �������������� - ������������� vbModeless ��� ���������� ���������
    With New UserForm1
        .Caption = "������������� ������"
        .Tag = CStr(rowNumLong)
        .Show vbModeless
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "������ ��� �������� �����: " & vbNewLine & Err.Description, vbCritical
End Sub
