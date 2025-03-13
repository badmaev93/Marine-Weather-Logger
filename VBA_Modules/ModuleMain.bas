Attribute VB_Name = "ModuleMain"
' ===================================
' Основной модуль для работы с записями
' ===================================
Option Explicit

' Константы
Private Const PASSWORD As String = "3timitimi3" ' Централизованное хранение паролей

'  Публичные методы для управления входом
Public Sub AddRecord()
    '  Инициализация формы для новой записи
    With New UserForm1
        .Tag = "New"
        .Show vbModeless
    End With
End Sub

Public Sub EditRecord()
    On Error GoTo ErrorHandler
    
    ' Активировать лист данных перед запросом номера строки
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    ws.Activate
    
    ' Получение номера целевого ряда от пользователя
    Dim rowNum As String
    rowNum = InputBox("Введите номер строки для редактирования:", "Редактирование записи")
    
    ' Проверка правильности ввода
    If rowNum = "" Then Exit Sub ' Пользователь отменил
    
    If Not IsNumeric(rowNum) Then
        MsgBox "Введите корректный номер строки!", vbExclamation
        Exit Sub
    End If
    
    ' Преобразование в длинные строки и проверка существования строк
    Dim rowNumLong As Long
    rowNumLong = CLng(rowNum)
    
    If rowNumLong < 2 Or rowNumLong > ws.Cells(ws.Rows.Count, 1).End(xlUp).Row Then
        MsgBox "Строка с таким номером не существует!", vbExclamation
        Exit Sub
    End If
    
    ' Выделение строки для наглядности
    ws.Rows(rowNumLong).Select
    
    ' Показать форму редактирования - использование vbModeless для разрешения прокрутки
    With New UserForm1
        .Caption = "Редактировать запись"
        .Tag = CStr(rowNumLong)
        .Show vbModeless
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при открытии формы: " & vbNewLine & Err.Description, vbCritical
End Sub
