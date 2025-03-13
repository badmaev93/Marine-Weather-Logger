VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ===================================
' Обработчики событий книги Excel
' ===================================


Private Sub Workbook_Open()
    On Error Resume Next
    
    ' Сначала снять защиту с листа данных, если он уже защищен
    ThisWorkbook.Sheets("Data").Unprotect PASSWORD:=PASSWORD
    
    ' Проверить/обновить структуру столбцов для Ice данных
    UpdateDataSheetStructure
    
    ' Применить защиту с правильными настройками для работы VBA
    ThisWorkbook.Sheets("Data").Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    
    On Error GoTo 0
End Sub

' Обновление структуры листа данных - переименование/создание столбцов при необходимости
Private Sub UpdateDataSheetStructure()
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Переименование заголовков Ice columns, если они существуют в старом формате
    Dim headerRow As Range
    Set headerRow = ws.Rows(1)
    
    ' Для корректной работы с данными в прямом коде, а не текстовых значениях
    If headerRow.Cells(1, 16).value = "Ice score (code)" Then
        ' Переименовать столбцы
        headerRow.Cells(1, 13).value = "Old Ice score"
        headerRow.Cells(1, 14).value = "Old Ice type"
        headerRow.Cells(1, 15).value = "Old Ice shape"
        
        headerRow.Cells(1, 16).value = "Ice score"
        headerRow.Cells(1, 17).value = "Ice type"
        headerRow.Cells(1, 18).value = "Ice shape"
        
        ' Перенести коды в основные столбцы для всех существующих записей
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        If lastRow > 1 Then
            Dim i As Long
            For i = 2 To lastRow
                ' Копировать коды в основные столбцы
                ws.Cells(i, 13).value = ws.Cells(i, 16).value
                ws.Cells(i, 14).value = ws.Cells(i, 17).value
                ws.Cells(i, 15).value = ws.Cells(i, 18).value
            Next i
        End If
    End If
    
    On Error GoTo 0
End Sub


