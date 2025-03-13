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



