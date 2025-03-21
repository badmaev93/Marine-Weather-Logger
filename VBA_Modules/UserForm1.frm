﻿VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Add Entry / Добавить запись"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13095
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ===================================
' Форма ввода данных
' ===================================

Option Explicit

Private Const COORD_FORMAT_DECIMAL As Boolean = False
Private Const COORD_FORMAT_DEGREES As Boolean = True

Private Const PASSWORD As String = "" ' !!!!!!!!!!!!!!!!!!1Установить  пароль!!!!!!!!!!!!

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As LongPtr, ByVal wMsg As Long, _
     ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, _
     ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr

Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const CB_SHOWDROPDOWN As Long = &H14F

Private Type CoordInput
    degrees As MSForms.TextBox
    minutes As MSForms.TextBox
    direction As MSForms.ComboBox
End Type

Private mCoordFormat As Boolean
Private mIsCalm As Boolean
Private mIsPort As Boolean
Private LatitudeInput As CoordInput
Private LongitudeInput As CoordInput
Private mIsIceNotated As Boolean

Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Starting UserForm_Initialize ==="
    Debug.Print "Form Tag: " & Me.Tag
    
    InitializeCoordinateFields
    InitializeControls
    InitializeIceControls
    
    mCoordFormat = COORD_FORMAT_DEGREES
    UpdateCoordinateControls
    
    If Me.Tag = "" Then
        Debug.Print "Empty tag - setting to New"
        Me.Tag = "New"
        SetDefaultValues
    End If
    
    Debug.Print "=== UserForm_Initialize completed ==="
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR in UserForm_Initialize: " & Err.Description
    MsgBox "Ошибка инициализации формы: " & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub UserForm_Activate()
    Debug.Print "Form Activated. Tag = " & Me.Tag
    
    If IsNumeric(Me.Tag) And Me.Tag <> "New" Then
        LoadExistingData CLng(Me.Tag)
    End If
End Sub

Private Sub UserForm_Terminate()
    On Error Resume Next
    ThisWorkbook.Sheets("Data").Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub InitializeCoordinateFields()
    With Me.fraMain.fraCoordinates

        Set LatitudeInput.degrees = .txtLatDegrees
        Set LatitudeInput.minutes = .txtLatMinutes
        Set LatitudeInput.direction = .cboLatDirection
        
        Set LongitudeInput.degrees = .txtLonDegrees
        Set LongitudeInput.minutes = .txtLonMinutes
        Set LongitudeInput.direction = .cboLonDirection
    End With
End Sub

Private Sub InitializeControls()
    On Error GoTo ErrorHandler
    
    With Me

        .optDecimalCoords.value = False
        .optDegreeCoords.value = True

        ClearAllFields

        InitializeIceControls
        InitializeDirectionControls

        .chkIceNotated = False
        .chkSeaSwell.value = True
    End With
    
    Exit Sub

ErrorHandler:
    Debug.Print "Error in InitializeControls: " & Err.Description
    Err.Raise Err.Number, "InitializeControls", _
              "Ошибка инициализации элементов управления."
End Sub

Private Sub InitializeDirectionControls()

    With LatitudeInput.direction
        .Clear
        .AddItem "N"
        .AddItem "S"
        .Text = "N"
    End With
    
    With LongitudeInput.direction
        .Clear
        .AddItem "E"
        .AddItem "W"
        .Text = "E"
    End With
End Sub

Private Sub InitializeIceControls()
    On Error GoTo ErrorHandler
    
    Dim wsIceScore As Worksheet
    Dim wsIceType As Worksheet
    Dim wsIceShape As Worksheet
    Set wsIceScore = ThisWorkbook.Sheets("IceScore")
    Set wsIceType = ThisWorkbook.Sheets("IceType")
    Set wsIceShape = ThisWorkbook.Sheets("IceShape")

    With Me.cboIceScore
        .Clear
        LoadComboBoxData wsIceScore, .Name
        .TextColumn = 1
        .BoundColumn = 2
        .ColumnWidths = "200;0"
        .Style = fmStyleDropDownList
    End With
    
    With Me.cboIceType
        .Clear
        LoadComboBoxData wsIceType, .Name
        .TextColumn = 1
        .BoundColumn = 2
        .ColumnWidths = "200;0"
        .Style = fmStyleDropDownList
    End With
    
    With Me.cboIceShape
        .Clear
        LoadComboBoxData wsIceShape, .Name
        .TextColumn = 1
        .BoundColumn = 2
        .ColumnWidths = "200;0"
        .Style = fmStyleDropDownList
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при инициализации данных льда: " & vbNewLine & Err.Description, vbCritical
End Sub

Private Sub LoadExistingData(ByVal rowNum As Long)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Starting LoadExistingData for row " & rowNum & " ==="
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    If ws Is Nothing Then
        Debug.Print "ERROR: Data sheet not found"
        Exit Sub
    End If
    
    If ws.Cells(rowNum, 1).value = "" Then
        Debug.Print "ERROR: Row " & rowNum & " is empty"
        Exit Sub
    End If
    
    Debug.Print "Reading values from row " & rowNum & ":"
    Debug.Print "Date/Time: " & ws.Cells(rowNum, 1).value
    Debug.Print "Latitude: " & ws.Cells(rowNum, 2).value
    Debug.Print "Longitude: " & ws.Cells(rowNum, 3).value
    
    With Me
        ClearAllFields
        
        .txtDateTime1.value = Format(ws.Cells(rowNum, 1).value, "dd.mm.yyyy hh:00")
        
        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            .fraMain.fraCoordinates.txtLatitude.Text = FormatNumber(ws.Cells(rowNum, 2).value, 4)
            .fraMain.fraCoordinates.txtLongitude.Text = FormatNumber(ws.Cells(rowNum, 3).value, 4)
        Else
            ConvertToDegreesMinutes CDbl(ws.Cells(rowNum, 2).value), _
                                  .fraMain.fraCoordinates.txtLatDegrees, _
                                  .fraMain.fraCoordinates.txtLatMinutes, _
                                  .fraMain.fraCoordinates.cboLatDirection, _
                                  True
                                  
            ConvertToDegreesMinutes CDbl(ws.Cells(rowNum, 3).value), _
                                  .fraMain.fraCoordinates.txtLonDegrees, _
                                  .fraMain.fraCoordinates.txtLonMinutes, _
                                  .fraMain.fraCoordinates.cboLonDirection, _
                                  False
        End If

        .txtTemp.Text = ws.Cells(rowNum, 4).Text
        .txtBarometer.Text = ws.Cells(rowNum, 5).Text
        .txtVisibility.Text = ws.Cells(rowNum, 6).Text
        .txtWindDirection.Text = ws.Cells(rowNum, 7).Text
        .txtWindSpeed.Text = ws.Cells(rowNum, 8).Text
        .txtSeaSwellDirection.Text = ws.Cells(rowNum, 9).Text
        .txtSeaSwell.Text = ws.Cells(rowNum, 10).Text
        .txtWindWaveDirection.Text = ws.Cells(rowNum, 11).Text
        .txtWindWaveHeight.Text = ws.Cells(rowNum, 12).Text

        If ws.Cells(rowNum, 13).Text = "CW" Then
            .chkIceNotated.value = False
        Else
            .chkIceNotated.value = True

            FindAndSelectComboValueByCode .cboIceScore, ws.Cells(rowNum, 13).Text
            FindAndSelectComboValueByCode .cboIceType, ws.Cells(rowNum, 14).Text
            FindAndSelectComboValueByCode .cboIceShape, ws.Cells(rowNum, 15).Text
        End If

        UpdateSeaControls
        UpdateCoordinateControls
        
        Debug.Print "Data loaded successfully"
    End With
    
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR in LoadExistingData: " & Err.Description
    Debug.Print "Error Line: " & Erl
    Resume Next
End Sub

Private Sub LoadComboBoxData(ws As Worksheet, comboName As String)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub  ' No data beyond header

    Dim dataRange As Range
    Set dataRange = ws.Range("A2:B" & lastRow)

    Select Case comboName
        Case "cboIceScore"
            Me.cboIceScore.List = dataRange.value
        Case "cboIceType"
            Me.cboIceType.List = dataRange.value
        Case "cboIceShape"
            Me.cboIceShape.List = dataRange.value
    End Select
End Sub

Private Function GetTwoColumnValues(ws As Worksheet) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Dim dataArray() As Variant
    ReDim dataArray(1 To lastRow - 1, 1 To 2)

    Dim i As Long
    For i = 2 To lastRow
        dataArray(i - 1, 1) = ws.Cells(i, "A").value
        dataArray(i - 1, 2) = ws.Cells(i, "B").value
    Next i
    
    GetTwoColumnValues = dataArray
End Function

Private Sub FindAndSelectComboValue(cmb As MSForms.ComboBox, value As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i, 0) = value Then
            cmb.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub FindAndSelectComboValueByCode(cmb As MSForms.ComboBox, code As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i, 1) = code Then
            cmb.ListIndex = i
            Exit For
        End If
    Next i
End Sub

Private Sub ClearAllFields()
    With Me
        .txtLongitude.Text = ""
        .txtLatitude.Text = ""
        .txtTemp.Text = ""
        .txtBarometer.Text = ""
        .txtWindDirection.Text = ""
        .txtWindSpeed.Text = ""
        .txtVisibility.Text = ""
        .txtSeaSwell.Text = ""
        .txtSeaSwellDirection.Text = ""
        .txtWindWaveDirection.Text = ""
        .txtWindWaveHeight.Text = ""
        
        If Not LatitudeInput.degrees Is Nothing Then LatitudeInput.degrees.Text = ""
        If Not LatitudeInput.minutes Is Nothing Then LatitudeInput.minutes.Text = ""
        If Not LongitudeInput.degrees Is Nothing Then LongitudeInput.degrees.Text = ""
        If Not LongitudeInput.minutes Is Nothing Then LongitudeInput.minutes.Text = ""
    End With
End Sub

Private Sub optDecimalCoords_Click()
    mCoordFormat = COORD_FORMAT_DECIMAL
    UpdateCoordinateControls
    ConvertAndUpdateCoordinates
End Sub

Private Sub optDegreeCoords_Click()
    mCoordFormat = COORD_FORMAT_DEGREES
    UpdateCoordinateControls
    ConvertAndUpdateCoordinates
End Sub

Private Sub chkIceNotated_Click()
    mIsIceNotated = Me.chkIceNotated.value
    UpdateSeaControls
End Sub

Private Sub chkSeaSwell_Click()
    UpdateSeaControls
End Sub

Private Sub txtLatitude_Click()
    If Not Me.fraMain.fraCoordinates.txtLatitude.Enabled Then
        optDecimalCoords.value = True
    End If
End Sub

Private Sub txtLongitude_Click()
    If Not Me.fraMain.fraCoordinates.txtLongitude.Enabled Then
        optDecimalCoords.value = True
    End If
End Sub

Private Sub txtLatDegrees_Click()
    If Not LatitudeInput.degrees.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub

Private Sub txtLonDegrees_Click()
    If Not LongitudeInput.degrees.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub

Private Sub txtLatMinutes_Click()
    If Not LatitudeInput.minutes.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub

Private Sub txtLonMinutes_Click()
    If Not LongitudeInput.minutes.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub

Private Sub cboLatDirection_Click()
    If Not LatitudeInput.direction.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub

Private Sub cboLonDirection_Click()
    If Not LongitudeInput.direction.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub

Private Sub UpdateCoordinateControls()
    Dim activeBackColor As Long, inactiveBackColor As Long
    Dim activeTextColor As Long, inactiveTextColor As Long
    
    activeBackColor = vbWhite
    inactiveBackColor = RGB(240, 240, 240)
    activeTextColor = vbBlack
    inactiveTextColor = RGB(192, 192, 192)
    
    With Me.fraMain.fraCoordinates
        If mCoordFormat = COORD_FORMAT_DECIMAL Then

            .txtLatitude.BackColor = activeBackColor
            .txtLongitude.BackColor = activeBackColor
            .txtLatitude.ForeColor = activeTextColor
            .txtLongitude.ForeColor = activeTextColor
            .lblLatitude.ForeColor = activeTextColor
            .lblLongitude.ForeColor = activeTextColor
            .txtLatitude.Locked = False
            .txtLongitude.Locked = False
            .txtLatitude.Enabled = True
            .txtLongitude.Enabled = True

            .txtLatDegrees.BackColor = inactiveBackColor
            .txtLatMinutes.BackColor = inactiveBackColor
            .cboLatDirection.BackColor = inactiveBackColor
            .txtLonDegrees.BackColor = inactiveBackColor
            .txtLonMinutes.BackColor = inactiveBackColor
            .cboLonDirection.BackColor = inactiveBackColor
            
            .txtLatDegrees.ForeColor = inactiveTextColor
            .txtLatMinutes.ForeColor = inactiveTextColor
            .cboLatDirection.ForeColor = inactiveTextColor
            .txtLonDegrees.ForeColor = inactiveTextColor
            .txtLonMinutes.ForeColor = inactiveTextColor
            .cboLonDirection.ForeColor = inactiveTextColor

            .txtLatDegrees.Locked = True
            .txtLatMinutes.Locked = True
            .txtLonDegrees.Locked = True
            .txtLonMinutes.Locked = True
            .cboLatDirection.Locked = True
            .cboLonDirection.Locked = True
            
            .txtLatDegrees.Enabled = False
            .txtLatMinutes.Enabled = False
            .txtLonDegrees.Enabled = False
            .txtLonMinutes.Enabled = False
            .cboLatDirection.Enabled = False
            .cboLonDirection.Enabled = False
            
        Else

            .txtLatDegrees.BackColor = activeBackColor
            .txtLatMinutes.BackColor = activeBackColor
            .cboLatDirection.BackColor = activeBackColor
            .txtLonDegrees.BackColor = activeBackColor
            .txtLonMinutes.BackColor = activeBackColor
            .cboLonDirection.BackColor = activeBackColor
            
            .txtLatDegrees.ForeColor = activeTextColor
            .txtLatMinutes.ForeColor = activeTextColor
            .cboLatDirection.ForeColor = activeTextColor
            .txtLonDegrees.ForeColor = activeTextColor
            .txtLonMinutes.ForeColor = activeTextColor
            .cboLonDirection.ForeColor = activeTextColor
            
            .txtLatDegrees.Locked = False
            .txtLatMinutes.Locked = False
            .txtLonDegrees.Locked = False
            .txtLonMinutes.Locked = False
            .cboLatDirection.Locked = False
            .cboLonDirection.Locked = False
            
            .txtLatDegrees.Enabled = True
            .txtLatMinutes.Enabled = True
            .txtLonDegrees.Enabled = True
            .txtLonMinutes.Enabled = True
            .cboLatDirection.Enabled = True
            .cboLonDirection.Enabled = True

            .txtLatitude.BackColor = inactiveBackColor
            .txtLongitude.BackColor = inactiveBackColor
            .txtLatitude.ForeColor = inactiveTextColor
            .txtLongitude.ForeColor = inactiveTextColor
            .lblLatitude.ForeColor = inactiveTextColor
            .lblLongitude.ForeColor = inactiveTextColor
            .txtLatitude.Locked = True
            .txtLongitude.Locked = True
            .txtLatitude.Enabled = False
            .txtLongitude.Enabled = False
        End If
    End With
End Sub

Private Sub UpdateSeaControls()
    Dim activeBackColor As Long, inactiveBackColor As Long
    Dim activeTextColor As Long, inactiveTextColor As Long
    
    activeBackColor = vbWhite
    inactiveBackColor = RGB(240, 240, 240)
    activeTextColor = vbBlack
    inactiveTextColor = RGB(192, 192, 192)
    
    With Me

        If .chkSeaSwell.value Then

            .txtSeaSwell.BackColor = activeBackColor
            .txtSeaSwellDirection.BackColor = activeBackColor
            .txtWindWaveDirection.BackColor = activeBackColor
            .txtWindWaveHeight.BackColor = activeBackColor
            
            .txtSeaSwell.ForeColor = activeTextColor
            .txtSeaSwellDirection.ForeColor = activeTextColor
            .txtWindWaveDirection.ForeColor = activeTextColor
            .txtWindWaveHeight.ForeColor = activeTextColor
            
            .lblSeaSwell.ForeColor = activeTextColor
            .lblSeaSwellDirection.ForeColor = activeTextColor
            .lblWindWaveDirection.ForeColor = activeTextColor
            .lblWindWaveHeight.ForeColor = activeTextColor
            
            .txtSeaSwell.Enabled = True
            .txtSeaSwellDirection.Enabled = True
            .txtWindWaveDirection.Enabled = True
            .txtWindWaveHeight.Enabled = True
            
            .txtSeaSwell.Locked = False
            .txtSeaSwellDirection.Locked = False
            .txtWindWaveDirection.Locked = False
            .txtWindWaveHeight.Locked = False

            If .txtSeaSwell.Text = "0" Then .txtSeaSwell.Text = ""
            If .txtSeaSwellDirection.Text = "0" Then .txtSeaSwellDirection.Text = ""
            If .txtWindWaveDirection.Text = "0" Then .txtWindWaveDirection.Text = ""
            If .txtWindWaveHeight.Text = "0" Then .txtWindWaveHeight.Text = ""
        Else

            .txtSeaSwell.BackColor = inactiveBackColor
            .txtSeaSwellDirection.BackColor = inactiveBackColor
            .txtWindWaveDirection.BackColor = inactiveBackColor
            .txtWindWaveHeight.BackColor = inactiveBackColor
            
            .txtSeaSwell.ForeColor = inactiveTextColor
            .txtSeaSwellDirection.ForeColor = inactiveTextColor
            .txtWindWaveDirection.ForeColor = inactiveTextColor
            .txtWindWaveHeight.ForeColor = inactiveTextColor
            
            .lblSeaSwell.ForeColor = inactiveTextColor
            .lblSeaSwellDirection.ForeColor = inactiveTextColor
            .lblWindWaveDirection.ForeColor = inactiveTextColor
            .lblWindWaveHeight.ForeColor = inactiveTextColor
            
            .txtSeaSwell.Enabled = False
            .txtSeaSwellDirection.Enabled = False
            .txtWindWaveDirection.Enabled = False
            .txtWindWaveHeight.Enabled = False
            
            .txtSeaSwell.Text = "0"
            .txtSeaSwellDirection.Text = "0"
            .txtWindWaveDirection.Text = "0"
            .txtWindWaveHeight.Text = "0"
        End If

        If .chkIceNotated.value Then

            .cboIceType.BackColor = activeBackColor
            .cboIceScore.BackColor = activeBackColor
            .cboIceShape.BackColor = activeBackColor
            
            .cboIceType.ForeColor = activeTextColor
            .cboIceScore.ForeColor = activeTextColor
            .cboIceShape.ForeColor = activeTextColor
            
            .lblIceType.ForeColor = activeTextColor
            .lblIceScore.ForeColor = activeTextColor
            .lblIceShape.ForeColor = activeTextColor
            
            .cboIceType.Enabled = True
            .cboIceScore.Enabled = True
            .cboIceShape.Enabled = True

            If .cboIceType.Text = "Чистая вода" Then .cboIceType.ListIndex = -1
            If .cboIceScore.Text = "Чистая вода" Then .cboIceScore.ListIndex = -1
            If .cboIceShape.Text = "Чистая вода" Then .cboIceShape.ListIndex = -1
        Else

            .cboIceType.BackColor = inactiveBackColor
            .cboIceScore.BackColor = inactiveBackColor
            .cboIceShape.BackColor = inactiveBackColor
            
            .cboIceType.ForeColor = inactiveTextColor
            .cboIceScore.ForeColor = inactiveTextColor
            .cboIceShape.ForeColor = inactiveTextColor
            
            .lblIceType.ForeColor = inactiveTextColor
            .lblIceScore.ForeColor = inactiveTextColor
            .lblIceShape.ForeColor = inactiveTextColor
            
            .cboIceType.Enabled = False
            .cboIceScore.Enabled = False
            .cboIceShape.Enabled = False

            .cboIceType.Text = "Чистая вода"
            .cboIceScore.Text = "Чистая вода"
            .cboIceShape.Text = "Чистая вода"
        End If
    End With
End Sub

Private Sub cboIceScore_DropDown()
    EnableMouseWheel Me.cboIceScore
End Sub

Private Sub cboIceType_DropDown()
    EnableMouseWheel Me.cboIceType
End Sub

Private Sub cboIceShape_DropDown()
    EnableMouseWheel Me.cboIceShape
End Sub

Private Sub EnableMouseWheel(cmb As MSForms.ComboBox)
    Dim hwndList As LongPtr
    hwndList = FindWindowEx(cmb.hwnd, 0, "ComboBox", vbNullString)
    If hwndList <> 0 Then
        SendMessage hwndList, WM_MOUSEWHEEL, 0, 0
    End If
End Sub

Private Sub ConvertAndUpdateCoordinates()
    On Error GoTo ErrorHandler
    
    With Me
        If mCoordFormat = COORD_FORMAT_DECIMAL Then

            If LatitudeInput.degrees.Text <> "" And LatitudeInput.minutes.Text <> "" Then
                Dim latVal As Double
                latVal = ConvertToDecimal(LatitudeInput.degrees.Text, _
                                        LatitudeInput.minutes.Text, _
                                        LatitudeInput.direction.Text)
                
                .txtLatitude.Text = FormatCoordinate(latVal)
            End If
            
            If LongitudeInput.degrees.Text <> "" And LongitudeInput.minutes.Text <> "" Then
                Dim lonVal As Double
                lonVal = ConvertToDecimal(LongitudeInput.degrees.Text, _
                                        LongitudeInput.minutes.Text, _
                                        LongitudeInput.direction.Text)
                
                .txtLongitude.Text = FormatCoordinate(lonVal)
            End If
        Else

            If .txtLatitude.Text <> "" Then
                ConvertToDegreesMinutes CDbl(Replace(.txtLatitude.Text, ".", ",")), _
                                      LatitudeInput.degrees, _
                                      LatitudeInput.minutes, _
                                      LatitudeInput.direction, _
                                      True
            End If
            
            If .txtLongitude.Text <> "" Then
                ConvertToDegreesMinutes CDbl(Replace(.txtLongitude.Text, ".", ",")), _
                                      LongitudeInput.degrees, _
                                      LongitudeInput.minutes, _
                                      LongitudeInput.direction, _
                                      False
            End If
        End If
    End With
    Exit Sub

ErrorHandler:
    Debug.Print "Error in ConvertAndUpdateCoordinates: " & Err.Description
End Sub

Private Sub ConvertToDegreesMinutes(ByVal decimalValue As Double, _
                                  degreesBox As MSForms.TextBox, _
                                  minutesBox As MSForms.TextBox, _
                                  directionBox As MSForms.ComboBox, _
                                  ByVal isLatitude As Boolean)

    Dim isNegative As Boolean
    isNegative = (decimalValue < 0)
    decimalValue = Abs(decimalValue)

    Dim degrees As Long
    Dim minutes As Double
    
    degrees = Int(decimalValue)
    minutes = (decimalValue - degrees) * 60
    minutes = Round(minutes, 1) ' Округляем до 1 знака

    If minutes >= 60 Then
        degrees = degrees + 1
        minutes = 0
    End If

    Dim minutesStr As String
    minutesStr = Trim(Str(minutes))

    If InStr(minutesStr, ".") = 0 Then
        minutesStr = minutesStr & ".0"
    End If

    degreesBox.Text = CStr(degrees)
    minutesBox.Text = minutesStr

    If isLatitude Then
        directionBox.Text = IIf(isNegative, "S", "N")
    Else
        directionBox.Text = IIf(isNegative, "W", "E")
    End If
End Sub

Private Function ConvertToDecimal(ByVal degrees As String, ByVal minutes As String, ByVal direction As String) As Double

    degrees = Trim(degrees)
    minutes = Trim(minutes)
    direction = Trim(direction)

    If degrees = "" Or minutes = "" Or direction = "" Then
        ConvertToDecimal = 0
        Exit Function
    End If

    minutes = Replace(minutes, ",", ".")

    Dim deg As Double, min As Double
    deg = Val(degrees)
    min = Val(minutes)

    ConvertToDecimal = deg + (min / 60)

    If direction = "S" Or direction = "W" Then
        ConvertToDecimal = -ConvertToDecimal
    End If
End Function
Private Sub ConvertDecimalToMinutes()
    On Error Resume Next

    If Me.txtLatitude.Text <> "" And IsNumeric(Me.txtLatitude.Text) Then
        Dim latValue As Double
        latValue = Val(Me.txtLatitude.Text)
        
        If Abs(latValue) <= 90 Then

            Dim latDegrees As Long
            Dim latMinutes As Double
            Dim latDirection As String
            
            latDirection = IIf(latValue < 0, "S", "N")
            latValue = Abs(latValue)
            
            latDegrees = Int(latValue)
            latMinutes = (latValue - latDegrees) * 60

            Dim latMinutesStr As String
            latMinutesStr = Format(latMinutes, "0.0")

            LatitudeInput.degrees.Text = CStr(latDegrees)
            LatitudeInput.minutes.Text = latMinutesStr
            LatitudeInput.direction.Text = latDirection
        End If
    End If

    If Me.txtLongitude.Text <> "" And IsNumeric(Me.txtLongitude.Text) Then
        Dim lonValue As Double
        lonValue = Val(Me.txtLongitude.Text)
        
        If Abs(lonValue) <= 180 Then

            Dim lonDegrees As Long
            Dim lonMinutes As Double
            Dim lonDirection As String
            
            lonDirection = IIf(lonValue < 0, "W", "E")
            lonValue = Abs(lonValue)
            
            lonDegrees = Int(lonValue)
            lonMinutes = (lonValue - lonDegrees) * 60

            Dim lonMinutesStr As String
            lonMinutesStr = Format(lonMinutes, "0.0")

            LongitudeInput.degrees.Text = CStr(lonDegrees)
            LongitudeInput.minutes.Text = lonMinutesStr
            LongitudeInput.direction.Text = lonDirection
        End If
    End If
End Sub

Private Function GetDecimalCoordinates(degrees As String, minutes As String, direction As String) As Double

    If degrees = "" Or minutes = "" Or direction = "" Then
        GetDecimalCoordinates = 0
        Exit Function
    End If

    Dim deg As Double, min As Double
    deg = Val(degrees)
    min = Val(minutes)
    
    GetDecimalCoordinates = deg + (min / 60)

    If direction = "S" Or direction = "W" Then
        GetDecimalCoordinates = -GetDecimalCoordinates
    End If
End Function

Private Function FormatCoordinate(ByVal value As Double) As String
    Dim result As String

    result = Trim(Str(Abs(value)))

    If InStr(result, ".") > 0 Then

        Dim decimalPart As String
        decimalPart = Mid(result, InStr(result, ".") + 1)
        
        While Len(decimalPart) < 4
            result = result & "0"
            decimalPart = Mid(result, InStr(result, ".") + 1)
        Wend
    Else

        result = result & ".0000"
    End If

    If value < 0 Then
        result = "-" & result
    End If
    
    FormatCoordinate = result
End Function
Private Sub txtLatitude_Change()
    Static isProcessing As Boolean
    If isProcessing Then Exit Sub
    
    isProcessing = True

    If InStr(Me.txtLatitude.Text, ",") > 0 Then
        Dim curPos As Integer
        curPos = Me.txtLatitude.selStart
        
        Me.txtLatitude.Text = Replace(Me.txtLatitude.Text, ",", ".")
        Me.txtLatitude.selStart = curPos
    End If

    If Not Me.txtLatitude.Locked And Me.txtLatitude.Text <> "" And _
       Me.txtLatitude.Text <> "-" And Me.txtLatitude.Text <> "." And _
       Me.txtLatitude.Text <> "-." Then
        ConvertCoordinates
    End If
    
    isProcessing = False
End Sub

Private Sub txtLongitude_Change()
    Static isProcessing As Boolean
    If isProcessing Then Exit Sub
    
    isProcessing = True

    If InStr(Me.txtLongitude.Text, ",") > 0 Then
        Dim curPos As Integer
        curPos = Me.txtLongitude.selStart
        
        Me.txtLongitude.Text = Replace(Me.txtLongitude.Text, ",", ".")
        Me.txtLongitude.selStart = curPos
    End If

    If Not Me.txtLongitude.Locked And Me.txtLongitude.Text <> "" And _
       Me.txtLongitude.Text <> "-" And Me.txtLongitude.Text <> "." And _
       Me.txtLongitude.Text <> "-." Then
        ConvertCoordinates
    End If
    
    isProcessing = False
End Sub
Private Sub ConvertCoordinates()
    If mCoordFormat = COORD_FORMAT_DECIMAL Then

        ConvertDecimalToMinutes
    Else

        ConvertDegreesToDecimal
    End If
End Sub
Private Sub txtLatDegrees_Change()
    If Me.txtLatDegrees.Text <> "" And Me.txtLatMinutes.Text <> "" And _
       Me.txtLatMinutes.Text <> "." And LatitudeInput.direction.Text <> "" Then
        ConvertCoordinates
    End If
End Sub

Private Sub txtLatMinutes_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    Me.txtLatMinutes.Text = Replace(Me.txtLatMinutes.Text, ",", ".")
End Sub

Private Sub txtLonMinutes_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)

    Me.txtLonMinutes.Text = Replace(Me.txtLonMinutes.Text, ",", ".")
End Sub

Private Sub txtLatMinutes_Change()
    Static isProcessing As Boolean
    If isProcessing Then Exit Sub
    
    isProcessing = True

    If InStr(Me.txtLatMinutes.Text, ",") > 0 Then
        Dim curPos As Integer
        curPos = Me.txtLatMinutes.selStart
        
        Me.txtLatMinutes.Text = Replace(Me.txtLatMinutes.Text, ",", ".")
        Me.txtLatMinutes.selStart = curPos
    End If

    Me.txtLatMinutes.ForeColor = RGB(0, 0, 0) ' Всегда черный текст

    If Me.txtLatMinutes.Text <> "" And Me.txtLatMinutes.Text <> "." And _
       LatitudeInput.degrees.Text <> "" And LatitudeInput.direction.Text <> "" Then
        ConvertCoordinates
    End If
    
    isProcessing = False
End Sub

Private Sub txtLonDegrees_Change()
    If Me.txtLonDegrees.Text <> "" And Me.txtLonMinutes.Text <> "" And _
       Me.txtLonMinutes.Text <> "." And LongitudeInput.direction.Text <> "" Then
        ConvertCoordinates
    End If
End Sub

Private Sub txtLonMinutes_Change()
    Static isProcessing As Boolean
    If isProcessing Then Exit Sub
    
    isProcessing = True

    If InStr(Me.txtLonMinutes.Text, ",") > 0 Then
        Dim curPos As Integer
        curPos = Me.txtLonMinutes.selStart
        
        Me.txtLonMinutes.Text = Replace(Me.txtLonMinutes.Text, ",", ".")
        Me.txtLonMinutes.selStart = curPos
    End If

    Me.txtLonMinutes.ForeColor = RGB(0, 0, 0) ' Всегда черный текст

    If Me.txtLonMinutes.Text <> "" And Me.txtLonMinutes.Text <> "." And _
       LongitudeInput.degrees.Text <> "" And LongitudeInput.direction.Text <> "" Then
        ConvertCoordinates
    End If
    
    isProcessing = False
End Sub

Private Sub cboLatDirection_Change()
    If Me.txtLatDegrees.Text <> "" And Me.txtLatMinutes.Text <> "" And _
       Me.txtLatMinutes.Text <> "." Then
        ConvertCoordinates
    End If
End Sub

Private Sub cboLonDirection_Change()
    If Me.txtLonDegrees.Text <> "" And Me.txtLonMinutes.Text <> "" And _
       Me.txtLonMinutes.Text <> "." Then
        ConvertCoordinates
    End If
End Sub

Private Sub ConvertMinutesToDecimal()

    On Error Resume Next
    
    With Me.fraMain.fraCoordinates

        If .txtLatitude.Text <> "" And IsNumeric(Replace(.txtLatitude.Text, ",", ".")) Then
            Dim latValue As Double
            latValue = CDbl(Replace(.txtLatitude.Text, ",", "."))
            ConvertToDegreesMinutes latValue, LatitudeInput.degrees, LatitudeInput.minutes, LatitudeInput.direction, True
        End If

        If .txtLongitude.Text <> "" And IsNumeric(Replace(.txtLongitude.Text, ",", ".")) Then
            Dim lonValue As Double
            lonValue = CDbl(Replace(.txtLongitude.Text, ",", "."))
            ConvertToDegreesMinutes lonValue, LongitudeInput.degrees, LongitudeInput.minutes, LongitudeInput.direction, False
        End If
    End With
End Sub

Private Sub ConvertDegreesToDecimal()
    On Error Resume Next

    If LatitudeInput.degrees.Text <> "" And LatitudeInput.minutes.Text <> "" And _
       LatitudeInput.minutes.Text <> "." And LatitudeInput.direction.Text <> "" Then

        Dim latDeg As Double, latMin As Double
        latDeg = Val(LatitudeInput.degrees.Text)
        latMin = Val(LatitudeInput.minutes.Text)

        Dim latDec As Double
        latDec = latDeg + (latMin / 60)

        If LatitudeInput.direction.Text = "S" Then
            latDec = -latDec
        End If

        Me.txtLatitude.Text = Format(latDec, "0.0000")
    End If

    If LongitudeInput.degrees.Text <> "" And LongitudeInput.minutes.Text <> "" And _
       LongitudeInput.minutes.Text <> "." And LongitudeInput.direction.Text <> "" Then

        Dim lonDeg As Double, lonMin As Double
        lonDeg = Val(LongitudeInput.degrees.Text)
        lonMin = Val(LongitudeInput.minutes.Text)

        Dim lonDec As Double
        lonDec = lonDeg + (lonMin / 60)

        If LongitudeInput.direction.Text = "W" Then
            lonDec = -lonDec
        End If

        Me.txtLongitude.Text = Format(lonDec, "0.0000")
    End If
End Sub
Private Sub InitializeCoordinateControls()

    Me.fraMain.fraCoordinates.txtLatitude.Text = ""
    Me.fraMain.fraCoordinates.txtLongitude.Text = ""
    Me.fraMain.fraCoordinates.txtLatDegrees.Text = ""
    Me.fraMain.fraCoordinates.txtLatMinutes.Text = ""
    Me.fraMain.fraCoordinates.txtLonDegrees.Text = ""
    Me.fraMain.fraCoordinates.txtLonMinutes.Text = ""
End Sub

Private Sub ValidateMinutes(txt As MSForms.TextBox)

    If Len(txt.Text) = 0 Then
        txt.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If

    If txt.Text = "." Then
        txt.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If

    Dim textValue As String
    textValue = Replace(txt.Text, ",", ".")

    If Not IsNumeric(textValue) Then
        txt.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If

    Dim numValue As Double
    numValue = Val(textValue)

    If numValue >= 60 Or numValue < 0 Then
        txt.ForeColor = RGB(255, 0, 0)
    Else
        txt.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub txtLatitude_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8 ' Backspace - всегда разрешен
            Exit Sub
            
        Case 45 ' Минус - только в начале и если еще нет
            If Me.txtLatitude.selStart > 0 Or InStr(Me.txtLatitude.Text, "-") > 0 Then
                KeyAscii = 0
            End If
            Exit Sub
            
        Case 46 ' Точка - только одна
            If InStr(Me.txtLatitude.Text, ".") > 0 Then
                KeyAscii = 0
            End If
            Exit Sub
            
        Case 44 ' Запятая - заменяем на точку
            KeyAscii = 46 ' ASCII код точки
            Exit Sub
            
        Case 48 To 57 ' Цифры - проверяем диапазон
            Dim newText As String

            If Me.txtLatitude.SelLength > 0 Then
                newText = Left(Me.txtLatitude.Text, Me.txtLatitude.selStart) & Chr(KeyAscii) & _
                        Mid(Me.txtLatitude.Text, Me.txtLatitude.selStart + Me.txtLatitude.SelLength + 1)
            Else
                newText = Left(Me.txtLatitude.Text, Me.txtLatitude.selStart) & Chr(KeyAscii) & _
                        Mid(Me.txtLatitude.Text, Me.txtLatitude.selStart + 1)
            End If

            If IsNumeric(newText) Then
                If Abs(CDbl(newText)) > 90 Then
                    KeyAscii = 0
                End If
            End If
            Exit Sub
            
        Case Else ' Другие символы запрещены
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtLongitude_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8 ' Backspace - всегда разрешен
            Exit Sub
            
        Case 45 ' Минус - только в начале и если еще нет
            If Me.txtLongitude.selStart > 0 Or InStr(Me.txtLongitude.Text, "-") > 0 Then
                KeyAscii = 0
            End If
            Exit Sub
            
        Case 46 ' Точка - только одна
            If InStr(Me.txtLongitude.Text, ".") > 0 Then
                KeyAscii = 0
            End If
            Exit Sub
            
        Case 44 ' Запятая - заменяем на точку
            KeyAscii = 46 ' ASCII код точки
            Exit Sub
            
        Case 48 To 57 ' Цифры - проверяем диапазон
            Dim newText As String

            If Me.txtLongitude.SelLength > 0 Then
                newText = Left(Me.txtLongitude.Text, Me.txtLongitude.selStart) & Chr(KeyAscii) & _
                        Mid(Me.txtLongitude.Text, Me.txtLongitude.selStart + Me.txtLongitude.SelLength + 1)
            Else
                newText = Left(Me.txtLongitude.Text, Me.txtLongitude.selStart) & Chr(KeyAscii) & _
                        Mid(Me.txtLongitude.Text, Me.txtLongitude.selStart + 1)
            End If

            If IsNumeric(newText) Then
                If Abs(CDbl(newText)) > 180 Then
                    KeyAscii = 0
                End If
            End If
            Exit Sub
            
        Case Else ' Другие символы запрещены
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtLatDegrees_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8  ' Backspace

            Exit Sub
            
        Case 48 To 57  ' Цифры 0-9

            Dim newText As String
            If Me.txtLatDegrees.SelLength > 0 Then
                newText = Left(Me.txtLatDegrees.Text, Me.txtLatDegrees.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtLatDegrees.Text, Me.txtLatDegrees.selStart + Me.txtLatDegrees.SelLength + 1)
            Else
                newText = Left(Me.txtLatDegrees.Text, Me.txtLatDegrees.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtLatDegrees.Text, Me.txtLatDegrees.selStart + 1)
            End If
            
            If IsNumeric(newText) Then

                If CDbl(newText) > 90 Then
                    KeyAscii = 0
                End If
            End If
            Exit Sub
            
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtLonDegrees_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8  ' Backspace

            Exit Sub
            
        Case 48 To 57  ' Цифры 0-9

            Dim newText As String
            If Me.txtLonDegrees.SelLength > 0 Then
                newText = Left(Me.txtLonDegrees.Text, Me.txtLonDegrees.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtLonDegrees.Text, Me.txtLonDegrees.selStart + Me.txtLonDegrees.SelLength + 1)
            Else
                newText = Left(Me.txtLonDegrees.Text, Me.txtLonDegrees.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtLonDegrees.Text, Me.txtLonDegrees.selStart + 1)
            End If
            
            If IsNumeric(newText) Then

                If CDbl(newText) > 180 Then
                    KeyAscii = 0
                End If
            End If
            Exit Sub
            
        Case Else
            KeyAscii = 0
            Exit Sub
    End Select
End Sub

Private Sub txtLatMinutes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8  ' Backspace

            
        Case 46  ' Точка (.)

            If InStr(Me.txtLatMinutes.Text, ".") > 0 Then KeyAscii = 0
            
        Case 44  ' Запятая (,) - заменяем на точку
            If InStr(Me.txtLatMinutes.Text, ".") > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = 46 ' Превращаем в точку
            End If
            
        Case 48 To 57  ' Цифры

            If InStr(Me.txtLatMinutes.Text, ".") = 0 Then

                Dim newText As String
                
                If Me.txtLatMinutes.SelLength > 0 Then
                    newText = Left(Me.txtLatMinutes.Text, Me.txtLatMinutes.selStart) & Chr(KeyAscii) & _
                             Mid(Me.txtLatMinutes.Text, Me.txtLatMinutes.selStart + Me.txtLatMinutes.SelLength + 1)
                Else
                    newText = Left(Me.txtLatMinutes.Text, Me.txtLatMinutes.selStart) & Chr(KeyAscii) & _
                             Mid(Me.txtLatMinutes.Text, Me.txtLatMinutes.selStart + 1)
                End If

                If IsNumeric(newText) And Val(newText) >= 60 Then
                    KeyAscii = 0
                End If
            Else

                Dim dotPos As Integer
                dotPos = InStr(Me.txtLatMinutes.Text, ".")
                
                If Me.txtLatMinutes.selStart > dotPos And _
                   Len(Me.txtLatMinutes.Text) - dotPos >= 1 And _
                   Me.txtLatMinutes.SelLength = 0 Then
                    KeyAscii = 0
                End If
            End If
            
        Case Else

            KeyAscii = 0
    End Select
End Sub
Private Sub txtLonMinutes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    Select Case KeyAscii
        Case 8  ' Backspace

            
        Case 46  ' Точка (.)

            If InStr(Me.txtLonMinutes.Text, ".") > 0 Then KeyAscii = 0
            
        Case 44  ' Запятая (,) - заменяем на точку
            If InStr(Me.txtLonMinutes.Text, ".") > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = 46 ' Превращаем в точку
            End If
            
        Case 48 To 57  ' Цифры

            If InStr(Me.txtLonMinutes.Text, ".") = 0 Then

                Dim newText As String
                
                If Me.txtLonMinutes.SelLength > 0 Then
                    newText = Left(Me.txtLonMinutes.Text, Me.txtLonMinutes.selStart) & Chr(KeyAscii) & _
                             Mid(Me.txtLonMinutes.Text, Me.txtLonMinutes.selStart + Me.txtLonMinutes.SelLength + 1)
                Else
                    newText = Left(Me.txtLonMinutes.Text, Me.txtLonMinutes.selStart) & Chr(KeyAscii) & _
                             Mid(Me.txtLonMinutes.Text, Me.txtLonMinutes.selStart + 1)
                End If

                If IsNumeric(newText) And Val(newText) >= 60 Then
                    KeyAscii = 0
                End If
            Else

                Dim dotPos As Integer
                dotPos = InStr(Me.txtLonMinutes.Text, ".")
                
                If Me.txtLonMinutes.selStart > dotPos And _
                   Len(Me.txtLonMinutes.Text) - dotPos >= 1 And _
                   Me.txtLonMinutes.SelLength = 0 Then
                    KeyAscii = 0
                End If
            End If
            
        Case Else

            KeyAscii = 0
    End Select
End Sub

Private Function ValidateData() As Boolean

    If Not ValidateRequiredFields Then Exit Function

    If Not ValidateCoordinates Then
        MsgBox "Incorrect coordinate format!" & Chr(13) & "Неверный формат координат!", vbExclamation
        Exit Function
    End If
    
    ValidateData = True
End Function

Private Function ValidateRequiredFields() As Boolean

    If Me.txtDateTime1.value = "" Then
        MsgBox "Fill in date/time field!" & Chr(13) & "Заполните поле даты/времени!", vbExclamation
        Exit Function
    End If

    If mCoordFormat = COORD_FORMAT_DECIMAL Then
        If Me.txtLongitude.value = "" Or Me.txtLatitude.value = "" Then
            MsgBox "Enter coordinates!" & Chr(13) & "Введите координаты!", vbExclamation
            Exit Function
        End If
    Else
        If LatitudeInput.degrees.Text = "" Or LatitudeInput.minutes.Text = "" Or _
           LongitudeInput.degrees.Text = "" Or LongitudeInput.minutes.Text = "" Then
            MsgBox "Enter coordinates!" & Chr(13) & "Введите координаты!", vbExclamation
            Exit Function
        End If
    End If
    
    ValidateRequiredFields = True
End Function
Private Function ValidateCoordinates() As Boolean
    If mCoordFormat = COORD_FORMAT_DECIMAL Then

        If Me.txtLatitude.Text = "" Or Me.txtLongitude.Text = "" Then Exit Function

        Dim lat As Double, lon As Double
        lat = Val(Me.txtLatitude.Text)
        lon = Val(Me.txtLongitude.Text)

        If lat < -90 Or lat > 90 Or lon < -180 Or lon > 180 Then Exit Function
    Else

        If LatitudeInput.degrees.Text = "" Or LatitudeInput.minutes.Text = "" Or _
           LatitudeInput.direction.Text = "" Or LongitudeInput.degrees.Text = "" Or _
           LongitudeInput.minutes.Text = "" Or LongitudeInput.direction.Text = "" Then
            Exit Function
        End If

        Dim latDeg As Double, lonDeg As Double
        latDeg = Val(LatitudeInput.degrees.Text)
        lonDeg = Val(LongitudeInput.degrees.Text)
        
        If latDeg < 0 Or latDeg > 90 Or lonDeg < 0 Or lonDeg > 180 Then Exit Function

        Dim latMin As Double, lonMin As Double
        latMin = Val(LatitudeInput.minutes.Text)
        lonMin = Val(LongitudeInput.minutes.Text)
        
        If latMin < 0 Or latMin >= 60 Or lonMin < 0 Or lonMin >= 60 Then Exit Function
    End If
    
    ValidateCoordinates = True
End Function

Private Sub SetDefaultValues()
    If Me.Tag = "New" Then
        Dim currentTime As Date
        currentTime = Now
        If Minute(currentTime) > 30 Then
            currentTime = DateAdd("h", 1, currentTime)
        End If
        
        Me.txtDateTime1.value = Format(DateSerial(Year(currentTime), Month(currentTime), day(currentTime)) + _
                           Hour(currentTime) / 24, "dd.mm.yyyy hh:00")
    End If
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler

    If Not ValidateData Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    On Error Resume Next
    ws.Unprotect PASSWORD:=PASSWORD
    On Error GoTo 0

    Dim targetRow As Long
    If Me.Tag = "New" Then
        targetRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    Else
        targetRow = CLng(Me.Tag)
    End If

    SaveDataToSheet ws, targetRow

    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0

    MsgBox "Данные успешно сохранены!", vbInformation
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка сохранения данных: " & vbNewLine & Err.Description, vbCritical

    On Error Resume Next
    ws.Protect PASSWORD:=PASSWORD, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub SaveDataToSheet(ByRef ws As Worksheet, ByVal targetRow As Long)
    On Error GoTo ErrorHandler
    
    With ws

        .Cells(targetRow, 1) = CDate(Me.txtDateTime1.value)

        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            Dim latText As String

            latText = Replace(Me.txtLatitude.value, ",", ".")
            If IsNumeric(latText) Then

                latText = Format(Val(latText), "0.0000")

                latText = Replace(latText, ",", ".")

                .Cells(targetRow, 2).value = latText
            End If
        Else
            Dim latVal As Double
            latVal = GetDecimalCoordinates(LatitudeInput.degrees.Text, _
                                          LatitudeInput.minutes.Text, _
                                          LatitudeInput.direction.Text)

            Dim latStr As String
            latStr = Format(latVal, "0.0000")
            latStr = Replace(latStr, ",", ".")
            .Cells(targetRow, 2).value = latStr
        End If

        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            Dim lonText As String

            lonText = Replace(Me.txtLongitude.value, ",", ".")
            If IsNumeric(lonText) Then

                lonText = Format(Val(lonText), "0.0000")

                lonText = Replace(lonText, ",", ".")

                .Cells(targetRow, 3).value = lonText
            End If
        Else
            Dim lonVal As Double
            lonVal = GetDecimalCoordinates(LongitudeInput.degrees.Text, _
                                          LongitudeInput.minutes.Text, _
                                          LongitudeInput.direction.Text)

            Dim lonStr As String
            lonStr = Format(lonVal, "0.0000")
            lonStr = Replace(lonStr, ",", ".")
            .Cells(targetRow, 3).value = lonStr
        End If

        If Me.txtTemp.Text <> "" Then
            If InStr(Me.txtTemp.Text, ",") > 0 Then
                .Cells(targetRow, 4) = Me.txtTemp.Text
            Else
                .Cells(targetRow, 4) = Val(Me.txtTemp.Text)
            End If
        End If

        If Me.txtBarometer.Text <> "" Then
            If InStr(Me.txtBarometer.Text, ",") > 0 Then
                .Cells(targetRow, 5) = Me.txtBarometer.Text
            Else
                .Cells(targetRow, 5) = Val(Me.txtBarometer.Text)
            End If
        End If

        If Me.txtVisibility.Text <> "" Then
            .Cells(targetRow, 6) = Val(Me.txtVisibility.value)
        End If

        If Me.txtWindDirection.Text = "0" And Me.txtWindSpeed.Text = "0" Then
            .Cells(targetRow, 7) = "0"
            .Cells(targetRow, 8) = "0"
        Else
            If Me.txtWindDirection.Text <> "" Then
                .Cells(targetRow, 7) = Val(Me.txtWindDirection.value)
            End If

            If Me.txtWindSpeed.Text <> "" Then
                If InStr(Me.txtWindSpeed.Text, ",") > 0 Then
                    .Cells(targetRow, 8) = Me.txtWindSpeed.Text
                Else
                    .Cells(targetRow, 8) = Val(Me.txtWindSpeed.Text)
                End If
            End If
        End If

        If Me.chkSeaSwell.value Then

            If Me.txtSeaSwellDirection.Text <> "" Then
                .Cells(targetRow, 9) = Val(Me.txtSeaSwellDirection.value)
            End If

            If Me.txtSeaSwell.Text <> "" Then
                If InStr(Me.txtSeaSwell.Text, ",") > 0 Then
                    .Cells(targetRow, 10) = Me.txtSeaSwell.Text
                Else
                    .Cells(targetRow, 10) = Val(Me.txtSeaSwell.Text)
                End If
            End If

            If Me.txtWindWaveDirection.Text <> "" Then
                .Cells(targetRow, 11) = Val(Me.txtWindWaveDirection.value)
            End If

            If Me.txtWindWaveHeight.Text <> "" Then
                If InStr(Me.txtWindWaveHeight.Text, ",") > 0 Then
                    .Cells(targetRow, 12) = Me.txtWindWaveHeight.Text
                Else
                    .Cells(targetRow, 12) = Val(Me.txtWindWaveHeight.Text)
                End If
            End If
        Else
            .Cells(targetRow, 9) = "0"   ' Sea Swell Direction
            .Cells(targetRow, 10) = "0"  ' Sea Swell
            .Cells(targetRow, 11) = "0"  ' Wind wave direction
            .Cells(targetRow, 12) = "0"  ' Wind wave height
        End If

        If Me.chkIceNotated.value Then

            If Me.cboIceScore.ListIndex <> -1 Then
                .Cells(targetRow, 13) = Me.cboIceScore.List(Me.cboIceScore.ListIndex, 1)
            Else
                .Cells(targetRow, 13) = ""
            End If

            If Me.cboIceType.ListIndex <> -1 Then
                .Cells(targetRow, 14) = Me.cboIceType.List(Me.cboIceType.ListIndex, 1)
            Else
                .Cells(targetRow, 14) = ""
            End If

            If Me.cboIceShape.ListIndex <> -1 Then
                .Cells(targetRow, 15) = Me.cboIceShape.List(Me.cboIceShape.ListIndex, 1)
            Else
                .Cells(targetRow, 15) = ""
            End If
        Else

            .Cells(targetRow, 13) = ""  ' Ice score - пусто
            .Cells(targetRow, 14) = ""  ' Ice type - пусто
            .Cells(targetRow, 15) = ""  ' Ice shape - пусто
        End If

        On Error Resume Next
        With .Range(.Cells(targetRow, 1), .Cells(targetRow, 15))
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        .Cells(targetRow, 2).NumberFormat = "@"
        .Cells(targetRow, 3).NumberFormat = "@"

        If InStr(Me.txtTemp.Text, ",") > 0 Then
            .Cells(targetRow, 4).NumberFormat = "0,0"
        End If
        
        If InStr(Me.txtBarometer.Text, ",") > 0 Then
            .Cells(targetRow, 5).NumberFormat = "0,0"
        End If
        
        If InStr(Me.txtWindSpeed.Text, ",") > 0 Then
            .Cells(targetRow, 8).NumberFormat = "0,0"
        End If
        
        If InStr(Me.txtSeaSwell.Text, ",") > 0 Then
            .Cells(targetRow, 10).NumberFormat = "0,0"
        End If
        
        If InStr(Me.txtWindWaveHeight.Text, ",") > 0 Then
            .Cells(targetRow, 12).NumberFormat = "0,0"
        End If
    End With
    
    Exit Sub

ErrorHandler:
    MsgBox "Error while saving data / Ошибка при сохранении данных" & vbNewLine & _
           "Error Description: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Error Source: " & Err.Source, vbCritical
End Sub

Private Sub txtVisibility_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If KeyAscii = 8 Then Exit Sub

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    Dim newText As String
    If Me.txtVisibility.SelLength > 0 Then
        newText = Left(Me.txtVisibility.Text, Me.txtVisibility.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtVisibility.Text, Me.txtVisibility.selStart + Me.txtVisibility.SelLength + 1)
    Else
        newText = Left(Me.txtVisibility.Text, Me.txtVisibility.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtVisibility.Text, Me.txtVisibility.selStart + 1)
    End If

    If IsNumeric(newText) Then
        If CLng(newText) > 50000 Then
            KeyAscii = 0
        End If
    End If
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtWindDirection_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If KeyAscii = 8 Then Exit Sub

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 48 And (Me.txtWindDirection.Text = "" Or Me.txtWindDirection.SelLength = Len(Me.txtWindDirection.Text)) Then
        Me.txtWindDirection.Text = "0"
        Me.txtWindSpeed.Text = "0"
        KeyAscii = 0
        Exit Sub
    End If

    Dim newText As String
    If Me.txtWindDirection.SelLength > 0 Then
        newText = Left(Me.txtWindDirection.Text, Me.txtWindDirection.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindDirection.Text, Me.txtWindDirection.selStart + Me.txtWindDirection.SelLength + 1)
    Else
        newText = Left(Me.txtWindDirection.Text, Me.txtWindDirection.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindDirection.Text, Me.txtWindDirection.selStart + 1)
    End If

    If IsNumeric(newText) Then
        If CLng(newText) > 360 Then
            KeyAscii = 0
        End If
    End If
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtSeaSwellDirection_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If Not Me.txtSeaSwellDirection.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Then
        Exit Sub
    End If

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    Dim newText As String
    If Me.txtSeaSwellDirection.SelLength > 0 Then
        newText = Left(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.selStart + Me.txtSeaSwellDirection.SelLength + 1)
    Else
        newText = Left(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.selStart + 1)
    End If

    If IsNumeric(newText) Then
        If CLng(newText) > 360 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtWindWaveDirection_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If Not Me.txtWindWaveDirection.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Then
        Exit Sub
    End If

    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If

    Dim newText As String
    If Me.txtWindWaveDirection.SelLength > 0 Then
        newText = Left(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.selStart + Me.txtWindWaveDirection.SelLength + 1)
    Else
        newText = Left(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.selStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.selStart + 1)
    End If

    If IsNumeric(newText) Then
        If CLng(newText) > 360 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtTemp_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If KeyAscii = 8 Then Exit Sub

    If KeyAscii = 45 And (Me.txtTemp.Text = "" Or Me.txtTemp.SelLength = Len(Me.txtTemp.Text)) Then
        Exit Sub
    End If

    Select Case KeyAscii
        Case 48 To 57 ' Цифры

            Dim newText As String
            If Me.txtTemp.SelLength > 0 Then
                newText = Left(Me.txtTemp.Text, Me.txtTemp.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtTemp.Text, Me.txtTemp.selStart + Me.txtTemp.SelLength + 1)
            Else
                newText = Left(Me.txtTemp.Text, Me.txtTemp.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtTemp.Text, Me.txtTemp.selStart + 1)
            End If
            
            If IsNumeric(Replace(newText, ",", ".")) Then
                If Abs(CDbl(Replace(newText, ",", "."))) > 100 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            
        Case 44, 46 ' Запятая или точка
            If InStr(Me.txtTemp.Text, ",") > 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            KeyAscii = 44 ' Всегда запятая
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtBarometer_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If KeyAscii = 8 Then Exit Sub

    Select Case KeyAscii
        Case 48 To 57 ' Цифры

            Dim newText As String
            If Me.txtBarometer.SelLength > 0 Then
                newText = Left(Me.txtBarometer.Text, Me.txtBarometer.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtBarometer.Text, Me.txtBarometer.selStart + Me.txtBarometer.SelLength + 1)
            Else
                newText = Left(Me.txtBarometer.Text, Me.txtBarometer.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtBarometer.Text, Me.txtBarometer.selStart + 1)
            End If
            
            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 9000 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            
        Case 44, 46 ' Запятая или точка
            If InStr(Me.txtBarometer.Text, ",") > 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            KeyAscii = 44 ' Всегда запятая
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtSeaSwell_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If Not Me.txtSeaSwell.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Then
        Exit Sub
    End If

    Select Case KeyAscii
        Case 48 To 57  ' Цифры 0-9

            Dim newText As String
            If Me.txtSeaSwell.SelLength > 0 Then
                newText = Left(Me.txtSeaSwell.Text, Me.txtSeaSwell.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtSeaSwell.Text, Me.txtSeaSwell.selStart + Me.txtSeaSwell.SelLength + 1)
            Else
                newText = Left(Me.txtSeaSwell.Text, Me.txtSeaSwell.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtSeaSwell.Text, Me.txtSeaSwell.selStart + 1)
            End If

            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 20 Then
                    KeyAscii = 0
                End If
            End If
            
        Case 44, 46  ' Запятая или точка

            If InStr(Me.txtSeaSwell.Text, ",") > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = 44  ' Всегда запятая
            End If

            If Me.txtSeaSwell.selStart = 0 Then
                KeyAscii = 0
            End If
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtWindWaveHeight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If Not Me.txtWindWaveHeight.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If

    If KeyAscii = 8 Then
        Exit Sub
    End If

    Select Case KeyAscii
        Case 48 To 57  ' Цифры 0-9

            Dim newText As String
            If Me.txtWindWaveHeight.SelLength > 0 Then
                newText = Left(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.selStart + Me.txtWindWaveHeight.SelLength + 1)
            Else
                newText = Left(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.selStart + 1)
            End If

            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 20 Then
                    KeyAscii = 0
                End If
            End If
            
        Case 44, 46  ' Запятая или точка

            If InStr(Me.txtWindWaveHeight.Text, ",") > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = 44  ' Всегда запятая
            End If

            If Me.txtWindWaveHeight.selStart = 0 Then
                KeyAscii = 0
            End If
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub txtWindSpeed_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler

    If KeyAscii = 8 Then Exit Sub

    If KeyAscii = 48 And (Me.txtWindSpeed.Text = "" Or Me.txtWindSpeed.SelLength = Len(Me.txtWindSpeed.Text)) Then
        Me.txtWindSpeed.Text = "0"
        Me.txtWindDirection.Text = "0"
        KeyAscii = 0
        Exit Sub
    End If

    Select Case KeyAscii
        Case 48 To 57 ' Цифры

            Dim newText As String
            If Me.txtWindSpeed.SelLength > 0 Then
                newText = Left(Me.txtWindSpeed.Text, Me.txtWindSpeed.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindSpeed.Text, Me.txtWindSpeed.selStart + Me.txtWindSpeed.SelLength + 1)
            Else
                newText = Left(Me.txtWindSpeed.Text, Me.txtWindSpeed.selStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindSpeed.Text, Me.txtWindSpeed.selStart + 1)
            End If

            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 100 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            
        Case 44, 46 ' Запятая или точка

            If InStr(Me.txtWindSpeed.Text, ",") > 0 Then
                KeyAscii = 0
                Exit Sub
            End If

            If Me.txtWindSpeed.selStart = 0 Then
                KeyAscii = 0
                Exit Sub
            End If

            KeyAscii = 44
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


