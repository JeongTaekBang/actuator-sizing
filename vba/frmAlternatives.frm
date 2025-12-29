VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAlternatives 
   Caption         =   "Select Alternative Model"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "frmAlternatives.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAlternatives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' ============================================
' frmAlternatives - Alternative Model Selection UserForm
' Noah Actuator Sizing Tool
' ============================================
'
' HOW TO CREATE THIS FORM IN VBA EDITOR:
' ======================================
' 1. Open VBA Editor (Alt+F11)
' 2. Insert > UserForm (creates new form)
' 3. Rename to "frmAlternatives" in Properties window
' 4. Set form properties:
'    - Caption: "Select Alternative Model"
'    - Width: 680
'    - Height: 440
'    - StartUpPosition: 1 - CenterOwner
'
' 5. Add controls from Toolbox:
'
'    [lblHeader] - Label for column headers
'    - Name: lblHeader
'    - Left: 12, Top: 8, Width: 640, Height: 16
'    - Font: Consolas, 9pt, Bold
'    - BackStyle: 0 - fmBackStyleTransparent
'
'    [lstAlternatives] - ListBox for alternatives
'    - Name: lstAlternatives
'    - Left: 12, Top: 28, Width: 640, Height: 310
'    - ColumnCount: 9
'    - ColumnWidths: 28;85;70;48;42;60;58;55;55
'    - Font: Consolas, 9pt
'
'    [lblInfo] - Label for info text
'    - Name: lblInfo
'    - Left: 12, Top: 345, Width: 440, Height: 16
'    - Font: 9pt
'    - ForeColor: &H606060 (gray)
'    - BackStyle: 0 - fmBackStyleTransparent
'
'    [btnOK] - OK button
'    - Name: btnOK
'    - Caption: "OK"
'    - Left: 480, Top: 370, Width: 80, Height: 28
'    - Default: True
'
'    [btnCancel] - Cancel button
'    - Name: btnCancel
'    - Caption: "Cancel"
'    - Left: 572, Top: 370, Width: 80, Height: 28
'    - Cancel: True
'
' 6. Double-click the form to open code window
' 7. Paste all the code below (from Option Explicit onwards)
' ============================================

Option Explicit

' Public properties for result
Public SelectedAlternative As String
Public UserCancelled As Boolean

' Private variables
Private mAlternatives As Collection
Private mSelectedRow As Long
Private mTargetWorksheet As Worksheet

' ============================================
' Form Initialize
' ============================================

Private Sub UserForm_Initialize()
    UserCancelled = True
    SelectedAlternative = ""
    
    ' Set up ListBox for multi-column display
    With lstAlternatives
        .Clear
        .ColumnCount = 9
        .ColumnWidths = "28;85;70;48;42;60;58;55;55"
        .Font.Name = "Consolas"
        .Font.Size = 9
    End With
    
    ' Set header label
    lblHeader.Caption = "No   Model            Gearbox        Ratio   RPM    Torque     Thrust     Time      Price"
    lblHeader.Font.Name = "Consolas"
    lblHeader.Font.Size = 9
    lblHeader.Font.Bold = True
End Sub

' ============================================
' Public Method - Load Alternatives
' ============================================

Public Sub LoadAlternatives(alternatives As Collection, selectedRow As Long, ws As Worksheet)
    Dim altInfo As Variant
    Dim parts() As String
    Dim gbDisplay As String
    Dim ratioDisplay As String
    Dim thrustVal As Double
    Dim i As Long
    
    Set mAlternatives = alternatives
    mSelectedRow = selectedRow
    Set mTargetWorksheet = ws
    
    ' Update info label
    lblInfo.Caption = "Found " & alternatives.Count & " alternative(s) for Line " & (selectedRow - 3) & _
        "  |  Double-click or select and click OK"
    
    ' Populate list
    With lstAlternatives
        .Clear
        
        i = 1
        For Each altInfo In alternatives
            parts = Split(CStr(altInfo), "|")
            ' Format: Model|Gearbox|Torque|OpTime|Price|Flange|RPM|Ratio|Thrust
            
            If UBound(parts) >= 7 Then
                gbDisplay = parts(1)
                If gbDisplay = "" Then gbDisplay = "-"
                
                ratioDisplay = parts(7)
                If ratioDisplay = "" Or ratioDisplay = "1" Then
                    ratioDisplay = "-"
                Else
                    ratioDisplay = ratioDisplay & ":1"
                End If
                
                ' Get thrust value
                If UBound(parts) >= 8 Then
                    thrustVal = Val(parts(8))
                Else
                    thrustVal = 0
                End If
                
                ' Add row to ListBox
                .AddItem Format(i, "00")
                .List(.ListCount - 1, 1) = parts(0)            ' Model
                .List(.ListCount - 1, 2) = gbDisplay           ' Gearbox
                .List(.ListCount - 1, 3) = ratioDisplay        ' Ratio
                .List(.ListCount - 1, 4) = parts(6)            ' RPM
                .List(.ListCount - 1, 5) = parts(2)            ' Torque(Nm)
                .List(.ListCount - 1, 6) = IIf(thrustVal > 0, CStr(Round(thrustVal, 1)), "-")
                .List(.ListCount - 1, 7) = parts(3)            ' OpTime(s)
                .List(.ListCount - 1, 8) = parts(4)            ' Price($)
                
                i = i + 1
            End If
        Next altInfo
        
        ' Select first item by default
        If .ListCount > 0 Then
            .ListIndex = 0
        End If
    End With
End Sub

' ============================================
' Event Handlers
' ============================================

Private Sub lstAlternatives_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Double-click to select and close
    If lstAlternatives.ListIndex >= 0 Then
        SelectAndClose
    End If
End Sub

Private Sub btnOK_Click()
    SelectAndClose
End Sub

Private Sub btnCancel_Click()
    CancelAndClose
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Handle X button click (same as Cancel)
    If CloseMode = vbFormControlMenu Then
        CancelAndClose
    End If
End Sub

' ============================================
' Helper Methods
' ============================================

Private Sub SelectAndClose()
    Dim selectedIdx As Long
    
    If lstAlternatives.ListIndex < 0 Then
        MsgBox "Please select a model from the list.", vbExclamation, "Noah Sizing Tool"
        Exit Sub
    End If
    
    selectedIdx = lstAlternatives.ListIndex + 1
    SelectedAlternative = CStr(mAlternatives(selectedIdx))
    UserCancelled = False
    
    Me.Hide
End Sub

Private Sub CancelAndClose()
    UserCancelled = True
    SelectedAlternative = ""
    Me.Hide
End Sub
