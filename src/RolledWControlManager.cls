VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledWControlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Controls")
Option Explicit
Implements IRolledShapeControl

Private Type TRolledWControlManager
    filePath As String
    fileNumber As Long
    properties() As String
    text As String
End Type

Private This As TRolledWControlManager

Private Const RolledShapeType As String = "W"
Private Const Delimiter As String = ","
Private Const TypeIndex As Long = 0
Private Const NameIndex As Long = 1

Private Sub Class_Initialize()
    With This
        .filePath = ThisWorkbook.Path & "\" & FileNames.DataFolder & "\" & FileNames.AiscCSVFileName
    End With
End Sub

Public Sub PopulateComboBox(ByRef control As ComboBox)
    On Error GoTo ErrorHandler
    
    control.Clear
    
    This.fileNumber = FreeFile
    Open This.filePath For Input As This.fileNumber
    
    Do While Not EOF(This.fileNumber)
        Line Input #This.fileNumber, This.text
        This.properties = Split(This.text, Delimiter)
        If This.properties(TypeIndex) = RolledShapeType Then
            control.AddItem This.properties(NameIndex)
        End If
    Loop
    
SubExit:
    Close #This.fileNumber
    Exit Sub
    
ErrorHandler:
    ' TODO: add errors
    Resume SubExit
End Sub
Private Sub NewMethod(ByRef control As ComboBox)
    Do While Not EOF(This.fileNumber)
        Line Input #This.fileNumber, This.text
        This.properties = Split(This.text, Delimiter)
        If This.properties(TypeIndex) = RolledShapeType Then
            control.AddItem This.properties(NameIndex)
        End If
    Loop
End Sub

Public Sub PopulateListBox(ByRef control As MSForms.ListBox)
    On Error GoTo ErrorHandler
    
    control.Clear
    
    This.fileNumber = FreeFile
    Open This.filePath For Input As This.fileNumber
    
    Do While Not EOF(This.fileNumber)
        Line Input #This.fileNumber, This.text
        This.properties = Split(This.text, Delimiter)
        If This.properties(TypeIndex) = RolledShapeType Then
            control.AddItem This.properties(NameIndex)
        End If
    Loop
    
SubExit:
    Close #This.fileNumber
    Exit Sub
    
ErrorHandler:
    ' TODO: add errors
    Resume SubExit
End Sub


' IRolledShapeControl Implementation
Private Sub IRolledShapeControl_PopulateComboBox(control As ComboBox)
    PopulateComboBox control
End Sub

Private Sub IRolledShapeControl_PopulateListBox(control As MSForms.ListBox)
    PopulateListBox control
End Sub
