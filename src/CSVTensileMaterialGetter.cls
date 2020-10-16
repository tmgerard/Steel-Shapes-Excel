VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVTensileMaterialGetter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Data")
Option Explicit
Implements ITensileMaterialGetter

Private Type TCSVTensileMaterialGetter
    filePath As String
    fileNumber As String
    text As String
    materialProperties() As String
    materialFound As Boolean
End Type

Private This As TCSVTensileMaterialGetter

Private Const Delimiter As String = ","

Public Enum MatArrayIndex
    matSpecification
    matGrade
    matFy
    matFu
End Enum

Private Sub Class_Initialize()
    With This
        .filePath = ThisWorkbook.Path & "\" & FileNames.DataFolder & "\" & FileNames.MaterialCSVFileName
    End With
End Sub

Public Function GetMaterial(ByVal specification As String, ByVal grade As String) As String()
    On Error GoTo ErrorHandler
    
    This.materialFound = False
    
    This.fileNumber = FreeFile
    Open This.filePath For Input As This.fileNumber
    
    PopulateMaterialArray specification, grade
    
    Close #This.fileNumber
    
    If Not This.materialFound Then
        ' TODO: throw error
    End If
    
    GetMaterial = This.materialProperties
    
    Exit Function
    
ErrorHandler:
    Close #This.fileNumber
    MsgBox Err.Description
    ' TODO: Throw CSVTensileShapeGetterError
End Function

Private Sub PopulateMaterialArray(ByRef specification As String, ByRef grade As String)
    Do While Not EOF(This.fileNumber)
        Line Input #This.fileNumber, This.text
        This.materialProperties = Split(This.text, Delimiter)
        If This.materialProperties(matSpecification) = specification And _
            This.materialProperties(matGrade) = grade Then
            This.materialFound = True
            Exit Do
        End If
    Loop
End Sub

' ITensileMaterialGetter Implementation
Private Function ITensileMaterialGetter_GetMaterial(ByVal specification As String, ByVal grade As String) As String()
    ITensileMaterialGetter_GetMaterial = Me.GetMaterial(specification, grade)
End Function

