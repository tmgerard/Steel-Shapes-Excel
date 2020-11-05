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
Private Const ModuleName As String = "CSVTensileMaterialGetter"
Private Const DataNotFoundErrorMessage As String = _
    "The requested data was not found within the CSV file."

Private Sub Class_Initialize()
    With This
        .filePath = ThisWorkbook.Path & "\" & FileNames.DataFolder & "\" & FileNames.MaterialCSVFileName
    End With
End Sub

Public Function GetMaterial(ByVal specificationName As String, ByVal materialGrade As String) As String()
    On Error GoTo ErrorHandler
    
    This.materialFound = False
    
    This.fileNumber = FreeFile
    Open This.filePath For Input As This.fileNumber
    
    PopulateMaterialArray specificationName, materialGrade
    
    Close #This.fileNumber
    
    If Not This.materialFound Then
        DataErrors.ThrowDataNotFoundError DataNotFoundErrorMessage, ModuleName
    End If
    
    GetMaterial = This.materialProperties
    
    Exit Function
    
ErrorHandler:
    Close #This.fileNumber
    ' Re-throw error so it is caught by the function caller
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub PopulateMaterialArray(ByVal specification As String, ByVal grade As String)
    Do While Not EOF(This.fileNumber)
        Line Input #This.fileNumber, This.text
        This.materialProperties = Split(This.text, Delimiter)
        If This.materialProperties(MaterialArrayIndex.specification) = specification And _
            This.materialProperties(MaterialArrayIndex.grade) = grade Then
            This.materialFound = True
            Exit Do
        End If
    Loop
End Sub

' ITensileMaterialGetter Implementation
Private Function ITensileMaterialGetter_GetMaterial(ByVal specificationName As String, ByVal materialGrade As String) As String()
    ITensileMaterialGetter_GetMaterial = Me.GetMaterial(specificationName, materialGrade)
End Function


