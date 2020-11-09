VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledTCrossSectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("SteelShapes.Shapes")
Option Explicit
'@PredeclaredId

'@Description "Create RolledTCrossSection object using data from the specified source file."
Public Function Create(ByVal shapeGetter As IRolledShapeGetter, ByVal shapeName As String) As IRolledTCrossSection
Attribute Create.VB_Description = "Create RolledTCrossSection object using data from the specified source file."
    Dim newShape As RolledTCrossSection
    Set newShape = New RolledTCrossSection
    newShape.Create shapeGetter, shapeName
    Set Create = newShape
End Function
