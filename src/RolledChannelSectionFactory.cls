VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledChannelSectionFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("SteelShapes.Shapes")
Option Explicit
'@PredeclaredId

'@Description "Create RolledICrossSection object using data from the specified source file."
Public Function Create(ByVal shapeGetter As IRolledShapeGetter, ByRef shapeName As String) As IRolledChannelCrossSection
    Dim newShape As RolledChannelCrossSection
    Set newShape = New RolledChannelCrossSection
    newShape.Create shapeGetter, shapeName
    Set Create = newShape
End Function
