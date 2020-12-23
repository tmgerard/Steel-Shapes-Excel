VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledTCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Shapes")
Option Explicit
Implements IRolledTCrossSection

Private properties As Object
Private Const PropertyNotFoundErrorMessage As String = "The requested property does not exists."
Private Const ModuleName As String = "RolledTCrossSection"

Public Sub Create(ByVal shapeGetter As IRolledShapeGetter, ByVal shapeName As String)
    Dim shapeProperties As Object
    Set shapeProperties = shapeGetter.GetRolledShape(shapeName)
    If Not TypeOf shapeProperties Is Dictionary Then
        ' TODO: throw an error
    End If
    Set properties = shapeProperties
End Sub

Private Property Get Area() As Double
    Area = properties.item(ShapePropertyNames.Area)
End Property

Private Property Get Cw() As Double
    Cw = properties.item(ShapePropertyNames.WarpingConstant)
End Property

Private Property Get Depth() As Double
    Depth = properties.item(ShapePropertyNames.OverallDepth)
End Property

Private Property Get FlangeThickness() As Double
    FlangeThickness = properties.item(ShapePropertyNames.FlangeThickness)
End Property

Private Property Get FlangeWidth() As Double
    FlangeWidth = properties.item(ShapePropertyNames.FlangeWidth)
End Property

Private Function GetProperty(ByRef propertyName As String) As Variant
    If Not properties.Exists(propertyName) Then
        DataErrors.ThrowPropertyNotFoundError PropertyNotFoundErrorMessage, ModuleName
    End If
    GetProperty = properties.item(propertyName)
End Function

Private Property Get Ix() As Double
    Ix = properties.item(ShapePropertyNames.XAxisMomentOfInertia)
End Property

Private Property Get Iy() As Double
    Iy = properties.item(ShapePropertyNames.YAxisMomentOfInertia)
End Property

Private Property Get J() As Double
    J = properties.item(ShapePropertyNames.TorsionalMomentOfInertia)
End Property

Private Property Get Name() As String
    Name = properties.item(ShapePropertyNames.AiscManualLabel)
End Property

Private Property Get NominalWeight() As Double
    NominalWeight = properties.item(ShapePropertyNames.NominalWeight)
End Property

Private Property Get rx() As Double
    rx = properties.item(ShapePropertyNames.XRadiusOfGyration)
End Property

Private Property Get ry() As Double
    ry = properties.item(ShapePropertyNames.YRadiusOfGyration)
End Property

Private Property Get ShapeType() As String
    ShapeType = properties.item(ShapePropertyNames.ShapeType)
End Property

Private Property Get Sx() As Double
    Sx = properties.item(ShapePropertyNames.XAxisSectionModulus)
End Property

Private Property Get Sy() As Double
    Sy = properties.item(ShapePropertyNames.YAxisSectionModulus)
End Property

Private Property Get webThickness() As Double
    webThickness = properties.item(ShapePropertyNames.webThickness)
End Property

Private Property Get Zx() As Double
    Zx = properties.item(ShapePropertyNames.XAxisPlasticSectionModulus)
End Property

Private Property Get Zy() As Double
    Zy = properties.item(ShapePropertyNames.YAxisPlasticSectionModulus)
End Property

' IRolledTCrossSection Implementation

Private Property Get IRolledTCrossSection_Area() As Double
    IRolledTCrossSection_Area = Area
End Property

Private Property Get IRolledTCrossSection_Cw() As Double
    IRolledTCrossSection_Cw = Cw
End Property

Private Property Get IRolledTCrossSection_Depth() As Double
    IRolledTCrossSection_Depth = Depth
End Property

Private Property Get IRolledTCrossSection_FlangeThickness() As Double
    IRolledTCrossSection_FlangeThickness = FlangeThickness
End Property

Private Property Get IRolledTCrossSection_FlangeWidth() As Double
    IRolledTCrossSection_FlangeWidth = FlangeWidth
End Property

Private Function IRolledTCrossSection_GetProperty(ByRef propertyName As String) As Variant
    IRolledTCrossSection_GetProperty = GetProperty(propertyName)
End Function

Private Property Get IRolledTCrossSection_Ix() As Double
    IRolledTCrossSection_Ix = Ix
End Property

Private Property Get IRolledTCrossSection_Iy() As Double
    IRolledTCrossSection_Iy = Iy
End Property

Private Property Get IRolledTCrossSection_J() As Double
    IRolledTCrossSection_J = J
End Property

Private Property Get IRolledTCrossSection_Name() As String
    IRolledTCrossSection_Name = Name
End Property

Private Property Get IRolledTCrossSection_NominalWeight() As Double
    IRolledTCrossSection_NominalWeight = NominalWeight
End Property

Private Property Get IRolledTCrossSection_rx() As Double
    IRolledTCrossSection_rx = rx
End Property

Private Property Get IRolledTCrossSection_ry() As Double
    IRolledTCrossSection_ry = ry
End Property

Private Property Get IRolledTCrossSection_Sx() As Double
    IRolledTCrossSection_Sx = Sx
End Property

Private Property Get IRolledTCrossSection_Sy() As Double
    IRolledTCrossSection_Sy = Sy
End Property

Private Property Get IRolledTCrossSection_ShapeType() As String
    IRolledTCrossSection_ShapeType = ShapeType
End Property

Private Property Get IRolledTCrossSection_WebThickness() As Double
    IRolledTCrossSection_WebThickness = webThickness
End Property

Private Property Get IRolledTCrossSection_Zx() As Double
    IRolledTCrossSection_Zx = Zx
End Property

Private Property Get IRolledTCrossSection_Zy() As Double
    IRolledTCrossSection_Zy = Zy
End Property

