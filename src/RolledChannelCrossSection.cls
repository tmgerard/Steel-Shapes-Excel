VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledChannelCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Shapes")
Option Explicit
Implements IRolledChannelCrossSection

Private properties As Object
Private Const PropertyNotFoundErrorMessage As String = "The requested property does not exists."
Private Const ModuleName As String = "RolledChannelCrossSection"

Public Sub Create(ByVal shapeGetter As IRolledShapeGetter, ByRef shapeName As String)
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

Private Property Get WebThickness() As Double
    WebThickness = properties.item(ShapePropertyNames.WebThickness)
End Property

Private Property Get Zx() As Double
    Zx = properties.item(ShapePropertyNames.XAxisPlasticSectionModulus)
End Property

Private Property Get Zy() As Double
    Zy = properties.item(ShapePropertyNames.YAxisPlasticSectionModulus)
End Property

' IRolledChannelCrossSection Implementation
Private Property Get IRolledChannelCrossSection_Area() As Double
    IRolledChannelCrossSection_Area = Area
End Property

Private Property Get IRolledChannelCrossSection_Cw() As Double
    IRolledChannelCrossSection_Cw = Cw
End Property

Private Property Get IRolledChannelCrossSection_Depth() As Double
    IRolledChannelCrossSection_Depth = Depth
End Property

Private Property Get IRolledChannelCrossSection_FlangeThickness() As Double
    IRolledChannelCrossSection_FlangeThickness = FlangeThickness
End Property

Private Property Get IRolledChannelCrossSection_FlangeWidth() As Double
    IRolledChannelCrossSection_FlangeWidth = FlangeWidth
End Property

Private Function IRolledChannelCrossSection_GetProperty(propertyName As String) As Variant
    IRolledChannelCrossSection_GetProperty = GetProperty(propertyName)
End Function

Private Property Get IRolledChannelCrossSection_Ix() As Double
    IRolledChannelCrossSection_Ix = Ix
End Property

Private Property Get IRolledChannelCrossSection_Iy() As Double
    IRolledChannelCrossSection_Iy = Iy
End Property

Private Property Get IRolledChannelCrossSection_J() As Double
    IRolledChannelCrossSection_J = Name
End Property

Private Property Get IRolledChannelCrossSection_Name() As String
    IRolledChannelCrossSection_Name = Name
End Property

Private Property Get IRolledChannelCrossSection_NominalWeight() As Double
    IRolledChannelCrossSection_NominalWeight = NominalWeight
End Property

Private Property Get IRolledChannelCrossSection_rx() As Double
    IRolledChannelCrossSection_rx = rx
End Property

Private Property Get IRolledChannelCrossSection_ry() As Double
    IRolledChannelCrossSection_ry = ry
End Property

Private Property Get IRolledChannelCrossSection_ShapeType() As String
    IRolledChannelCrossSection_ShapeType = ShapeType
End Property

Private Property Get IRolledChannelCrossSection_Sx() As Double
    IRolledChannelCrossSection_Sx = Sx
End Property

Private Property Get IRolledChannelCrossSection_Sy() As Double
    IRolledChannelCrossSection_Sy = Sy
End Property

Private Property Get IRolledChannelCrossSection_WebThickness() As Double
    IRolledChannelCrossSection_WebThickness = WebThickness
End Property

Private Property Get IRolledChannelCrossSection_Zx() As Double
    IRolledChannelCrossSection_Zx = Zx
End Property

Private Property Get IRolledChannelCrossSection_Zy() As Double
    IRolledChannelCrossSection_Zy = Zy
End Property
