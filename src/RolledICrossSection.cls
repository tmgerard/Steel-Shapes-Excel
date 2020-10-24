VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledICrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Shapes")
'@IgnoreModule UseMeaningfulName
Option Explicit
Implements IRolledICrossSection

Private properties As Object
Private Const PropertyNotFoundErrorMessage As String = "The requested property does not exists."
Private Const ModuleName As String = "RolledICrossSection"

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

Private Property Get WebThickness() As Double
    WebThickness = properties.item(ShapePropertyNames.WebThickness)
End Property

Private Property Get Zx() As Double
    Zx = properties.item(ShapePropertyNames.XAxisPlasticSectionModulus)
End Property

Private Property Get Zy() As Double
    Zy = properties.item(ShapePropertyNames.YAxisPlasticSectionModulus)
End Property

' IRolledCrossSection Implementation
Private Property Get IRolledICrossSection_Area() As Double
    IRolledICrossSection_Area = Area
End Property

Private Property Get IRolledICrossSection_Cw() As Double
    IRolledICrossSection_Cw = Cw
End Property

Private Property Get IRolledICrossSection_Depth() As Double
    IRolledICrossSection_Depth = Depth
End Property

Private Property Get IRolledICrossSection_FlangeThickness() As Double
    IRolledICrossSection_FlangeThickness = FlangeThickness
End Property

Private Property Get IRolledICrossSection_FlangeWidth() As Double
    IRolledICrossSection_FlangeWidth = FlangeWidth
End Property

Private Function IRolledICrossSection_GetProperty(ByRef propertyName As String) As Variant
    IRolledICrossSection_GetProperty = GetProperty(propertyName)
End Function

Private Property Get IRolledICrossSection_Ix() As Double
    IRolledICrossSection_Ix = Ix
End Property

Private Property Get IRolledICrossSection_Iy() As Double
    IRolledICrossSection_Iy = Iy
End Property

Private Property Get IRolledICrossSection_J() As Double
    IRolledICrossSection_J = J
End Property

Private Property Get IRolledICrossSection_Name() As String
    IRolledICrossSection_Name = Name
End Property

Private Property Get IRolledICrossSection_NominalWeight() As Double
    IRolledICrossSection_NominalWeight = NominalWeight
End Property

Private Property Get IRolledICrossSection_rx() As Double
    IRolledICrossSection_rx = rx
End Property

Private Property Get IRolledICrossSection_ry() As Double
    IRolledICrossSection_ry = ry
End Property

Private Property Get IRolledICrossSection_ShapeType() As String
    IRolledICrossSection_ShapeType = ShapeType
End Property

Private Property Get IRolledICrossSection_Sx() As Double
    IRolledICrossSection_Sx = Sx
End Property

Private Property Get IRolledICrossSection_Sy() As Double
    IRolledICrossSection_Sy = Sy
End Property

Private Property Get IRolledICrossSection_WebThickness() As Double
    IRolledICrossSection_WebThickness = WebThickness
End Property

Private Property Get IRolledICrossSection_Zx() As Double
    IRolledICrossSection_Zx = Zx
End Property

Private Property Get IRolledICrossSection_Zy() As Double
    IRolledICrossSection_Zy = Zy
End Property
