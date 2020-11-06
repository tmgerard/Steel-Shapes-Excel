VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RolledLCrossSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SteelShapes.Shapes")
'@IgnoreModule UseMeaningfulName
Option Explicit
Implements IRolledLCrossSection

Private properties As Object
Private Const PropertyNotFoundErrorMessage As String = "The requested property does not exists."
Private Const ModuleName As String = "Rolled>CrossSection"

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

Private Property Get Iz() As Double
    Iz = properties.item(ShapePropertyNames.ZAxisMomentOfInertia)
End Property

Private Property Get J() As Double
    J = properties.item(ShapePropertyNames.TorsionalMomentOfInertia)
End Property

Private Property Get LengthLongLeg() As Double
    LengthLongLeg = properties.item(ShapePropertyNames.LengthLongAngleLeg)
End Property

Private Property Get LengthShortLeg() As Double
    LengthShortLeg = properties.item(ShapePropertyNames.LengthShortAngleLeg)
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

Private Property Get rz() As Double
    rz = properties.item(ShapePropertyNames.ZRadiusOfGyration)
End Property

Private Property Get Sx() As Double
    Sx = properties.item(ShapePropertyNames.XAxisSectionModulus)
End Property

Private Property Get Sy() As Double
    Sy = properties.item(ShapePropertyNames.YAxisSectionModulus)
End Property

Private Property Get Sz() As Double
    Sz = properties.item(ShapePropertyNames.ZAxisSectionModulus)
End Property

Private Property Get ShapeType() As String
    ShapeType = properties.item(ShapePropertyNames.ShapeType)
End Property

Private Property Get Thickness() As Double
    Thickness = properties.item(ShapePropertyNames.AngleLegThickness)
End Property

Private Property Get Zx() As Double
    Zx = properties.item(ShapePropertyNames.XAxisPlasticSectionModulus)
End Property

Private Property Get Zy() As Double
    Zy = properties.item(ShapePropertyNames.YAxisPlasticSectionModulus)
End Property

' IRolledLCrossSection Implementation
Private Property Get IRolledLCrossSection_Area() As Double
    IRolledLCrossSection_Area = Area
End Property

Private Property Get IRolledLCrossSection_Cw() As Double
    IRolledLCrossSection_Cw = Cw
End Property

Private Function IRolledLCrossSection_GetProperty(ByRef propertyName As String) As Variant
    IRolledLCrossSection_GetProperty = GetProperty(propertyName)
End Function

Private Property Get IRolledLCrossSection_Ix() As Double
    IRolledLCrossSection_Ix = Ix
End Property

Private Property Get IRolledLCrossSection_Iy() As Double
    IRolledLCrossSection_Iy = Iy
End Property

Private Property Get IRolledLCrossSection_Iz() As Double
    IRolledLCrossSection_Iz = Iz
End Property

Private Property Get IRolledLCrossSection_J() As Double
    IRolledLCrossSection_J = J
End Property

Private Property Get IRolledLCrossSection_LengthLongLeg() As Double
    IRolledLCrossSection_LengthLongLeg = LengthLongLeg
End Property

Private Property Get IRolledLCrossSection_LengthShortLeg() As Double
    IRolledLCrossSection_LengthShortLeg = LengthShortLeg
End Property

Private Property Get IRolledLCrossSection_Name() As String
    IRolledLCrossSection_Name = Name
End Property

Private Property Get IRolledLCrossSection_NominalWeight() As Double
    IRolledLCrossSection_NominalWeight = NominalWeight
End Property

Private Property Get IRolledLCrossSection_rx() As Double
    IRolledLCrossSection_rx = rx
End Property

Private Property Get IRolledLCrossSection_ry() As Double
    IRolledLCrossSection_ry = ry
End Property

Private Property Get IRolledLCrossSection_rz() As Double
    IRolledLCrossSection_rz = rz
End Property

Private Property Get IRolledLCrossSection_Sx() As Double
    IRolledLCrossSection_Sx = Sx
End Property

Private Property Get IRolledLCrossSection_Sy() As Double
    IRolledLCrossSection_Sy = Sy
End Property

Private Property Get IRolledLCrossSection_Sz() As Double
    IRolledLCrossSection_Sz = Sz
End Property

Private Property Get IRolledLCrossSection_ShapeType() As String
    IRolledLCrossSection_ShapeType = ShapeType
End Property

Private Property Get IRolledLCrossSection_Thickness() As Double
    IRolledLCrossSection_Thickness = Thickness
End Property

Private Property Get IRolledLCrossSection_Zx() As Double
    IRolledLCrossSection_Zx = Zx
End Property

Private Property Get IRolledLCrossSection_Zy() As Double
    IRolledLCrossSection_Zy = Zy
End Property

