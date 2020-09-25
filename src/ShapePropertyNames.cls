VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapePropertyNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@IgnoreModule UseMeaningfulName
'@Folder("SteelShapes.Data")
Option Explicit
'@PredeclaredId

' Property names given in the AISC shapes database. These have
' been included in a predeclared class because of the tan(alpha)
' name. A class will be used to make it a read only property.

Public Property Get ShapeType() As String
    ShapeType = "Type"
End Property

Public Property Get EdiStdNomenclature() As String
    EdiStdNomenclature = "EDI_Std_Nomenclature"
End Property

Public Property Get AiscManualLabel() As String
    AiscManualLabel = "AISC_Manual_Label"
End Property

Public Property Get HasSpecialNote() As String
    HasSpecialNote = "T_F"
End Property

Public Property Get NominalWeight() As String
    NominalWeight = "W"
End Property

Public Property Get Area() As String
    Area = "A"
End Property

Public Property Get OverallDepth() As String
    OverallDepth = "d"
End Property

Public Property Get DetailingDepth() As String
    DetailingDepth = "ddet"
End Property

Public Property Get HssOverallDepth() As String
    HssOverallDepth = "Ht"
End Property

Public Property Get FlatWallHssDepth() As String
    FlatWallHssDepth = "h"
End Property

Public Property Get OutsideDiameter() As String
    OutsideDiameter = "OD"
End Property

Public Property Get FlangeWidth() As String
    FlangeWidth = "bf"
End Property

Public Property Get DetailingFlangeWidth() As String
    DetailingFlangeWidth = "bfdet"
End Property

Public Property Get HssOverallWidth() As String
    HssOverallWidth = "B"
End Property

Public Property Get FlatWallHssWidth() As String
    FlatWallHssWidth = "b"
End Property

Public Property Get InsideDiameter() As String
    InsideDiameter = "ID"
End Property

Public Property Get WebThickness() As String
    WebThickness = "tw"
End Property

Public Property Get DetailingWebThickness() As String
    DetailingWebThickness = "twdet"
End Property

Public Property Get DetailingWebThicknessDiv2() As String
    DetailingWebThicknessDiv2 = "twdet/2"
End Property

Public Property Get FlangeThickness() As String
    FlangeThickness = "tf"
End Property

Public Property Get DetailingFlangeThickness() As String
    DetailingFlangeThickness = "tfdet"
End Property

Public Property Get AngleLegThickness() As String
    AngleLegThickness = "t"
End Property

Public Property Get NominalWallThickness() As String
    NominalWallThickness = "tnom"
End Property

Public Property Get DesignWallThickness() As String
    DesignWallThickness = "tdes"
End Property

Public Property Get DesignFlangeToWebToeOfFillet() As String
    DesignFlangeToWebToeOfFillet = "kdes"
End Property

Public Property Get DetailingFlangeToWebToeOfFIllet() As String
    DetailingFlangeToWebToeOfFIllet = "kdet"
End Property

Public Property Get CenterOfFlangeToToeOfFillet() As String
    CenterOfFlangeToToeOfFillet = "k1"
End Property

Public Property Get ToNeutralXAxis() As String
    ToNeutralXAxis = "x"
End Property

Public Property Get ToNeutralYAxis() As String
    ToNeutralYAxis = "y"
End Property

Public Property Get ToShearCenter() As String
    ToShearCenter = "eo"
End Property

Public Property Get ToPlasticNeutralXAxis() As String
    ToPlasticNeutralXAxis = "xp"
End Property

Public Property Get ToPlasticNeutralYAxis() As String
    ToPlasticNeutralYAxis = "yp"
End Property

Public Property Get SlendernessRatioForFlange() As String
    SlendernessRatioForFlange = "bf/2tf"
End Property

Public Property Get SlendernessRatioForWeb() As String
    SlendernessRatioForWeb = "h/tw"
End Property

Public Property Get SlendernessRatioForAngles() As String
    SlendernessRatioForAngles = "b/t"
End Property

Public Property Get SlendernessRatioForHss1() As String
    SlendernessRatioForHss1 = "b/tdes"
End Property

Public Property Get SlendernessRatioForHss2() As String
    SlendernessRatioForHss2 = "h/tdes"
End Property

Public Property Get SlendernessRatioForPipesAndTees() As String
    SlendernessRatioForPipesAndTees = "D/t"
End Property

Public Property Get XAxisMomentOfInertia() As String
    XAxisMomentOfInertia = "Ix"
End Property

Public Property Get XAxisPlasticSectionModulus() As String
    XAxisPlasticSectionModulus = "Zx"
End Property

Public Property Get XAxisSectionModulus() As String
    XAxisSectionModulus = "Sx"
End Property

Public Property Get XRadiusOfGyration() As String
    XRadiusOfGyration = "rx"
End Property

Public Property Get YAxisMomentOfInertia() As String
    YAxisMomentOfInertia = "Iy"
End Property

Public Property Get YAxisPlasticSectionModulus() As String
    YAxisPlasticSectionModulus = "Zy"
End Property

Public Property Get YAxisSectionModulus() As String
    YAxisSectionModulus = "Sy"
End Property

Public Property Get YRadiusOfGyration() As String
    YRadiusOfGyration = "ry"
End Property

Public Property Get ZAxisMomentOfInertia() As String
    ZAxisMomentOfInertia = "Iz"
End Property

Public Property Get ZRadiusOfGyration() As String
    ZRadiusOfGyration = "rz"
End Property

Public Property Get ZAxisSectionModulus() As String
    ZAxisSectionModulus = "Sz"
End Property

Public Property Get TorsionalMomentOfInertia() As String
    TorsionalMomentOfInertia = "J"
End Property

Public Property Get WarpingConstant() As String
    WarpingConstant = "Cw"
End Property

Public Property Get NormalizedWarpingFunction() As String
    NormalizedWarpingFunction = "Wno"
End Property

Public Property Get WarpingStaticalMomentPoint1() As String
    WarpingStaticalMomentPoint1 = "Sw1"
End Property

Public Property Get WarpingStaticalMomentPoint2() As String
    WarpingStaticalMomentPoint2 = "Sw2"
End Property

Public Property Get WarpingStaticalMomentPoint3() As String
    WarpingStaticalMomentPoint3 = "Sw3"
End Property

Public Property Get StaticalMomentInFlangeAboveWeb() As String
    StaticalMomentInFlangeAboveWeb = "Qf"
End Property

Public Property Get StaticalMomentMidDepthOfCrossSection() As String
    StaticalMomentMidDepthOfCrossSection = "Qw"
End Property

Public Property Get PolarRadiusOfGyration() As String
    PolarRadiusOfGyration = "ro"
End Property

Public Property Get FlexuralConstant() As String
    FlexuralConstant = "H"
End Property

Public Property Get TanAlpha() As String
    Const alphaCharacterCode As Long = 945
    TanAlpha = "tan(" & ChrW$(alphaCharacterCode) & ")"
End Property

Public Property Get WAxisMomentOfInertia() As String
    WAxisMomentOfInertia = "Iw"
End Property

Public Property Get PointAtoCenterOfGravityAlongZAxis() As String
    PointAtoCenterOfGravityAlongZAxis = "zA"
End Property

Public Property Get PointBtoCenterOfGravityAlongZAxis() As String
    PointBtoCenterOfGravityAlongZAxis = "zB"
End Property

Public Property Get PointCtoCenterOfGravityAlongZAxis() As String
    PointCtoCenterOfGravityAlongZAxis = "zC"
End Property

Public Property Get PointAtoCenterOfGravityAlongWAxis() As String
    PointAtoCenterOfGravityAlongWAxis = "wA"
End Property

Public Property Get PointBtoCenterOfGravityAlongWAxis() As String
    PointBtoCenterOfGravityAlongWAxis = "wB"
End Property

Public Property Get PointCtoCenterOfGravityAlongWAxis() As String
    PointCtoCenterOfGravityAlongWAxis = "wC"
End Property

Public Property Get WAxisSectionModulusAboutA() As String
    WAxisSectionModulusAboutA = "SwA"
End Property

Public Property Get WAxisSectionModulusAboutB() As String
    WAxisSectionModulusAboutB = "SwB"
End Property

Public Property Get WAxisSectionModulusAboutC() As String
    WAxisSectionModulusAboutC = "SwC"
End Property

Public Property Get ZAxisSectionModulusAboutA() As String
    ZAxisSectionModulusAboutA = "SzA"
End Property

Public Property Get ZAxisSectionModulusAboutB() As String
    ZAxisSectionModulusAboutB = "SzB"
End Property

Public Property Get ZAxisSectionModulusAboutC() As String
    ZAxisSectionModulusAboutC = "SzC"
End Property

Public Property Get EffectiveRadiusOfGyration() As String
    EffectiveRadiusOfGyration = "rts"
End Property

Public Property Get FlangesCentroidalDistance() As String
    FlangesCentroidalDistance = "ho"
End Property

Public Property Get ShapePerimeterMinusOneFlangeSurface() As String
    ShapePerimeterMinusOneFlangeSurface = "PA"
End Property

Public Property Get SingleAngleShapePermieterMinusLongLeg() As String
    SingleAngleShapePermieterMinusLongLeg = "PA2"
End Property

Public Property Get ShapePerimeter() As String
    ShapePerimeter = "PB"
End Property

Public Property Get BoxPerimeterMinusOneFlangeSurface() As String
    BoxPerimeterMinusOneFlangeSurface = "PC"
End Property

Public Property Get BoxPerimeter() As String
    BoxPerimeter = "PD"
End Property

Public Property Get DistanceBetweenWebToesOfFillet() As String
    DistanceBetweenWebToesOfFillet = "T"
End Property

Public Property Get InnerWorkableGage() As String
    InnerWorkableGage = "WGi"
End Property

Public Property Get OuterWorkableGage() As String
    OuterWorkableGage = "WGo"
End Property
