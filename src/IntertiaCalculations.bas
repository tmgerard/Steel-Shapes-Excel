Attribute VB_Name = "IntertiaCalculations"
'@Folder("SteelShapes.General")
Option Explicit

Public Function ParallelAxisTheorem(ByVal Inertia As Double, _
                                    ByVal Area As Double, _
                                    ByVal DistanceToAxis As Double) As Double
                                    
    ParallelAxisTheorem = Inertia + Area * DistanceToAxis ^ 2
                                    
End Function

