Public MustInherit Class Hole

    Private f_holeDia As Double
    Private f_holeDepth As Double
    Private f_holeXOffset As Double
    Private f_holeYOffset As Double
    Private f_IsCounterbored As Boolean
    Private f_isCounterSunk As Boolean
    Private f_isThruHole As Boolean
    Private f_isSpotFace As Boolean
    Private f_counterBoreDia As Integer
    Private f_counterBoreDepth As Integer
    Private f_spotFaceDia As Integer
    Private f_spotFaceDepth As Integer
    Private f_counterSinkAngle As Integer
    Private f_countersinkDia As Integer
    Private f_holeIsSuppressed As Integer

    Public Property HoleDia As Double
        Get
            Return f_holeDia
        End Get
        Set(value As Double)
            f_holeDia = value
        End Set
    End Property

    Public Property HoleDepth As Double
        Get
            Return f_holeDepth
        End Get
        Set(value As Double)
            f_holeDepth = value
        End Set
    End Property

    Public Property HoleXOffset As Double
        Get
            Return f_holeXOffset
        End Get
        Set(value As Double)
            f_holeXOffset = value
        End Set
    End Property

    Public Property HoleYOffset As Double
        Get
            Return f_holeYOffset
        End Get
        Set(value As Double)
            f_holeYOffset = value
        End Set
    End Property

    Public Property CounterBoreDepth As Double
        Get
            Return f_counterBoreDepth
        End Get
        Set(value As Double)
            f_counterBoreDepth = value
        End Set
    End Property

    Public Property CounterBoreDia As Double
        Get
            Return f_counterBoreDia
        End Get
        Set(value As Double)
            f_counterBoreDia = value
        End Set
    End Property

    Public Property CounterSinkDia As Double
        Get
            Return f_countersinkDia
        End Get
        Set(value As Double)
            f_countersinkDia = value
        End Set
    End Property

    Public Property CounterSinkAngle As Double
        Get
            Return f_counterSinkAngle
        End Get
        Set(value As Double)
            f_counterSinkAngle = value
        End Set
    End Property

    Public Property IsCounterBored As Boolean
        Get
            Return f_IsCounterbored
        End Get
        Set(value As Boolean)
            f_IsCounterbored = value
        End Set
    End Property

    Public Property IsCounterSunk As Boolean
        Get
            Return f_isCounterSunk
        End Get
        Set(value As Boolean)
            f_isCounterSunk = value
        End Set
    End Property

    Public Property IsThruHole As Boolean
        Get
            Return f_isThruHole
        End Get
        Set(value As Boolean)
            f_isThruHole = value
        End Set
    End Property

    Public Property IsSpotFace As Boolean
        Get
            Return f_isSpotFace
        End Get
        Set(value As Boolean)
            f_isSpotFace = value
        End Set
    End Property

    Public Property HoleIsSuppressed As Integer
        Get
            Return f_holeIsSuppressed
        End Get
        Set(value As Integer)
            f_holeIsSuppressed = value
        End Set
    End Property

    Public Overridable Sub Hole(Dia As Double, Depth As Double)
        HoleDia = Dia
        HoleDepth = Depth
    End Sub

    ''' <summary>
    ''' used if no hole properties are specified.
    ''' </summary>
    Public Overridable Sub Hole()
        f_holeDia = 13.37
        f_holeDepth = 13.37
        f_holeXOffset = 133.7
        f_holeYOffset = 133.7
    End Sub
End Class

Public Class ConnectingHoleInterModule
    Inherits Hole
    Public Overrides Sub Hole()
        HoleDia = Convert.ToDouble(13.37 * 1.5)
    End Sub
End Class

Public Class ConnectingHoleTransition
    Inherits Hole
    Public Overrides Sub Hole()
        HoleDia = Convert.ToDouble(13.37 * 1.25)
    End Sub
End Class

Public Class ConnectingHoleVertical
    Inherits Hole
    Public Overrides Sub Hole()
        HoleDia = Convert.ToDouble(13.37 * 2)
    End Sub

End Class

Public Class MountingHoleElement
    Inherits Hole
    Public Overrides Sub Hole()
        HoleDia = Convert.ToDouble(13.37 * 0.9)
    End Sub
End Class

Public Class ConnectingHoleHorizontal
    Inherits Hole
    Public Overrides Sub Hole()
        HoleDia = Convert.ToDouble(13.37 * 1.2)
    End Sub
End Class
