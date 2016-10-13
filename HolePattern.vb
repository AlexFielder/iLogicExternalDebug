Public MustInherit Class HolePattern
    Private f_numHolesDirection1 As Integer
    Private f_NumHolesDirection2 As Integer
    Private f_HoleSpacingDirection1 As Double
    Private f_HoleSpacingDirection2 As Double

    Public Property HoleSpacingDirection1 As Double
        Get
            Return f_numHolesDirection1
        End Get
        Set(value As Double)
            f_numHolesDirection1 = value
        End Set
    End Property

    Public Property HoleSpacingDirection2 As Double
        Get
            Return f_HoleSpacingDirection2
        End Get
        Set(value As Double)
        End Set
    End Property

    Public Property NumHolesDirection1 As Integer
        Get
            Return f_numHolesDirection1
        End Get
        Set(value As Integer)
            f_numHolesDirection1 = value
        End Set
    End Property

    Public Property NumHolesDirection2 As Integer
        Get
            Return f_NumHolesDirection2
        End Get
        Set(value As Integer)
            f_NumHolesDirection2 = value
        End Set
    End Property

    Public Overridable Sub HolePattern()
        HoleSpacingDirection1 = 13.37
        HoleSpacingDirection2 = 13.37
        NumHolesDirection1 = 13
        NumHolesDirection2 = 13

    End Sub
End Class

Public Class HolePattern3WideElement
    Inherits HolePattern
    Public Overrides Sub HolePattern()
        NumHolesDirection1 = 5
    End Sub
End Class

Public Class HolePattern4WideElement
    Inherits HolePattern
    Public Overrides Sub HolePattern()
        NumHolesDirection1 = 8
    End Sub
End Class
