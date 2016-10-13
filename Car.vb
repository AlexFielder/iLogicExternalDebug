Public Class Car
    Private currSpeed As Integer
    Private maxSpeed As Integer

    Public Property Speed As Integer
        Get
            Return Nothing
        End Get
        Set(value As Integer)
        End Set
    End Property

    Public Sub Car(max As Integer)
        maxSpeed = max
    End Sub
End Class
