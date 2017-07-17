Imports Inventor

Public Class BomRowItem

    Public Property ItemNo() As Long
        Get
            Return m_ItemNo
        End Get
        Set(value As Long)
            m_ItemNo = value
        End Set
    End Property
    Private m_ItemNo As Long

    Private m_occurrence As String
    Public Property Document() As String
        Get
            Return m_occurrence
        End Get
        Set(ByVal value As String)
            m_occurrence = value
        End Set
    End Property

    Private m_Material As String
    Public Property Material() As String
        Get
            Return m_Material
        End Get
        Set(ByVal value As String)
            m_Material = value
        End Set
    End Property

    Private m_Quantity As Integer
    Public Property Quantity() As Integer
        Get
            Return m_Quantity
        End Get
        Set(ByVal value As Integer)
            m_Quantity = value
        End Set
    End Property
    Private m_WhereUsed As List(Of String)
    Public Property WhereUsed() As List(Of String)
        Get
            Return m_WhereUsed
        End Get
        Set(ByVal value As List(Of String))
            m_WhereUsed = value
        End Set
    End Property
End Class
