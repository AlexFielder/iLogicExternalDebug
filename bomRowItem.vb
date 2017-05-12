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
End Class
