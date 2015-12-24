Public Class VarChar
    ' This class will simply wrap a string in order to enforce length limits
    Private m_strValue As String
    Private m_iLengthLimit As Integer

    Public Property Value() As String
        Get
            Return m_strValue
        End Get
        Set(ByVal Value As String)
            If Value.Length <= m_iLengthLimit Then
                m_strValue = Value
            Else
                m_strValue = ""
                Throw New Exception("Length Limit of VarChar Exceeded")
            End If
        End Set
    End Property

    Sub New(ByVal p_iLengthLimit As Integer)
        m_iLengthLimit = p_iLengthLimit
        m_strValue = ""
    End Sub

End Class
