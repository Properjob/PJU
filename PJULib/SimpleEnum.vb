Public Class SimpleEnum
    Private pValues As New Collection

    ReadOnly Property Values
        Get
            Values = pValues
        End Get
    End Property


    ReadOnly Property Count
        Get
            Count = pValues.Count
        End Get
    End Property

    Public Sub Add(Key As String, Value As Object)
        pValues.Add(Value, Key)
    End Sub

End Class
