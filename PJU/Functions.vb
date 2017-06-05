
Public Class Functions

    Public Function EncodeURI(StringIn As String) As String
        Return Uri.EscapeDataString(StringIn)
    End Function

    Public Function DecodeURI(StringIn As String) As String
        Return Uri.UnescapeDataString(StringIn)
    End Function

End Class

