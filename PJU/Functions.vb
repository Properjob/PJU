
Public Class Functions

    Public Function EncodeURI(StringIn As String) As String
        Return Uri.EscapeDataString(StringIn)
    End Function

    Public Function DecodeURI(StringIn As String) As String
        Return Uri.UnescapeDataString(StringIn)
    End Function

    Public Function GetFilesFromZip(archiveFileName As String) As String()
        Dim fileArchive As IO.Compression.ZipArchive
        Dim Files(0) As String
        '
        fileArchive = IO.Compression.ZipFile.OpenRead(archiveFileName)
        '
        For i = 0 To fileArchive.Entries.Count - 1
            ReDim Preserve Files(i)
            Files(i) = fileArchive.Entries(i).Name
        Next
        '
        Return Files
    End Function

End Class

