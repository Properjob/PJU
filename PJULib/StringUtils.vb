Module StringUtils

    Public Function StringUtils() As String
        StringUtils = "1.2" ' Version

        ' 1.0 Initial Work, Added Piece, Extract, TextFromCell
        ' 1.1 Fix to Extract
        ' 1.2 Fix to TextFromCell
    End Function

    Public Function Piece(str As String, Delim As String, StartPiece As Integer, Optional EndPiece As Integer = -1) As String
        Dim RetStr As String
        Dim RetArr() As String
        Dim LBnd As Integer
        Dim UBnd As Integer
        '
        RetStr = ""
        '
        ' Check that delim exists in string
        If Not InStr(1, str, Delim) > 0 Then
            RetStr = str
            GoTo Finish
        End If
        '
        ' Split into array
        RetArr = Strings.Split(str, Delim)
        LBnd = LBound(RetArr) + 1
        UBnd = UBound(RetArr) + 1
        '
        ' Check StartPiece is Lower than the last piece
        If StartPiece > UBnd Then GoTo Finish
        If StartPiece = -1 Then StartPiece = UBnd
        '
        If Not EndPiece = 0 Then
            If EndPiece = -1 Then EndPiece = UBnd
            If EndPiece > UBnd Then GoTo Finish
        Else : EndPiece = StartPiece
        End If
        '
        ' Give pieces requested
        For Pce = StartPiece - 1 To EndPiece - 1
            RetStr = RetStr & RetArr(Pce)
            If Not Pce = EndPiece - 1 Then RetStr = RetStr & Delim
        Next Pce

Finish:
        Piece = RetStr
    End Function

    Public Function Extract(str As String, Optional StartPos As Integer = 1, Optional EndPos As Integer = -1)
        Dim StrLen As Integer
        '
        StrLen = Len(str)
        '
        If (StartPos * -1 > StrLen) Then StartPos = StrLen * -1
        If (EndPos * -1 > StrLen) Then EndPos = StrLen * -1
        If EndPos < 0 Then EndPos = StrLen + (EndPos + 1)
        If StartPos < 0 Then StartPos = StrLen + (StartPos + 1)
        Extract = Strings.Mid$(str, StartPos, EndPos - StartPos + 1)
    End Function

    Public Function Contains(str As String, StrIn As String, Delim As String) As Boolean
        Dim SplitArr() As String
        Dim RetStr As Boolean
        Dim Pce As Integer
        '
        RetStr = False
        '
        SplitArr = Strings.Split(StrIn, Delim)

        For Pce = LBound(SplitArr) To UBound(SplitArr)
            If SplitArr(Pce) = str Then
                RetStr = True
                Exit For
            End If
        Next Pce

        Contains = RetStr
    End Function

    Sub Test()
        Console.Write(Extract("123456", -1, -1))
    End Sub

End Module
