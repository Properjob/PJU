Module EnumHelper
    Public Example As SimpleEnum
    '
    Dim ExampleLoaded As Boolean

    Public Function ExampleItem(IndexVal) As String
        ExampleItem = Example.Values.Item(IndexVal)
    End Function

    Sub Load()
        Example = New SimpleEnum
        '
        ExampleLoad()
    End Sub

    Sub ExampleLoad()
        Example.Add("Example1", "Just an Example")
    End Sub
End Module
