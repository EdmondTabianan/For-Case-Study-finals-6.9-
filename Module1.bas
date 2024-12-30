Attribute VB_Name = "Module1"
        Public con As New ADODB.Connection
Public AID As Integer

Sub main()
    Set con = New ADODB.Connection
        With con
            .Provider = "MSDASQL"
            .ConnectionString = "Data Source=" & "librarydsn"
            .Open
        End With
            Call Load(frmsplash)
            frmsplash.Show
End Sub

