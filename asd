Sub CDO_Send_Selection_Or_Range_Body()
    Dim rng As Range
    Dim iMsg As Object
    Dim iConf As Object
    '    Dim Flds As Variant

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

    iConf.Load -1    ' CDO Source Defaults
      Set Flds = iConf.Fields
      With Flds
           .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
           .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
           .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "konmalyshkin@gmail.com"
           .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""
           .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

           .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
           .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
           .Update
      End With


    Set rng = Nothing
    On Error Resume Next

    Set rng = Selection.SpecialCells(xlCellTypeVisible)

    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    With iMsg
        Set .Configuration = iConf
        .To = "konmalyshkin@gmail.com"
        .CC = ""
        .BCC = ""
        .From = """Ron"" <ron@something.nl>"
        .Subject = "This is a test"
        .HTMLBody = RangetoHTML(rng)
        .Send
    End With

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub
