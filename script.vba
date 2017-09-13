Sub MajCotations()
Dim i%, k%, URL$, COT
k = Cells(Rows.Count, 2).End(xlUp).Row
Range(Cells(2, 4), Cells(3 + k, 4)).Clear
On Error Resume Next
For i = 2 To k
   DoEvents
          ReDim COT(1 To k, 1 To 1)
                COT(1, 1) = Cells(i, 2).Value
                      URL = Cells(i, 3).Value
    Application.StatusBar = "Mise à jour des cotations en cours …"
    On Error Resume Next
    With CreateObject("MSXML2.XMLHTTP")
            .Open "GET", URL, False
            .Send
            If .Status = 200 Then COT(i, 1) = Val(Split(.responseText, "cotation"">", 2)(1))
    End With
    Application.StatusBar = False
        Cells(i, 4).Value = COT(i, 1)
Next
End Sub
