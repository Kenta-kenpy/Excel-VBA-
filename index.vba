Sub GenerateHTML()
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' HTMLの生成
    Dim html As String
    html = "<!DOCTYPE html>" & vbCrLf
    html = html & "<html>" & vbCrLf
    html = html & "  <head>" & vbCrLf
    html = html & "    <meta charset=""utf-8"">" & vbCrLf
    html = html & "    <title>Progate</title>" & vbCrLf
    html = html & "  </head>" & vbCrLf
    html = html & "  <body>" & vbCrLf
    html = html & "    <h1 class=""title"">Excel VBA</h1>" & vbCrLf
    html = html & "    <a href=""https://github.com/Kenta-kenpy/Excel-VBA-"">マクロ</a>" & vbCrLf
    html = html & "  </body>" & vbCrLf
    html = html & "</html>"
    
    ' HTMLを表示
    ie.Visible = True
    ie.navigate "about:blank"
    ie.document.write html
End Sub
