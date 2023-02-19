Attribute VB_Name = "Module1"

' https://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

Public Function doiToName(doi As String, Optional style As String = "")

Dim xmlhttp As Object
Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
Dim myurl As String

'Methold 1: https://citation.crosscite.org/docs.html
If style <> "" Then
    myurl = "https://citation.crosscite.org/format?doi=" & URLEncode(doi) & "&style=" & style & "&lang=en-US"
Else
    myurl = "https://citation.crosscite.org/format?doi=" & URLEncode(doi) & "&style=g3&lang=en-US" 'long, no [1]
    myurl = "https://citation.crosscite.org/format?doi=" & URLEncode(doi) & "&style=elsevier-vancouver&lang=en-US" 'long with doi, but generates [1]
    myurl = "https://citation.crosscite.org/format?doi=" & URLEncode(doi) & "&style=advanced-materials&lang=en-US" 'short, but generates [1]
End If

xmlhttp.Open "GET", myurl, False
xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

'Method 2: http://doi.fyicenter.com/1000082_crossref_org_API-_works_%7Bdoi%7D_transform.html
'myurl = "https://api.crossref.org/works/" & doi & "/transform/text/x-bibliography"
'xmlhttp.Open "GET", myurl, False
'xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

'Method 3: doi.org ====> failed, dunno why :(
'myurl = "https://doi.org/" & doi
'myurl = "https://doi.org/10.5284/1015681"
'myurl = "https://doi.org/10.1126/science.169.3946.635"
'xmlhttp.Open "GET", myurl, False
'xmlhttp.setRequestHeader "Connection", "keep-alive"
'xmlhttp.setRequestHeader "Accept", "text/x-bibliography; style=modern-language-association; locale=fr-FR"
'xmlhttp.setRequestHeader "style", "apa"

'xmlhttp.setRequestHeader "X-Requested-With", "XMLHTTP60Request"
'xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.102 Safari/537.36"
'xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

' for use inside HA
xmlhttp.setProxy 2, "proxy.ha.org.hk:8080", ""
xmlhttp.setProxyCredentials "cks982", "med2020!"
xmlhttp.send

shortName = Replace(xmlhttp.responseText, vbCrLf, "")
shortName = Right(shortName, Len(shortName) - 3)

doiToName = shortName 'StrConv(.responseBody, vbUnicode)

End Function

Function ExtractDOI(str As String) As String()

'doiRegex = "(10[.][0-9]{2,}(?:[.][0-9]+)*)/(?:(?![%""#? ])\\S)+"

Dim doiRegex(4) As String
' https://www.crossref.org/blog/dois-and-matching-regular-expressions/
' Total doi 74.9M
' This matches 74.4M
doiRegex(0) = "10.\d{4,9}/[-._;()/:A-Z0-9]+$"
' This matches 300k more
doiRegex(1) = "10.1002/[^\s]+$"
' Adding the 3 below leaves 72k unmatched
doiRegex(2) = "10.\d{4}/\d+-\d+X?(\d+)\d+<[\d\w]+:[\d\w]*>\d+.\d+.\w+;\d$"
doiRegex(3) = "10.1021/\w\w\d+$"
doiRegex(4) = "10.1207/[\w\d]+\&\d+_\d+$"

Dim DOIs() As String, k As Long
DOIs = VBA.Strings.Split(vbNullString)
k = 0

For i = 0 To UBound(doiRegex)

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = True
            .Pattern = doiRegex(i)
    End With
    
    If regEx.Test(str) Then
        Set matches = regEx.Execute(str)
        For j = 0 To matches.Count - 1
          ReDim Preserve DOIs(0 To k)
          DOIs(k) = matches(j).Value
          str = Replace(str, matches(j).Value, "", , 1) 'replace once only (though the effect will be same if replace all)
          k = k + 1
        Next
    Else
        
    End If

Next

ExtractDOI = DOIs

End Function

Public Sub MainSub_ReplaceDoiToName()

Dim curSlide As Slide
Dim curShape As Shape

Dim textShapes() As Shape, i As Long
'ReDim textShapes(0 To 0)
i = 0

Dim refDOIs As New Collection ' all DOIs used in reference page

For Each curSlide In ActivePresentation.Slides

    Dim footnotes As String
    footnotes = ""
    
    For Each curShape In curSlide.Shapes
        If curShape.HasTextFrame Then
            ReDim Preserve textShapes(0 To i) As Shape
            Set textShapes(i) = curShape
            ''''''''''''''''''''' DONE: add footnotes for current page DOI '''''''''''''''''''
            Dim text As String
            Dim DOIs() As String
            text = textShapes(i).TextFrame.TextRange.text
            DOIs = ExtractDOI(text)
            
            For j = 0 To UBound(DOIs)
                refDOIs.Add (DOIs(j))
                textShapes(i).TextFrame.TextRange.text = Replace(text, DOIs(j), "")
                footnotes = footnotes & doiToName(DOIs(j))
                'If j <> UBound(DOIs) Then footnotes = footnotes & vbCrLf
            Next
            
            i = i + 1
        End If
    Next curShape
    
    If footnotes <> "" Then
        ' add footnotes
        If Right(footnotes, 1) = vbLf Then footnotes = Left(footnotes, Len(footnotes) - 1)
        footnotes = Replace(footnotes, vbCrLf & vbCrLf, vbCrLf)
        
        Dim ss As Shape
        Set ss = curSlide.Shapes.AddShape(msoShapeRectangle, 0, 0, ActivePresentation.PageSetup.SlideWidth, 10)
        ss.Select
        With ActiveWindow.Selection.ShapeRange
            .Fill.Visible = msoTrue
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(162, 30, 36)
            .Fill.Transparency = 0.5
            .Line.Visible = msoFalse
            
        End With
        ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Select
        ActiveWindow.Selection.ShapeRange.TextFrame.AutoSize = ppAutoSizeShapeToFitText
        ActiveWindow.Selection.ShapeRange.TextFrame.TextRange.Characters(Start:=1, Length:=0).Select
                
        With ActiveWindow.Selection.TextRange
            .text = footnotes
            .ParagraphFormat.Alignment = ppAlignLeft
            With .Font
                .Name = "Arial"
                .Size = 10
                .Bold = msoFalse
                .Italic = msoFalse
                .Underline = msoFalse
                .Shadow = msoFalse
                .Emboss = msoFalse
                .BaselineOffset = 0
                .AutoRotateNumbers = msoFalse
                .Color.SchemeColor = ppForeground
            End With
        End With
        
        ss.Top = ActivePresentation.PageSetup.SlideHeight - ss.Height
        
        
        
    End If
Next curSlide

'''''''''''''' DONE: add all DOIs into reference page '''''''''''''''''''
Dim RefList As String
RefList = ""

For i = 0 To refDOIs.Count - 1
    Dim longName As String
    longName = doiToName(refDOIs.Item(i + 1), "elsevier-vancouver")
    RefList = RefList & longName
    'If i <> refDOIs.Count - 1 Then RefList = RefList & vbCrLf
Next

Dim RefSlide As Slide
Set RefSlide = ActivePresentation.Slides.Add(Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutText)
RefSlide.Shapes(1).TextFrame.TextRange.text = "Reference"
RefSlide.Shapes(2).TextFrame.TextRange.text = RefList

End Sub





