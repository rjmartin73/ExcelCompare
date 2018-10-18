Option Explicit
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Sub RangeCompare(Optional In_r1 As String, Optional In_r2 As String)
    
    Dim i As Integer, j As Integer, m As Integer, n As Integer, w As Integer, y As Integer, dCount As Integer
    Dim r1 As Range, r2 As Range, r1Address As String, r2Address As String, strDiffCount As String
    Dim rr1() As Variant, rr2() As Variant, rrmatch As Variant
    ActiveSheet.Cells.Borders.LineStyle = xlLineStyleNone
    Dim myStr As String ', In_r1 As String, In_r2 As String
    Dim filePath As String, fileName As String
    i = 1
    j = 0
    
    
    filePath = CStr(VBA.Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\" & Application.Version & "\ExcelCompareFiles")
    fileName = "xlCompare.html"
    
    Do While i <> j
    
    'On Error GoTo ExitMe
    
        If In_r1 = "" Then
            Set r1 = Application.InputBox(prompt:="Enter your first range", Title:="Range 1", Default:="", Type:=8) 'Sheets("Sheet1").Range("A1:H11")  '
            
            r1Address = r1.Worksheet.Name & "!" & r1.Address
                With Range(r1Address)
                    .BorderAround _
                    ColorIndex:=3, LineStyle:=xlDashDotDot, weight:=xlThin
                End With
        
            Set r2 = Application.InputBox(prompt:="Enter your second range", Title:="Range 2", Default:="", Type:=8) 'Sheets("Sheet1").Range("A11:L17") 'Sheets("Sheet1").Range("J1:Q11") '
        Else: Set r1 = Range(In_r1)
                Set r2 = Range(In_r2)
        End If
        
        'Debug.Print r1.Address
        
        r1Address = r1.Worksheet.Name & "!" & r1.Address
        With Range(r1Address)
            .BorderAround _
            ColorIndex:=0, LineStyle:=xlDashDotDot, weight:=xlThin
        End With

        r2Address = r2.Worksheet.Name & "!" & r2.Address
        With Range(r2Address)
            .BorderAround _
            ColorIndex:=0, LineStyle:=xlDashDotDot, weight:=xlThin
        End With
'Exit Sub
        i = r1.Columns.count
        j = r2.Columns.count
        m = r1.Rows.count
        n = r2.Rows.count
        
        
        If i <> j Then
            MsgBox prompt:="You must select the same number of cells for each range.", Buttons:=vbOKOnly
        End If
        
    Loop
    
    ReDim rr1(m - 1, i - 1)
    ReDim rr2(n - 1, j - 1)
    Dim x As Integer, z As Integer
    
    x = 1
    z = 1
    
    ReDim rrmatch(WorksheetFunction.Max(m, n))
    
    
    While x <= m
        While z <= i
            rr1(x - 1, z - 1) = r1.Cells(x, z).Value
            z = z + 1
        Wend
        x = x + 1
        z = 1
    Wend
    
    x = 1
    z = 1
    
    While x <= n
        While z <= j
            rr2(x - 1, z - 1) = r2.Cells(x, z).Value
            z = z + 1
        Wend
        x = x + 1
        z = 1
    Wend
    
    Dim Val_1 As Variant
    Dim Val_2 As Variant
    
    x = 0
    z = 0
    y = 0
    w = 0
    
    While x <= m - 1
        While z <= i - 1
            Val_1 = rr1(x, z)
            Val_2 = rr2(x, z)
            If Val_1 <> Val_2 Then
                w = w + 1
                dCount = dCount + 1
            End If
            z = z + 1
        Wend
        z = 0
        If w > 0 Then
            rrmatch(x) = 0
            Else: rrmatch(x) = 1
        End If
        w = 0
        x = x + 1
    Wend
  
  Call CreateFile(fileName, filePath, i, m, dCount)
  
    x = 0
    z = 0
    y = 0
    w = 0
    dCount = 0
    
    While x <= m - 1
        If x = 0 Then
            myStr = "<thead id='header'><tr><th>&nbsp;</th>"
        Else
            If rrmatch(x) = 1 Then
                myStr = "<tr class='tr-match'  id='r" & x & "'><td class='clickme' onclick='viewRecords(" & x & ")' rowspan=2>" & x & " <span class='badge badge-pill badge-success'>Match</span></td>"
            Else
                myStr = "<tr class='tr-nomatch'  id='r" & x & "'><td class='clickme' onclick='viewRecords(" & x & ")' rowspan=2>" & x & " <span class='badge badge-pill badge-danger'>No Match</span></td>"
                dCount = dCount + 1
            End If
        End If
        While z <= i - 1
            If x = 0 Then
                myStr = myStr & "<th scope='col'>" & rr1(x, z) & "</th>"
            Else
                If rr2(x, z) <> rr1(x, z) Then
                    myStr = myStr & "<td class='nomatch'>" & rr1(x, z) & "</td>"
                Else
                    myStr = myStr & "<td>" & rr1(x, z) & "</td>"
                End If
            End If
            z = z + 1
 'Debug.Print myStr
        Wend
        If rrmatch(x) = 1 And x > 0 Then
                If x > 0 And z Mod 2 = 0 Then
                    myStr = myStr & "</tr><tr class='tr-match-last' id='r2_" & x & "'>"
                ElseIf x > 0 Then
                    myStr = myStr & "</tr><tr class='tr-match-last' id='r2_" & x & "'>"
                End If
        Else
            If x > 0 And z Mod 2 = 0 Then
                myStr = myStr & "</tr><tr class='tr-nomatch-last' id='r2_" & x & "'>"
            ElseIf x > 0 Then
                myStr = myStr & "</tr><tr class='tr-nomatch-last' id='r2_" & x & "'>"
            End If
        End If
        'Call WriteToFile(myStr, fileName, filePath)
        'myStr = ""
        z = 0
        While z <= i - 1
            If x = 0 Then
                myStr = myStr '& "" "<td scope='row'>" & rr2(x, z) & "</td>"
            Else
                If rr2(x, z) <> rr1(x, z) Then
                    myStr = myStr & "<td class='nomatch'>" & rr2(x, z) & "</td>"
                Else
                    myStr = myStr & "<td>" & rr2(x, z) & "</td>"
                End If
            End If
            z = z + 1
 'Debug.Print myStr
        Wend
        If x = 0 Then
            myStr = myStr & "</tr></thead><tbody>"
        Else
            myStr = myStr & "</tr>"
        End If
'Debug.Print myStr
        Call WriteToFile(myStr, fileName, filePath)
        z = 0
        x = x + 1
        myStr = ""
    Wend
    
 Call CloseFile(fileName, filePath, i, m, dCount)
    
    With UserForm1
        .Label1.caption = r1Address
        .Label2.caption = r2Address
    End With
    
    'If UserForm1.Visible = True Then
    '    UserForm1.WebBrowser1.Refresh
    'Else: UserForm1.Show
    'End If
    
Call openWeb
    'ThisWorkbook.FollowHyperlink "file:///C:/Windows/Temp/test.html"
       
    
'Debug.Print "Done"
ExitMe:     Exit Sub

End Sub

Sub CreateFile(fileName As String, filePath As String, colspan As Integer, m As Integer, dCount As Integer)
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim strPath As String
strPath = filePath
Dim strHTML As String
Dim ofile As Object

If FileFolderExists(strPath) = False Then fso.createfolder (strPath)
    
Set ofile = fso.CreateTextFile(strPath & "\" & fileName)

ofile.writeline "<!DOCTYPE html><html><head><meta charset='utf-8'><title>xL Comparison</title>"

'ofile.writeline "<script src='" & strPath & "\compare_xl.js'></script>"
ofile.writeline "<link rel='stylesheet' href='compare_xl.css' />"
ofile.writeline "<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.2/css/bootstrap.min.css' integrity='sha384-PsH8R72JQ3SOdhVi3uxftmaW6Vc51MKb0q5P2rRUpPvrszuE4W1povHYgTpBfshb' crossorigin='anonymous'>"
ofile.writeline "</head><body>"
ofile.writeline "<div id='summary' class='container-fluid'>"
ofile.writeline "<table><tr><td style='border:0;'>Total Records:</td><td style='border:0;'>" & m - 1
ofile.writeline "</td></tr><tr><td style='border:0;'>Matching Records:</td><td style='border:0;'>" & (m - 1) - Abs((dCount))
ofile.writeline "</td></tr><tr><td style='border:0;'>Non-Matching Records:</td><td style='border:0;'>" & Abs((dCount))
ofile.writeline "</td></tr></table></div>"
ofile.writeline "<div id='comparecontainer' class='hidden' role='alert'><a href='#' class='close-thik' onclick='hideMe()'></a>"
ofile.writeline "<table><thead><tr id='compare'><th></th></thead><tr id=record_1></tr><tr id=record_2></tr></table></div>"
ofile.writeline "<div class='container-fluid'><table class='table'>"
'Dim i As Integer
'While i <= 50
'ofile.writeline "<tr id=tr-nomatch><td>36<td>36<td>36<td>36<td>36<td>0<td>0<td>100%<td><td>36<td>36<td>72<td>36<td>36<td>0<td>0<td>200%"
'i = i + 1
'Wend
'ofile.writeline "</tbody></table></body></html>"
'ofile.Close
Set fso = Nothing
Set ofile = Nothing

End Sub
Sub WriteToFile(strHTML As String, fileName As String, filePath As String)
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim strPath As String
strPath = Replace(filePath & "\" & fileName, "\\", "\")
Dim ofile As Object
Set ofile = fso.OpenTextFile(strPath, 8)
ofile.writeline strHTML
Set fso = Nothing
Set ofile = Nothing
End Sub
Sub CloseFile(fileName As String, filePath As String, colspan As Integer, m As Integer, dCount As Integer)
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Dim strPath As String
strPath = filePath & "\" & fileName
Dim ofile As Object
Set ofile = fso.OpenTextFile(strPath, 8)

' These are the paths to the js and css files on my public dropbox share
Dim scriptFile As String
    scriptFile = "https://www.dropbox.com/s/gzilymb7gcm2b3k/compare_xl.js?dl=1"
Dim cssFile As String
    cssFile = "https://www.dropbox.com/s/6mchmefiaqccuzs/compare_xl.css?dl=1"
      
ofile.writeline "</tbody></table></div>"
ofile.writeline "<script src='scripts/jquery-3.2.1.js'></script>"
ofile.writeline "<script src='https://code.jquery.com/ui/1.12.1/jquery-ui.js'></script>"
ofile.writeline "<script src='https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.3/umd/popper.min.js' integrity='sha384-vFJXuSJphROIrBnz7yo7oB41mKfc8JzQZiCq4NCceLEaO4IHwicKwpJf9c9IpFgh' crossorigin='anonymous'></script>"
ofile.writeline "<script src='https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-beta.2/js/bootstrap.min.js' integrity='sha384-alpBpkh1PFOepccYVYDB4do5UnbKysX5WZXm3XxPqe5iKTfUKjNkCk9SaVuEZflJ' crossorigin='anonymous'></script>"
ofile.writeline "<script src='scripts/compare_xl.js'></script>"
ofile.writeline "</body></html>"
ofile.Close
Set fso = Nothing
Set ofile = Nothing

Call dl_compareJS(filePath, "compare_xl.js", scriptFile)
Call dl_compareJS(filePath, "compare_xl.css", cssFile)

End Sub
Function killHyperlinkWarning()
    Dim oShell As Object
    Dim strReg As String

    strReg = "Software\Microsoft\Office\16.0\Common\Security\DisableHyperlinkWarning"

    Set oShell = CreateObject("Wscript.Shell")
    oShell.RegWrite "HKCU\" & strReg, 1, "REG_DWORD"
End Function
Public Function FileFolderExists(strFullPath As String) As Boolean

If strFullPath = vbNullString Then Exit Function
On Error GoTo EarlyExit
If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True

EarlyExit:
    On Error GoTo 0
End Function

Sub dl_compareJS(filePath As String, fileName As String, fileSource As String)
    Dim FileNum As Long
    Dim FileData() As Byte
    Dim myFile As String
    Dim myPath As String
    Dim cssFile As String
    Dim WHTTP As Object
    
    myFile = fileSource
    myPath = filePath
    
    On Error Resume Next
        Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5")
        If Err.number <> 0 Then
            Set WHTTP = CreateObject("WinHTTP.WinHTTPrequest.5.1")
        End If
    On Error GoTo 0
    

    
    ' script file
    WHTTP.Open "GET", fileSource, False
    WHTTP.send
    FileData = WHTTP.responseBody
    Set WHTTP = Nothing
    
    If Dir(filePath, vbDirectory) = Empty Then MkDir myPath
    
    FileNum = FreeFile
    Open myPath & "\" & fileName For Binary Access Write As #FileNum
        Put #FileNum, 1, FileData
    Close #FileNum
    
End Sub
Private Sub openWeb()
 Dim filePath As String, fileName As String
    filePath = CStr(VBA.Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\" & Application.Version & "\ExcelCompareFiles")
    fileName = "xlCompare.html"
    ThisWorkbook.FollowHyperlink Replace("file:///" & filePath & "/" & fileName, "\", "/")
End Sub