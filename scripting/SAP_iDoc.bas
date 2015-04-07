Attribute VB_Name = "SAP_IdOC"
Public Function SAP_MM90(EXT_NUMBER As Long)
Dim objApp As Object
Dim objGUI As Object
Dim objConnect As Object
Dim objSession As Object

    If objApp Is Nothing Then
       Set objGUI = GetObject("SAPGUI")
       Set objApp = objGUI.GetScriptingEngine
    End If
    If objConnect Is Nothing Then
       Set objConnect = objApp.Children(0)
    End If
    If objSession Is Nothing Then
       Set objSession = objConnect.Children(0)
    End If
    
    With objSession
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nmm90"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/usr/txtEXTNO").Text = EXT_NUMBER
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]").sendVKey 2
    End With
    
Dim objUsr As Object
    
    Set objUsr = objSession.findById("/app/con[0]/ses[0]/wnd[0]/usr")
    
Dim objChild As Object
Dim strId As String
Dim strType As String

Dim intPreY As Integer
'Dim intPreX As Integer

Dim objBook As Workbook
Dim objSheet As Worksheet

    Set objBook = ActiveSheet.Parent
    Set objSheet = objBook.Worksheets(2)
    objSheet.Select
    
Dim lngRow As Long
Dim intCol As Integer

    lngRow = (objSheet.Range("A1").CurrentRegion.Rows.Count + 1)
    intCol = 1
    intPreY = 0

    For Each objChild In objUsr.Children: DoEvents
        strType = objChild.Type
        If strType = "GuiLabel" Then
            Dim intX As Integer
            Dim intY As Integer
            Dim intFrtBracket As Integer
            Dim intComma As Integer
            Dim intBckBracket As Integer
            
            strId = objChild.ID
            
            intFrtBracket = InStr(1, strId, "lbl[", vbTextCompare) + 4
            intComma = InStr(1, strId, ",", vbTextCompare)
            intBckBracket = Len(strId)
            
            intX = Mid(strId, intFrtBracket, intComma - intFrtBracket)
            intY = Mid(strId, intComma + 1, intBckBracket - (intComma + 1))
            
            'OUTPUT TO EXCEL
            Select Case intPreY
                Case 0
                    intCol = 1
                    objSheet.Cells(lngRow, intCol).Value = objChild.Text
                    intPreY = intY
                Case intY
                    intCol = (intCol + 1)
                    objSheet.Cells(lngRow, intCol).Value = objChild.Text
                Case Else
                    intPreY = intY
                    lngRow = (objSheet.Range("A1").CurrentRegion.Rows.Count + 1)
                    intCol = 1
                    objSheet.Cells(lngRow, intCol).Value = objChild.Text
            End Select
        End If
    Next objChild
End Function

Private Sub TEST_LABEL()
Dim intFrtBracket As Integer
Dim intX As Integer
Dim intComma As Integer
Dim intY As Integer
Dim intBckBracket As Integer

Dim strId As String

    strId = "/app/con[0]/ses[0]/wnd[0]/usr/lbl[16,11]"
    
    intFrtBracket = InStr(1, strId, "lbl[", vbTextCompare) + 4
    intComma = InStr(1, strId, ",", vbTextCompare)
    intBckBracket = Len(strId)
    
    Debug.Print "Front: " & intFrtBracket & " | " & Mid(strId, intFrtBracket, intComma - intFrtBracket)
    Debug.Print "Comma: " & intComma & " | " & Mid(strId, intComma, 1)
    Debug.Print "Back: " & intBckBracket & " | " & Mid(strId, intComma + 1, intBckBracket - (intComma + 1))
End Sub
