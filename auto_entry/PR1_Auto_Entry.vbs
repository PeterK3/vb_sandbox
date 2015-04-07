If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
REM int(a+(abs(fix(a)<>a)))
Set objExcel = GetObject(,"Excel.Application")
Set objSheet = objExcel.ActiveWorkbook.ActiveSheet

lngBegin = 1
lngLast = objSheet.Usedrange.Rows.Count
rem intBulk = 1000
rem sngTemp = objSheet.Usedrange.Rows.Count / 1000
rem intFileT = INT(sngTemp + (ABS(FIX(sngTemp)<>sngTemp)))

rem For intFile = 1 to intFileT
	rem session.findById("wnd[0]").resizeWorkingPane 103,22,false
	rem session.findById("wnd[0]/tbar[0]/okcd").text = "ZTABDIS"
	rem session.findById("wnd[0]").sendVKey 0
	rem session.findById("wnd[0]/usr/ctxtP_TABNAM").text = "MARD"
	rem session.findById("wnd[0]/usr/ctxtP_TABNAM").caretPosition = 4
	rem session.findById("wnd[0]").sendVKey 8
	rem session.findById("wnd[0]/mbar/menu[3]/menu[2]").select
	rem session.findById("wnd[1]/tbar[0]/btn[14]").press
	rem session.findById("wnd[1]/usr/chk[2,5]").selected = true
	rem session.findById("wnd[1]/usr/chk[2,6]").selected = true
	rem session.findById("wnd[1]/tbar[0]/btn[0]").press
	rem session.findById("wnd[0]/usr/ctxtI2-LOW").text = "1010"
	rem session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press
	rem session.findById("wnd[1]/tbar[0]/btn[16]").press

	For lngTemp = lngBegin To lngLast
		strTemp = objSheet.Cells(lngTemp, 1).Value
		session.findById("wnd[1]/usr/sub:SAPLALDB:2020[0]/ctxtRSCSEL-SLOW_I[0,5]").text = TRIM(strTemp)
		rem session.findById("wnd[1]/usr/sub:SAPLALDB:2020[0]/ctxtRSCSEL-SLOW_I[0,5]").caretPosition = 7
		session.findById("wnd[1]").sendVKey 13
		If strTemp = "" Then Exit For
	Next
	
	rem session.findById("wnd[1]/tbar[0]/btn[5]").press
	rem session.findById("wnd[0]/tbar[1]/btn[8]").press
	
	rem session.findById("wnd[0]/usr/txt[57,0]").text = "1023"
	rem session.findById("wnd[0]/mbar/menu[3]/menu[0]/menu[1]").select
	rem session.findById("wnd[1]/tbar[0]/btn[14]").press
	rem session.findById("wnd[1]/tbar[0]/btn[9]").press
	rem session.findById("wnd[1]/usr/chk[1,3]").selected = false
	rem session.findById("wnd[1]/tbar[0]/btn[6]").press
	rem session.findById("wnd[0]/mbar/menu[6]/menu[5]/menu[2]/menu[2]").select
	rem session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]").select
	rem session.findById("wnd[1]/usr/sub:SAPLSPO5:0101/radSPOPLI-SELFLAG[1,0]").setFocus
	rem session.findById("wnd[1]").sendVKey 0
	rem session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text = "C:\Users\10096168\Documents\SAP\MARD 2012-02-09(" & intFile & ").xls"
	rem session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").caretPosition = 51
	rem session.findById("wnd[1]").sendVKey 0
	rem session.findById("wnd[0]/tbar[0]/btn[3]").press
	rem session.findById("wnd[0]/tbar[0]/btn[3]").press
	rem session.findById("wnd[0]/tbar[0]/btn[3]").press
	rem lngBegin = (lngTemp + 1)
rem Next