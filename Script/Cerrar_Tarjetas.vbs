If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

Main
Sub Main()
Set objExcel = CreateObject("Excel.Application")
   objExcel.Visible = True ' Haz que Excel sea visible para verificar

' Abrir el archivo Excel
   Set objWorkbook = objExcel.Workbooks.Open("C:\Users\NCDRPRACPROD\Downloads\Close_Cards_Data.xlsx")
   Set objSheet = objWorkbook.Sheets(1) ' Hoja donde están los datos

Dim lastRow, i
lastRow = objSheet.Cells(objSheet.Rows.Count, 1).End(-4162).Row ' Última fila con datos
For i = 2 To lastRow
Dim colNumAvisoCerrar, colDescription

colNumAvisoCerrar = objSheet.Cells(i, 1).Value ' Columna A
colDescription = objSheet.Cells(i, 2).Value ' Columna B

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/NIW22"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = colNumAvisoCerrar
session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").caretPosition = 10
session.findById("wnd[0]").sendVKey 0

session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setUnprotectedTextPart 0,"" + vbCr + colDescription + vbCr + ""
Wscript.Sleep 2000
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").setSelectionIndexes 392,392
session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").firstVisibleLine = "5"
session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1050/btnANWENDERSTATUS").press
session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,4]").selected = true
session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,4]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
Wscript.Sleep 3000
session.findById("wnd[0]/tbar[1]/btn[16]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press

Next
objWorkbook.Close False
objExcel.Quit
End Sub
