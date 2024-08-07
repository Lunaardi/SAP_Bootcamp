Option Explicit

Public Appl, SapGuiAuto, Connection, session, WScript 'Application, SapGuiAuto, Connection, session, WScript'

Sub CARGA_Domic_Fiscal_Sap()

Dim UltimaLinha As Integer
Dim i As Integer
Dim batchSize As Integer
Dim count As Integer

UltimaLinha = ThisWorkbook.Sheets("Sheet1").UsedRange.Rows.count
batchSize = 10
count = 0

'Conexao com o SAP'

If Not IsObject(Appl) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set Appl = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(Connection) Then
   Set Connection = Appl.Children(0)
End If
If Not IsObject(session) Then
   Set session = Connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject Application, "on"
End If

'Rotina Gravada'

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "se16"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").Text = "J_1BTXJURV"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/sub:SAPLSVIX:0100/ctxtD0100_FIELD_TAB-LOWER_LIMIT[0,37]").Text = "BR"
session.findById("wnd[1]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[25]").press
session.findById("wnd[0]/tbar[1]/btn[5]").press 'Entradas Novas'

'Loop para inserção de múltiplos registros'

For i = 2 To UltimaLinha

    session.findById("wnd[0]/usr/tblSAPLJ1BISSCUSTTCTRL_J_1BTXJURV/txtJ_1BTXJURV-TAXJURCODE[0," & (count Mod batchSize) & "]").Text = Cells(i, 3).Value
    session.findById("wnd[0]/usr/tblSAPLJ1BISSCUSTTCTRL_J_1BTXJURV/txtJ_1BTXJURV-TEXT[1," & (count Mod batchSize) & "]").Text = Cells(i, 2).Value
    session.findById("wnd[0]/usr/tblSAPLJ1BISSCUSTTCTRL_J_1BTXJURV/txtJ_1BTXJURV-TEXT[1,0]").SetFocus
    session.findById("wnd[0]/usr/tblSAPLJ1BISSCUSTTCTRL_J_1BTXJURV/txtJ_1BTXJURV-TEXT[1,0]").caretPosition = Len(Cells(i, 2).Value)

    ThisWorkbook.Sheets("Sheet1").Cells(i, 4).Value = 1

    count = count + 1

    ' Verifica se atingiu o batchSize
    If count Mod batchSize = 0 Then
        ' Salva e retorna para inserção de registros
        session.findById("wnd[0]").sendVKey 11
        session.findById("wnd[1]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/tbar[1]/btn[28]").press
        session.findById("wnd[0]/tbar[1]/btn[25]").press
        session.findById("wnd[0]/tbar[1]/btn[5]").press
    End If
  
Next i

' Envia a tecla Enter para salvar os registros restantes após inserir todos
If count Mod batchSize <> 0 Then
    session.findById("wnd[0]").sendVKey 11
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[1]/btn[28]").press
End If

session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0

End Sub

