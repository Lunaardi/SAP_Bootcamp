Option Explicit

Public Appl, SapGuiAuto, Connection, session, WScript 'Application, SapGuiAuto, Connection, session, WScript'


Sub CARGACFOP_Sap()

Dim UltimaLinha As Integer
Dim i As Integer
Dim icmsExemptReas As Variant
Dim Finnfe As Variant


UltimaLinha = Cells(Planilha1.UsedRange.Rows.Count, 1).Row


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
session.findById("wnd[0]/tbar[0]/okcd").Text = "sm30"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVIEWNAME").Text = "J_1BAGNV"
session.findById("wnd[0]/usr/ctxtVIEWNAME").caretPosition = 8
session.findById("wnd[0]/usr/btnUPDATE_PUSH").press
session.findById("wnd[0]/tbar[1]/btn[5]").press 'Entradas Novas'

 
'Loop para inserção de múltiplos registros'

For i = 2 To UltimaLinha

'Obtenha os valores das células
icmsExemptReas = Cells(i, 4).Value
Finnfe = Cells(i, 5).Value

'session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/ctxtJ_1BAGNV-VERSION[0,0]").Text = Cells(i, 1).Value
'session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/txtJ_1BAGNV-CFOP[1,0]").Text = Cells(i, 2).Value
'session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/txtJ_1BAGNV-CFOTXT[2,0]").Text = Cells(i, 3).Value


session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/ctxtJ_1BAGNV-VERSION[0," & (i - 2) & "]").Text = Cells(i, 1).Value
session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/txtJ_1BAGNV-CFOP[1," & (i - 2) & "]").Text = Cells(i, 2).Value
session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/txtJ_1BAGNV-CFOTXT[2," & (i - 2) & "]").Text = Cells(i, 3).Value


' Verifique se há registro na célula Cells(i, 4)
'If Not IsEmpty(icmsExemptReas) Then
'    session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-ICMS_EXEMPT_REAS[3,0]").Key = icmsExemptReas
'End If

' Verifique se há registro na célula Cells(i, 4)
If Not IsEmpty(icmsExemptReas) Then
    session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-ICMS_EXEMPT_REAS[3," & (i - 2) & "]").Key = icmsExemptReas
End If

'session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-ICMS_EXEMPT_REAS[3,0]").Key = Cells(i, 4).Value

' Verifique se há registro na célula Cells(i, 5)
'If Not IsEmpty(Finnfe) Then
'    session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-FINNFE[4,0]").Key = Finnfe
'End If


' Verifique se há registro na célula Cells(i, 5)
If Not IsEmpty(Finnfe) Then
    session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-FINNFE[4," & (i - 2) & "]").Key = Finnfe
End If

'session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-FINNFE[4,0]").Key = Cells(i, 5).Value

'session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-FINNFE[4,0]").SetFocus

Cells(i, 6).Value = 1

    ' Vá para a próxima linha na tabela SAP
    session.findById("wnd[0]").SetFocus

Next i

' Envia a tecla Enter para salvar os registros após inserir todos
session.findById("wnd[0]").sendVKey 0


session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
session.findById("wnd[0]").sendVKey 0


End Sub


