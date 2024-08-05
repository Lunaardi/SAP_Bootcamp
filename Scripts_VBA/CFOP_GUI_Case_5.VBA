Option Explicit

Public Appl, SapGuiAuto, Connection, session, WScript 'Application, SapGuiAuto, Connection, session, WScript'

Sub CARGACFOP_Sap()

    Dim UltimaLinha As Integer
    Dim i As Integer
    Dim icmsExemptReas As Variant
    Dim Finnfe As Variant
    Dim cfotxtValue As Variant
    Dim batchSize As Integer
    Dim count As Integer

    UltimaLinha = Cells(Sheet1.UsedRange.Rows.Count, 1).Row
    batchSize = 27
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

        session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/ctxtJ_1BAGNV-VERSION[0," & (count Mod batchSize) & "]").Text = Cells(i, 1).Value
        session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/txtJ_1BAGNV-CFOP[1," & (count Mod batchSize) & "]").Text = Cells(i, 2).Value

        ' Verifique e ajuste o tamanho do texto para J_1BAGNV-CFOTXT
        cfotxtValue = Cells(i, 3).Value
        If Len(cfotxtValue) > 50 Then
            cfotxtValue = Left(cfotxtValue, 50)
        End If
        session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/txtJ_1BAGNV-CFOTXT[2," & (count Mod batchSize) & "]").Text = cfotxtValue

        If Not IsEmpty(icmsExemptReas) And icmsExemptReas <> "" Then
            session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-ICMS_EXEMPT_REAS[3," & (count Mod batchSize) & "]").Key = icmsExemptReas
        End If

        If Not IsEmpty(Finnfe) And Finnfe <> "" Then
            session.findById("wnd[0]/usr/tblSAPLJ1BVTCTRL_J_1BAGNV/cmbJ_1BAGNV-FINNFE[4," & (count Mod batchSize) & "]").Key = Finnfe
        End If

        Cells(i, 6).Value = 1
        count = count + 1

        ' Verifique se atingiu o tamanho do lote
        If count Mod batchSize = 0 Then
            ' Envia a tecla Enter para salvar os registros após inserir o lote
            session.findById("wnd[0]").sendVKey 0

            ' Retorna à tela de entrada de novos registros
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
            session.findById("wnd[0]").sendVKey 0

            ' Reinicie o processo de entrada de novos registros
            session.findById("wnd[0]/tbar[0]/okcd").Text = "sm30"
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/ctxtVIEWNAME").Text = "J_1BAGNV"
            session.findById("wnd[0]/usr/ctxtVIEWNAME").caretPosition = 8
            session.findById("wnd[0]/usr/btnUPDATE_PUSH").press
            session.findById("wnd[0]/tbar[1]/btn[5]").press 'Entradas Novas'
        End If

    Next i

    ' Envia a tecla Enter para salvar os registros restantes após inserir todos
    If count Mod batchSize <> 0 Then
        session.findById("wnd[0]").sendVKey 0
    End If

    ' Retorna à tela inicial
    session.findById("wnd[0]/tbar[0]/btn[11]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0

End Sub
