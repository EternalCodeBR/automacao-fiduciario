VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ==========================
' CONFIGURAÇÃO GERAL
' ==========================
' NOTA: este arquivo é publicado em CP1252 no ambiente real (ver
' src/sincronizar_modulo_vba.py). No repositório está em UTF-8 apenas para
' leitura — ajuste o encoding ao reutilizar em um projeto VBA real.
Private Const FORMATO_DATA As String = "dd/mm/yyyy"
Private Const DATA_INICIA_EM As Long = 2

' Ignorar as primeiras N linhas do HISTÓRICO ao procurar a data (protege o cabeçalho)
Private Const HIST_IGNORE_TOP_ROWS As Long = 30

' CONFIGURE: pasta raiz local onde ficam as calculadoras e históricos
Private Const BASE_LOCAL As String = "C:\caminho\para\pasta\raiz"
' CONFIGURE: URL do repositório remoto (SharePoint/OneDrive) usado como fallback
' quando o arquivo local não está disponível
Private Const SHAREPOINT_URL As String = "https://exemplo.sharepoint.com/sites/SeuTime/Documentos/"
Private Const HIST_PASSWORD As String = ""

' Caminho da planilha de dados de mercado (fonte de taxas para as
' calculadoras). A macro garante que ela esteja aberta durante todo o processo.
Private Const BASE_DADOS_MERCADO_PATH As String = "C:\caminho\para\BaseDadosMercado.xlsm"
Private Const BASE_DADOS_MERCADO_NOME As String = "BaseDadosMercado.xlsm"

' ==========================
' MAPA DOS PARES
' ==========================
Private Const IDX_CalcPathRel  As Long = 1
Private Const IDX_CalcSheet    As Long = 2
Private Const IDX_CalcRange    As Long = 3   ' "B:S" - primeira col = col da data na calc
Private Const IDX_HistPathRel  As Long = 4
Private Const IDX_HistSheet    As Long = 5
Private Const IDX_HistRange    As Long = 6   ' "B:S" - primeira col = col da data no hist
Private Const IDX_AllowPartial As Long = 7   ' True = se D0 não existir na calc, vai até a última data presente

' Lista de ativos monitorados. Os nomes/caminhos reais dos ativos foram
' substituídos por identificadores genéricos (Ativo A..D) — ajuste conforme
' a estrutura real de pastas do seu ambiente.
Private Function GetPairs() As Variant
    Dim a(1 To 4, 1 To 7) As Variant

    ' 1) Ativo A - calc B:S, hist B:S, data na B (em ambos)
    a(1, 1) = "Fluxo de pagamento\Calculadora\Ativo A (Agente Fiduciário)\Calculadora Ativo A.xlsx"
    a(1, 2) = "":          a(1, 3) = "B:S"
    a(1, 4) = "PUs site\Historico_PU_Ativo_A.xlsx"
    a(1, 5) = "":          a(1, 6) = "B:S"
    a(1, 7) = False

    ' 2) Ativo B - aba "PU" da calc, B:S em ambos, data na B
    a(2, 1) = "Fluxo de pagamento\Calculadora\Ativo B (Agente Fiduciário)\Calculadora Ativo B.xlsx"
    a(2, 2) = "PU":        a(2, 3) = "B:S"
    a(2, 4) = "PUs site\Historico_PU_Ativo_B.xlsx"
    a(2, 5) = "":          a(2, 6) = "B:S"
    a(2, 7) = False

    ' 3) Ativo C - calc B:R, hist B:R, data na B
    a(3, 1) = "Fluxo de pagamento\Calculadora\Ativo C (Agente Fiduciário)\Calculadora Ativo C.xlsx"
    a(3, 2) = "":          a(3, 3) = "B:R"
    a(3, 4) = "PUs site\Historico_PU_Ativo_C.xlsx"
    a(3, 5) = "":          a(3, 6) = "B:R"
    a(3, 7) = False

    ' 4) Ativo D - calc B:T (data na C, col B descartada), hist B:S (data na B).
    a(4, 1) = "Fluxo de pagamento\Calculadora\Ativo D (Agente Fiduciário)\Calculadora Ativo D.xlsx"
    a(4, 2) = "":          a(4, 3) = "C:T"
    a(4, 4) = "PUs site\Historico_PU_Ativo_D.xlsx"
    a(4, 5) = "":          a(4, 6) = "B:S"
    a(4, 7) = False

    GetPairs = a
End Function

' ==========================
' FUNÇÃO PRINCIPAL
' ==========================
Public Sub Atualizar_Historicos_PU()
    Dim sData As Variant, alvo As Date
    Dim pairs As Variant, i As Long
    Dim ok As Long, logText As String
    Dim partes() As String
    Dim wbBaseDados As Workbook
    Dim baseDadosAbertaAqui As Boolean

    On Error GoTo fim
    Application.ScreenUpdating = True
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationAutomatic

    Set wbBaseDados = GarantirBaseDadosAberta(baseDadosAbertaAqui)
    If wbBaseDados Is Nothing Then
        logText = logText & "AVISO: planilha de dados de mercado nao pode ser aberta." & vbCrLf
    Else
        logText = logText & "INFO: planilha de dados de mercado " & IIf(baseDadosAbertaAqui, "aberta pela macro", "ja estava aberta") & "." & vbCrLf
    End If

    sData = Application.InputBox("Informe a DATA (ex.: 22/10/2025):", "Data base", Type:=2)
    If (VarType(sData) = vbBoolean And sData = False) Or Trim$(CStr(sData)) = "" Then GoTo fim
    partes = Split(Trim$(CStr(sData)), "/")
    If UBound(partes) <> 2 Then MsgBox "Data inválida. Use o formato dd/mm/aaaa.": GoTo fim
    alvo = DateSerial(CLng(partes(2)), CLng(partes(1)), CLng(partes(0)))

    pairs = GetPairs()
    For i = LBound(pairs, 1) To UBound(pairs, 1)
        If ProcessarPar( _
            CStr(pairs(i, IDX_CalcPathRel)), CStr(pairs(i, IDX_CalcSheet)), CStr(pairs(i, IDX_CalcRange)), _
            CStr(pairs(i, IDX_HistPathRel)), CStr(pairs(i, IDX_HistSheet)), CStr(pairs(i, IDX_HistRange)), _
            CBool(pairs(i, IDX_AllowPartial)), alvo, logText) Then ok = ok + 1
    Next i

    If Len(logText) = 0 Then logText = "Sem mensagens adicionais."
    MsgBox "Concluído." & vbCrLf & _
           "Colou com sucesso em " & ok & " de " & (UBound(pairs, 1)) & " pares." & vbCrLf & vbCrLf & logText, _
           IIf(ok = UBound(pairs, 1), vbInformation, vbExclamation)

fim:
    ' Fecha a planilha de dados de mercado apenas se a macro abriu
    If baseDadosAbertaAqui And Not (wbBaseDados Is Nothing) Then
        On Error Resume Next
        wbBaseDados.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' ==========================
' ENTRADA AUTOMATIZADA (chamada via COM pelo Python)
' Recebe a data como string dd/mm/aaaa e retorna o log; sem InputBox e sem MsgBox.
' Execução em segundo plano: ScreenUpdating desligado.
' ==========================
Public Function Atualizar_Historicos_PU_Auto(ByVal dataStr As String) As String
    Dim alvo As Date
    Dim pairs As Variant, i As Long
    Dim ok As Long, logText As String
    Dim partes() As String
    Dim wbBaseDados As Workbook
    Dim baseDadosAbertaAqui As Boolean

    On Error GoTo fim
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationAutomatic

    Set wbBaseDados = GarantirBaseDadosAberta(baseDadosAbertaAqui)
    If wbBaseDados Is Nothing Then
        logText = logText & "AVISO: planilha de dados de mercado nao pode ser aberta." & vbCrLf
    Else
        logText = logText & "INFO: planilha de dados de mercado " & IIf(baseDadosAbertaAqui, "aberta pela macro", "ja estava aberta") & "." & vbCrLf
    End If

    partes = Split(Trim$(dataStr), "/")
    If UBound(partes) <> 2 Then
        Atualizar_Historicos_PU_Auto = "ERRO: data inválida (use dd/mm/aaaa). Recebido: " & dataStr
        GoTo fim
    End If
    alvo = DateSerial(CLng(partes(2)), CLng(partes(1)), CLng(partes(0)))

    pairs = GetPairs()
    For i = LBound(pairs, 1) To UBound(pairs, 1)
        If ProcessarPar( _
            CStr(pairs(i, IDX_CalcPathRel)), CStr(pairs(i, IDX_CalcSheet)), CStr(pairs(i, IDX_CalcRange)), _
            CStr(pairs(i, IDX_HistPathRel)), CStr(pairs(i, IDX_HistSheet)), CStr(pairs(i, IDX_HistRange)), _
            CBool(pairs(i, IDX_AllowPartial)), alvo, logText) Then ok = ok + 1
    Next i

    If Len(logText) = 0 Then logText = "Sem mensagens adicionais."
    Atualizar_Historicos_PU_Auto = "OK " & ok & "/" & UBound(pairs, 1) & vbCrLf & logText

fim:
    ' Fecha a planilha de dados de mercado apenas se a macro abriu (Python pode ter aberto antes).
    If baseDadosAbertaAqui And Not (wbBaseDados Is Nothing) Then
        On Error Resume Next
        wbBaseDados.Close SaveChanges:=False
        On Error GoTo 0
    End If
    If Len(Atualizar_Historicos_PU_Auto) = 0 Then _
        Atualizar_Historicos_PU_Auto = "ERRO: falha inesperada na execução."
    ' Grava o log em arquivo para que o orquestrador Python consiga ler
    ' mesmo quando Application.Run não propaga o retorno pelo COM.
    On Error Resume Next
    Dim fNum As Integer, logPath As String
    logPath = Environ$("TEMP") & "\pu_macro_log.txt"
    fNum = FreeFile
    Open logPath For Output As #fNum
    Print #fNum, Atualizar_Historicos_PU_Auto
    Close #fNum
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

' ==========================
' PROCESSAMENTO DE UM PAR
' Lê a última data do histórico, localiza-a na calc, e copia (só valores) o
' bloco de linhas da calc da próxima linha até a linha que contém ALVO (D0),
' colando no hist a partir da próxima linha vazia. Replica formato da última
' linha do hist em todas as novas linhas.
' ==========================
Private Function ProcessarPar( _
    ByVal CalcPathRel As String, ByVal CalcSheet As String, ByVal CalcRange As String, _
    ByVal HistPathRel As String, ByVal HistSheet As String, ByVal HistRange As String, _
    ByVal AllowPartial As Boolean, _
    ByVal alvo As Date, ByRef logText As String) As Boolean

    On Error GoTo falha

    Dim wbCalc As Workbook, wbHist As Workbook
    Dim wsCalc As Worksheet, wsHist As Worksheet
    Dim partesCalc() As String, partesHist() As String
    Dim calcFirstCol As Long, calcLastCol As Long
    Dim histFirstCol As Long, histLastCol As Long
    Dim calcWidth As Long, histWidth As Long
    Dim lastHistRow As Long, lastHistDate As Date
    Dim linhaUltimaHistNaCalc As Long
    Dim calcStartRow As Long, calcEndRow As Long
    Dim destRow As Long, numRows As Long
    Dim vals As Variant
    Dim wasProtected As Boolean
    Dim dataInicial As Date, dataFinal As Date

    ' --- abre arquivos ---
    Set wbCalc = AbrirLocalOuSharePoint(CalcPathRel)
    If wbCalc Is Nothing Then logText = logText & "- Não abriu calculadora: " & CalcPathRel & vbCrLf: GoTo sair
    Set wbHist = AbrirLocalOuSharePoint(HistPathRel)
    If wbHist Is Nothing Then logText = logText & "- Não abriu histórico: " & HistPathRel & vbCrLf: GoTo sair

    ' --- abas ---
    Set wsCalc = EscolherAbaPreferindo(wbCalc, CalcSheet, Array("pu", "calc", "calculadora", "sheet"))
    Set wsHist = EscolherAbaPreferindo(wbHist, HistSheet, Array("hist", "pu", "cálculo", "sheet"))

    ' --- parse dos ranges ---
    partesCalc = Split(CalcRange, ":")
    partesHist = Split(HistRange, ":")
    If UBound(partesCalc) <> 1 Or UBound(partesHist) <> 1 Then
        logText = logText & "- " & NomeArq(wbCalc.fullname) & ": range inválido (Calc=" & CalcRange & " Hist=" & HistRange & ")." & vbCrLf
        GoTo sair
    End If
    calcFirstCol = ColLetraParaNum(partesCalc(0))
    calcLastCol = ColLetraParaNum(partesCalc(1))
    histFirstCol = ColLetraParaNum(partesHist(0))
    histLastCol = ColLetraParaNum(partesHist(1))
    calcWidth = calcLastCol - calcFirstCol + 1
    histWidth = histLastCol - histFirstCol + 1
    If calcWidth <> histWidth Or calcWidth <= 0 Then
        logText = logText & "- " & NomeArq(wbCalc.fullname) & ": larguras incompatíveis (Calc=" & calcWidth & " Hist=" & histWidth & ")." & vbCrLf
        GoTo sair
    End If

    ' --- última linha preenchida do hist (pela col de data) ---
    lastHistRow = wsHist.Cells(wsHist.Rows.Count, histFirstCol).End(xlUp).Row
    If lastHistRow < (HIST_IGNORE_TOP_ROWS + 1) Then
        logText = logText & "- " & NomeArq(wbHist.fullname) & ": não achou linha de dados (última preenchida na col " & partesHist(0) & " = " & lastHistRow & ")." & vbCrLf
        GoTo sair
    End If
    If Not IsDate(wsHist.Cells(lastHistRow, histFirstCol).Value) Then
        logText = logText & "- " & NomeArq(wbHist.fullname) & ": última célula (" & wsHist.Cells(lastHistRow, histFirstCol).Address(False, False) & ") não contém data." & vbCrLf
        GoTo sair
    End If
    lastHistDate = DateValue(wsHist.Cells(lastHistRow, histFirstCol).Value)

    ' --- nada a fazer se hist já está >= alvo ---
    If CLng(lastHistDate) >= CLng(alvo) Then
        logText = logText & "= " & NomeArq(wbHist.fullname) & ": já atualizado (última " & Format$(lastHistDate, FORMATO_DATA) & " >= alvo " & Format$(alvo, FORMATO_DATA) & ")." & vbCrLf
        ProcessarPar = True
        GoTo sair_ok
    End If

    ' --- localiza lastHistDate na calc ---
    linhaUltimaHistNaCalc = EncontrarLinhaPorData(wsCalc, partesCalc(0), lastHistDate, DATA_INICIA_EM)
    If linhaUltimaHistNaCalc < 0 Then
        logText = logText & "- " & NomeArq(wbCalc.fullname) & ": última data do hist (" & Format$(lastHistDate, FORMATO_DATA) & ") não encontrada na calc." & vbCrLf
        GoTo sair
    End If
    calcStartRow = linhaUltimaHistNaCalc + 1

    ' --- localiza alvo (D0) na calc ---
    calcEndRow = EncontrarLinhaPorData(wsCalc, partesCalc(0), alvo, calcStartRow)
    If calcEndRow < 0 Then
        If AllowPartial Then
            calcEndRow = wsCalc.Cells(wsCalc.Rows.Count, calcFirstCol).End(xlUp).Row
            If calcEndRow < calcStartRow Then
                logText = logText & "= " & NomeArq(wbCalc.fullname) & ": já no fim da calc (liquidado)." & vbCrLf
                ProcessarPar = True
                GoTo sair_ok
            End If
        Else
            logText = logText & "- " & NomeArq(wbCalc.fullname) & ": alvo " & Format$(alvo, FORMATO_DATA) & " não encontrado na calc." & vbCrLf
            GoTo sair
        End If
    End If

    numRows = calcEndRow - calcStartRow + 1
    If numRows < 1 Then
        logText = logText & "= " & NomeArq(wbHist.fullname) & ": nada a colar." & vbCrLf
        ProcessarPar = True
        GoTo sair_ok
    End If

    ' --- lê bloco de valores (só .Value, sem fórmula) ---
    ' Garante que a CALCULADORA terminou de resolver links/calculos.
    ' Sequencia: UpdateLink (refresca links externos) -> RefreshAll (data conn)
    ' -> Calculation=Automatic -> CalculateFullRebuild -> aguarda xlDone.
    On Error Resume Next
    Dim linkSources As Variant
    linkSources = wbCalc.LinkSources(xlExcelLinks)
    If Not IsEmpty(linkSources) Then wbCalc.UpdateLink Name:=linkSources, Type:=xlExcelLinks
    wbCalc.RefreshAll
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild
    Dim spinCount As Long
    For spinCount = 1 To 100
        DoEvents
        If Application.CalculationState = xlDone Then Exit For
    Next spinCount

    ' Diagnóstico: registra fórmula e valor da primeira célula que será copiada
    Dim firstCell As Range
    Set firstCell = wsCalc.Cells(calcStartRow, calcFirstCol + 1)
    logText = logText & "DBG " & NomeArq(wbCalc.FullName) & " " & firstCell.Address(False, False) & _
              ": HasFormula=" & firstCell.HasFormula & " Formula=" & Chr(34) & firstCell.Formula & Chr(34) & " Value=" & CStr(firstCell.Value) & vbCrLf
    On Error GoTo falha

    vals = wsCalc.Range( _
              wsCalc.Cells(calcStartRow, calcFirstCol), _
              wsCalc.Cells(calcEndRow, calcLastCol)).Value

    destRow = lastHistRow + 1

    ' --- desproteger se necessário ---
    wasProtected = wsHist.ProtectContents Or wsHist.ProtectDrawingObjects Or wsHist.ProtectScenarios
    If wasProtected Then
        If Len(HIST_PASSWORD) > 0 Then
            On Error Resume Next
            wsHist.Unprotect Password:=HIST_PASSWORD
            On Error GoTo falha
        Else
            logText = logText & "- " & NomeArq(wbHist.fullname) & ": aba protegida (sem senha informada)." & vbCrLf
            GoTo sair
        End If
    End If

    ' --- cola valores ---
    wsHist.Range( _
        wsHist.Cells(destRow, histFirstCol), _
        wsHist.Cells(destRow + numRows - 1, histLastCol)).Value = vals

    ' --- replica formato da última linha do hist nas novas linhas (só nas col do range) ---
    wsHist.Range(wsHist.Cells(lastHistRow, histFirstCol), wsHist.Cells(lastHistRow, histLastCol)).Copy
    wsHist.Range( _
        wsHist.Cells(destRow, histFirstCol), _
        wsHist.Cells(destRow + numRows - 1, histLastCol)).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    If wasProtected Then wsHist.Protect Password:=HIST_PASSWORD, DrawingObjects:=True, Contents:=True, Scenarios:=True

    wbHist.Save

    dataInicial = DateValue(wsCalc.Cells(calcStartRow, calcFirstCol).Value)
    dataFinal = DateValue(wsCalc.Cells(calcEndRow, calcFirstCol).Value)
    ProcessarPar = True
    logText = logText & "OK " & NomeArq(wbCalc.fullname) & " -> " & NomeArq(wbHist.fullname) & _
              " | " & numRows & " linha(s) [" & Format$(dataInicial, FORMATO_DATA) & ".." & Format$(dataFinal, FORMATO_DATA) & _
              "] em " & wsHist.Cells(destRow, histFirstCol).Address(False, False) & ":" & _
              wsHist.Cells(destRow + numRows - 1, histLastCol).Address(False, False) & vbCrLf

sair_ok:
    On Error Resume Next
    If Not (wbCalc Is Nothing) Then wbCalc.Close SaveChanges:=False
    If Not (wbHist Is Nothing) Then wbHist.Close SaveChanges:=True
    On Error GoTo 0
    Exit Function

sair:
    On Error Resume Next
    If Not (wbCalc Is Nothing) Then wbCalc.Close SaveChanges:=False
    If Not (wbHist Is Nothing) Then wbHist.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

falha:
    logText = logText & "- Erro em " & CalcPathRel & " -> " & HistPathRel & ": " & Err.Description & vbCrLf
    ProcessarPar = False
    Resume sair
End Function

' ==========================
' FUNÇÕES DE SUPORTE
' ==========================
Private Function GarantirBaseDadosAberta(ByRef abriuAqui As Boolean) As Workbook
    Dim wb As Workbook
    abriuAqui = False
    On Error Resume Next
    Set wb = Workbooks(BASE_DADOS_MERCADO_NOME)
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set GarantirBaseDadosAberta = wb
        Exit Function
    End If
    Dim ev As Boolean
    ev = Application.EnableEvents
    Application.EnableEvents = True
    On Error Resume Next
    Set wb = Workbooks.Open(Filename:=BASE_DADOS_MERCADO_PATH, UpdateLinks:=0)
    If Not wb Is Nothing Then
        wb.RefreshAll
        Application.CalculateFullRebuild
        Dim sp As Long
        For sp = 1 To 100
            DoEvents
            If Application.CalculationState = xlDone Then Exit For
        Next sp
        abriuAqui = True
    End If
    Application.EnableEvents = ev
    On Error GoTo 0
    Set GarantirBaseDadosAberta = wb
End Function

Private Function AbrirLocalOuSharePoint(relPath As String) As Workbook
    Dim wb As Workbook, fullLocal As String, fullSp As String
    Dim eventsState As Boolean
    eventsState = Application.EnableEvents
    Application.EnableEvents = True
    fullLocal = JuntarLocal(relPath)
    If Len(fullLocal) > 0 Then
        If Dir(fullLocal, vbNormal) <> "" Then
            Set wb = Workbooks.Open(fullLocal)
            Application.EnableEvents = eventsState
            Set AbrirLocalOuSharePoint = wb
            Exit Function
        End If
    End If
    fullSp = SHAREPOINT_URL & Replace(relPath, "\", "/")
    On Error Resume Next
    Set wb = Workbooks.Open(fullSp)
    On Error GoTo 0
    Application.EnableEvents = eventsState
    Set AbrirLocalOuSharePoint = wb
End Function

Private Function JuntarLocal(relPath As String) As String
    If InStr(1, relPath, ":\", vbTextCompare) > 0 Or Left$(relPath, 2) = "\\" Then
        JuntarLocal = relPath: Exit Function
    End If
    If Len(BASE_LOCAL) = 0 Then JuntarLocal = "": Exit Function
    Dim basePath As String: basePath = BASE_LOCAL
    If Right$(basePath, 1) <> "\" Then basePath = basePath & "\"
    JuntarLocal = basePath & relPath
End Function

Private Function EscolherAbaPreferindo(wb As Workbook, byName As String, prefers As Variant) As Worksheet
    Dim ws As Worksheet, p As Variant
    If Len(byName) > 0 Then
        On Error Resume Next
        Set ws = wb.Worksheets(byName)
        On Error GoTo 0
        If Not ws Is Nothing Then Set EscolherAbaPreferindo = ws: Exit Function
    End If
    For Each p In prefers
        For Each ws In wb.Worksheets
            If InStr(1, LCase$(ws.Name), CStr(p), vbTextCompare) > 0 Then Set EscolherAbaPreferindo = ws: Exit Function
        Next ws
    Next p
    Set EscolherAbaPreferindo = wb.Worksheets(1)
End Function

Private Function EncontrarLinhaPorData(ws As Worksheet, colLetra As String, alvo As Date, startRow As Long) As Long
    Dim colNum As Long, rngCol As Range, cell As Range
    Dim serialAlvo As Long, serialCell As Long
    EncontrarLinhaPorData = -1

    colNum = ColLetraParaNum(colLetra)
    If colNum <= 0 Then Exit Function

    Set rngCol = ws.Range(ws.Cells(startRow, colNum), ws.Cells(ws.Rows.Count, colNum).End(xlUp))
    serialAlvo = CLng(alvo)

    For Each cell In rngCol.Cells
        If Not IsEmpty(cell.Value) Then
            On Error Resume Next
            serialCell = CLng(CDate(cell.Value))
            On Error GoTo 0
            If serialCell = serialAlvo Then
                EncontrarLinhaPorData = cell.Row
                Exit Function
            End If
        End If
    Next cell
End Function

Private Function ColLetraParaNum(colLetter As String) As Long
    Dim i As Long, n As Long
    colLetter = UCase$(Trim$(colLetter))
    If Len(colLetter) = 0 Then Exit Function
    For i = 1 To Len(colLetter)
        n = n * 26 + (Asc(Mid$(colLetter, i, 1)) - 64)
    Next i
    ColLetraParaNum = n
End Function

Private Function ColNumParaLetra(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(((n - 1) Mod 26) + 65) & s
        n = (n - 1) \ 26
    Loop
    ColNumParaLetra = s
End Function

Private Function NomeArq(fullname As String) As String
    Dim i As Long: i = InStrRev(fullname, "\")
    If i > 0 Then
        NomeArq = Mid$(fullname, i + 1)
    Else
        NomeArq = fullname
    End If
End Function
