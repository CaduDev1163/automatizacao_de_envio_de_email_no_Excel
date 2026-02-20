Attribute VB_Name = "Email1EnviarTabelaEmail"

--->Código VBA de envio de relatórios (tabelas) para o e-mail (Outlook)

    
Sub EnviarTabelaEmail()

    Dim wsDados As Worksheet
    Dim wsConfig As Worksheet
    Dim OutApp As Object
    Dim OutMail As Object
    Dim ultimaLinha As Long
    Dim i As Long

    Dim emails As String
    Dim tipoRelatorio As String

    Dim rngRelatorio As Range
    Dim rngStatus As Range

    Dim htmlRelatorio As String
    Dim htmlStatus As String
    Dim corpoEmail As String
    
    Dim caminhoImagem As String
    

    ' ===== PLANILHAS =====
    Set wsDados = ThisWorkbook.Sheets("Relatorios")
    Set wsConfig = ThisWorkbook.Sheets("Config_Email")

	--->Para esse caso específico os relatórios (tabelas) estavam sendo separados por uma função do Excel chamada "Segmentação de Dados". O usuário precisa selecionar um item da segmentação para que gere um relatório (tabela) <---
    ' ===== LER SEGMENTAÇÃO DE DADOS =====
    tipoRelatorio = ObterTipoPelaSegmentacao("segmentaçãodeDados")
    
    If tipoRelatorio = "" Then
        MsgBox "Selecione apenas UM item na segmentação antes de enviar o e-mail.", vbExclamation
        Exit Sub
    End If
 

	---> A tabela relatorio e a tabela status geradas sempre começavam em uma célula específica. <---
    ' ===== TABELAS (INÍCIO FIXO) =====
    Set rngRelatorio = wsDados.Range("C15").CurrentRegion
    Set rngStatus = wsDados.Range("A24").CurrentRegion

	--->Para conseguir edita-las ao ponto do Outlook compreender é necessário converte-las para HTML. <---
    ' ===== CONVERTE PARA HTML =====
    htmlRelatorio = RangetoHTML(rngRelatorio)
    htmlStatus = RangetoHTML(rngStatus)


	---> Aqui é toda a estrutura que será impressa no corpo de texto do e-mail outlook <---
    ' ===== CORPO DO E-MAIL =====
    corpoEmail = _
         "<div style='font-family:Calibri, Arial, sans-serif; font-size:11pt; color:#333333;'>" & _
    "<h2 style='color:#1F4E79; margin-bottom:10px;'>MONITORAMENTO VEICULOS ATRASADOS/DEAD LINE</h2>" & _
    "<h3>Bom dia!</h3>" & _
    "<h3>Poderiam nos informar os horários em que os relatórios serão enviados?</h3>" & _
    "<h3>Caso ocorra algum problema entrar em contato.</h3>" & _

    "<hr style='border:1px solid #D9D9D9;'>" & _
    "<h2 style='color:#1F4E79; margin-top:20px;'>Status:</h2>" & _
    htmlRelatorio & _
    "<br>" & _
    "<h2 style='color:#1F4E79; margin-top:20px;'>Relatório:</h2>" & _
    htmlStatus & _
"</div>"
 

    
---> VBA abre o Outlook <---
    ' ===== OUTLOOK =====
    Set OutApp = CreateObject("Outlook.Application")

    ' ===== ÚLTIMA LINHA CONFIG =====
    ultimaLinha = wsConfig.Cells(wsConfig.Rows.Count, "A").End(xlUp).Row

    ' ===== LOOP DE ENVIO =====
    For i = 2 To ultimaLinha

        If wsConfig.Cells(i, "A").Value = tipoRelatorio Then

            emails = wsConfig.Cells(i, "C").Value

            Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .To = emails
        .Subject = "RELATORIO ONLINE " & tipoRelatorio
         
        .htmlBody = corpoEmail & AssinaturaPadrao
        .Display ' troque para .Send se quiser envio automático
    End With
            
            Set OutMail = Nothing
        End If

    Next i

    Set OutApp = Nothing

End Sub

---> Função que trabalha toda a robusta lógica que identifica o item da segmentação para saber para qual e-mail sera enviado o relatório. Exemplo: item x -> gera relatorio x -> que deve ser enviado para e-mail x. <---

Function ObterTipoPelaSegmentacao(segmentaçãodeDados As String) As String

    Dim sc As SlicerCache
    Dim si As SlicerItem
    Dim selecionados As Long
    Dim valor As String

    ' Acessa o SlicerCache corretamente
    Set sc = ThisWorkbook.SlicerCaches(segmentaçãodeDados)

    ' Percorre os itens da segmentação
    For Each si In sc.SlicerCacheLevels(1).SlicerItems
        If si.Selected Then
            selecionados = selecionados + 1
            valor = si.Caption
        End If
    Next si

    ' Validação
    If selecionados <> 1 Then
        ObterTipoPelaSegmentacao = ""
    Else
        ObterTipoPelaSegmentacao = valor
    End If

End Function

---> Função em que formata toda a parte de assinatura do corpo de texto do email Outlook <---
Function AssinaturaPadrao() As String

    AssinaturaPadrao = _
        "<br><br>" & _
        "<hr>" & _
        "<p style='font-family:Calibri; font-size:10pt; color:#555555;'>" & _
        "<br>Função em que formata toda a parte de assinatura do corpo de texto do email Outlook<br>"& _
        "</p>"

End Function



---> Função que converte os intervalos das tabelas em HTML para que o Outlook entenda e possa copiar em seu corpo de texto email <---

Function RangetoHTML(rng As Range) As String
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
    ' Copia o intervalo
    rng.Copy
    ' Cria um workbook temporário
    Set TempWB = Workbooks.Add(1)
    
    'Problema de layout "resolvido" com esse bloco de código"
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=xlPasteAll
        .Cells(1).PasteSpecial xlPasteValuesAndNumberFormats
        .Cells(1).PasteSpecial Paste:=8
        '.Cells(1).PasteSpecial xlPasteValues
        .Cells(1).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
    End With

    ' Salva como HTML
    TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=TempFile, _
        Sheet:=TempWB.Sheets(1).Name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic) _
        .Publish True
    

    ' Lê o HTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll

    ' Fecha tudo
    ts.Close
    TempWB.Close SaveChanges:=False
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function






