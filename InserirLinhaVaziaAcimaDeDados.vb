Option Explicit

'====================================================================================
' Sub-rotina: InserirLinhaVaziaAcimaDeDados
' Propósito: Insere uma linha vazia acima de cada linha que contém dados
' Melhorias: Validação, tratamento de erros, performance e documentação
' Autor: Ricardo Cambraia
' Data: 13/07/2025 às 18h: 18m: 35s
' Versão: 2.0 (Revisada)
'====================================================================================
Sub InserirLinhaVaziaAcimaDeDados()
    ' ==================================================================================
    ' DECLARAÇÃO DE VARIÁVEIS
    ' ==================================================================================
    Dim ws As Worksheet                    ' Referência explícita à planilha
    Dim lUltimaLinhaComDados As Long      ' Última linha com dados em qualquer coluna
    Dim j As Long                         ' Contador do loop
    Dim lLinhasInseridas As Long          ' Contador de linhas inseridas
    Dim tempoInicio As Double             ' Para medir performance
    Dim mensagemResultado As String       ' Mensagem final para o usuário
    
    ' ==================================================================================
    ' CONFIGURAÇÕES DE PERFORMANCE
    ' ==================================================================================
    Application.ScreenUpdating = False    ' Desativa atualização da tela
    Application.Calculation = xlCalculationManual  ' Desativa cálculo automático
    Application.EnableEvents = False      ' Desativa eventos (opcional)
    
    ' Marca tempo de início
    tempoInicio = Timer
    
    ' ==================================================================================
    ' DEFINIÇÃO DA PLANILHA E TRATAMENTO DE ERROS
    ' ==================================================================================
    On Error GoTo TratarErro
    
    ' Define a planilha de trabalho (modifique conforme necessário)
    Set ws = ActiveSheet
    ' Alternativa para planilha específica:
    ' Set ws = ThisWorkbook.Worksheets("NomeDaPlanilha")
    
    ' ==================================================================================
    ' VALIDAÇÕES INICIAIS
    ' ==================================================================================
    
    ' Verifica se a planilha está protegida
    If ws.ProtectContents Then
        MsgBox "? A planilha '" & ws.Name & "' está protegida." & vbCrLf & _
               "Desproteja a planilha antes de executar esta operação.", _
               vbExclamation, "Operação Não Permitida"
        GoTo Finalizar
    End If
    
    ' Verifica se há dados na planilha
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        MsgBox "?? A planilha '" & ws.Name & "' está vazia." & vbCrLf & _
               "Não há linhas para processar.", _
               vbInformation, "Planilha Vazia"
        GoTo Finalizar
    End If
    
    ' ==================================================================================
    ' DETERMINAÇÃO DA ÚLTIMA LINHA COM DADOS (MÉTODO ROBUSTO)
    ' ==================================================================================
    
    ' Método mais confiável: encontra a última célula usada em toda a planilha
    Dim ultimaCelula As Range
    Set ultimaCelula = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If ultimaCelula Is Nothing Then
        ' Planilha realmente vazia (backup)
        MsgBox "?? Nenhum dado encontrado na planilha.", vbInformation
        GoTo Finalizar
    Else
        lUltimaLinhaComDados = ultimaCelula.Row
    End If
    
    ' ==================================================================================
    ' CONFIRMAÇÃO DO USUÁRIO (OPCIONAL)
    ' ==================================================================================
    Dim resposta As VbMsgBoxResult
    resposta = MsgBox("?? Esta operação irá inserir linhas vazias acima de cada linha com dados." & vbCrLf & vbCrLf & _
                      "Planilha: " & ws.Name & vbCrLf & _
                      "Linhas a processar: 1 a " & lUltimaLinhaComDados & vbCrLf & vbCrLf & _
                      "?? Esta operação não pode ser desfeita facilmente." & vbCrLf & _
                      "Deseja continuar?", _
                      vbYesNo + vbQuestion, "Confirmar Operação")
    
    If resposta = vbNo Then
        MsgBox "? Operação cancelada pelo usuário.", vbInformation
        GoTo Finalizar
    End If
    
    ' ==================================================================================
    ' LOOP PRINCIPAL DE INSERÇÃO
    ' ==================================================================================
    
    ' Inicializa contador de linhas inseridas
    lLinhasInseridas = 0
    
    ' Loop de baixo para cima (CRUCIAL para manter índices válidos)
    For j = lUltimaLinhaComDados To 1 Step -1
        ' Verifica se a linha atual contém algum dado
        ' CountA conta células não vazias na linha inteira
        If Application.WorksheetFunction.CountA(ws.Rows(j)) <> 0 Then
            ' Insere uma nova linha ACIMA da linha atual
            ws.Rows(j).Insert Shift:=xlDown
            
            ' Incrementa contador
            lLinhasInseridas = lLinhasInseridas + 1
            
            ' Feedback de progresso para operações longas (opcional)
            If lLinhasInseridas Mod 100 = 0 Then
                Application.StatusBar = "Processando... " & lLinhasInseridas & " linhas inseridas"
            End If
        End If
    Next j
    
    ' ==================================================================================
    ' FINALIZAÇÃO E FEEDBACK
    ' ==================================================================================
    
    ' Calcula tempo decorrido
    Dim tempoDecorrido As Double
    tempoDecorrido = Timer - tempoInicio
    
    ' Monta mensagem de resultado
    mensagemResultado = "? Operação concluída com sucesso!" & vbCrLf & vbCrLf & _
                       "?? Resumo da operação:" & vbCrLf & _
                       "   • Planilha processada: " & ws.Name & vbCrLf & _
                       "   • Linhas analisadas: " & lUltimaLinhaComDados & vbCrLf & _
                       "   • Linhas vazias inseridas: " & lLinhasInseridas & vbCrLf & _
                       "   • Tempo decorrido: " & Format(tempoDecorrido, "0.00") & " segundos"
    
    ' Exibe resultado
    MsgBox mensagemResultado, vbInformation, "Operação Concluída"
    
    ' Log opcional para depuração
    Debug.Print "InserirLinhaVaziaAcimaDeDados - " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & _
                " - Linhas inseridas: " & lLinhasInseridas & " - Tempo: " & Format(tempoDecorrido, "0.00") & "s"

Finalizar:
    ' ==================================================================================
    ' RESTAURAÇÃO DAS CONFIGURAÇÕES
    ' ==================================================================================
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False  ' Limpa barra de status
    
    ' Libera objetos da memória
    Set ws = Nothing
    Set ultimaCelula = Nothing
    
    Exit Sub

TratarErro:
    ' ==================================================================================
    ' TRATAMENTO DE ERROS
    ' ==================================================================================
    Dim mensagemErro As String
    
    ' Identifica tipos comuns de erro
    Select Case Err.Number
        Case 1004  ' Erro comum: planilha protegida ou célula bloqueada
            mensagemErro = "? Erro de acesso: A planilha pode estar protegida ou você não tem permissão para inserir linhas."
            
        Case 9      ' Subscript out of range
            mensagemErro = "? Erro de referência: Problema ao acessar a planilha especificada."
            
        Case 1001   ' Interrompido pelo usuário
            mensagemErro = "? Operação interrompida pelo usuário."
            
        Case Else
            mensagemErro = "? Erro inesperado:" & vbCrLf & _
                          "Número: " & Err.Number & vbCrLf & _
                          "Descrição: " & Err.Description & vbCrLf & vbCrLf & _
                          "Entre em contato com o suporte técnico."
    End Select
    
    ' Exibe mensagem de erro
    MsgBox mensagemErro, vbCritical, "Erro na Operação"
    
    ' Log do erro
    Debug.Print "ERRO em InserirLinhaVaziaAcimaDeDados - " & Format(Now(), "dd/mm/yyyy hh:mm:ss") & _
                " - Erro " & Err.Number & ": " & Err.Description
    
    ' Garante que as configurações sejam restauradas mesmo em caso de erro
    Resume Finalizar
End Sub

'====================================================================================
' Sub-rotina: InserirLinhaVaziaComOpcoes
' Propósito: Versão avançada com opções de personalização
' Parâmetros: Permite especificar planilha, range e tipo de inserção
'====================================================================================
Sub InserirLinhaVaziaComOpcoes(Optional ByVal nomePlanilha As String = "", _
                              Optional ByVal primeiraLinha As Long = 1, _
                              Optional ByVal ultimaLinha As Long = 0, _
                              Optional ByVal inserirAcima As Boolean = True)
    
    Dim ws As Worksheet
    Dim j As Long
    Dim lLinhasInseridas As Long
    
    ' Configurações de performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo TratarErroOpcoes
    
    ' Define planilha
    If nomePlanilha = "" Then
        Set ws = ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(nomePlanilha)
    End If
    
    ' Define última linha se não especificada
    If ultimaLinha = 0 Then
        Dim ultimaCelula As Range
        Set ultimaCelula = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        If ultimaCelula Is Nothing Then
            ultimaLinha = 1
        Else
            ultimaLinha = ultimaCelula.Row
        End If
    End If
    
    ' Validação de parâmetros
    If primeiraLinha < 1 Then primeiraLinha = 1
    If ultimaLinha < primeiraLinha Then ultimaLinha = primeiraLinha
    
    lLinhasInseridas = 0
    
    ' Loop principal
    For j = ultimaLinha To primeiraLinha Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(j)) <> 0 Then
            If inserirAcima Then
                ws.Rows(j).Insert Shift:=xlDown
            Else
                ws.Rows(j + 1).Insert Shift:=xlDown
            End If
            lLinhasInseridas = lLinhasInseridas + 1
        End If
    Next j
    
    ' Feedback
    MsgBox "? Operação personalizada concluída!" & vbCrLf & _
           "Linhas inseridas: " & lLinhasInseridas, vbInformation

FinalizarOpcoes:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Set ws = Nothing
    Set ultimaCelula = Nothing
    Exit Sub

TratarErroOpcoes:
    MsgBox "? Erro na operação personalizada: " & Err.Description, vbCritical
    Resume FinalizarOpcoes
End Sub

'====================================================================================
' Sub-rotina: TestarInserirLinhaVazia
' Propósito: Função de teste para validar o funcionamento
'====================================================================================
Sub TestarInserirLinhaVazia()
    Debug.Print "=== TESTE: InserirLinhaVaziaAcimaDeDados ==="
    Debug.Print "Iniciando teste em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    
    ' Executa a função principal
    InserirLinhaVaziaAcimaDeDados
    
    Debug.Print "Teste concluído em: " & Format(Now(), "dd/mm/yyyy hh:mm:ss")
    Debug.Print "=== FIM DO TESTE ==="
End Sub
    
