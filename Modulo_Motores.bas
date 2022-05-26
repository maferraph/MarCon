Attribute VB_Name = "Modulo_Motores"
 Option Explicit
'Todo o posicionamento será realizado direto na tela que estiver comandando a marcaçao,
'sendo que esta só utilizará as funcoes deste modulo para posicionamento e marcação

'Configuração dos pinos do DB-25 (RS232)
'Saídas: Dados e Controle
'Entradas: Status
'
'Nome dos pinos (número do pino no conector)
'Dados: D0(P2), D1(P3) , D2(P4) , D3(P5) , D4(P6) , D5(P7) , D6(P8) , D7(P9)
'Controle: Strobe(P1) , AutoFeed(P14) , Init(P16) , SlctIn(P17)
'Status: Aknowledge(P10) , Busy(P11) , PaperEnd(P12) , SlctOut(P13) , Error(P15)

'********************************** LPT1 **********************************
'Configurações das Ligações dos pinos x hardware
'Motor de passo de 6 fios
'
'MOTOR 1 - motor eixo X
'P2: bobina 1 - fio amarelo
'P3: bobina 2 - fio vermelho
'P4: bobina 3 - fio marrom
'P5: bobina 4 - fio branco
'P1: Vdc 12v primeiro e segundo enrrolamento
'
'MOTOR 2 - motor eixo Y
'P6: bobina 1 - fio amarelo
'P7: bobina 2 - fio vermelho
'P8: bobina 3 - fio marrom
'P9: bobina 4 - fio branco
'P14: Vdc 12v primeiro e segundo enrrolamento
'
'PISTÃO
'P16: solenóide do pistão
'
'SENSOR HOME EIXO X
'P10: chave eixo X
'
'SENSOR HOME EIXO Y
'P11: chave eixo Y
'
'SENSOR PISTÃO RECUADO
'P12: pistão recuado

'Funcoes externas para controle da porta paralela
Public Declare Function LePorta Lib "inpout32.dll" _
Alias "Inp32" (ByVal EnderecoPortaH As Integer) As Integer

Public Declare Sub EscrevePorta Lib "inpout32.dll" _
Alias "Out32" (ByVal EnderecoPortaH As Integer, ByVal VALOR As Integer)

'Constantes para a maquina
Public Const ICONST_TEMPO_ESPERA_PASSO_MOTOR As Integer = 1 'em milisegundos
Private Const SGCONST_PASSO_ROSCA_EIXO_MOTOR As Single = 2.11582 'em milimetros
Private Const ICONST_PASSO_PONTO_CARACTER7x7 As Integer = 4 'em passos
Private Const ICONST_ESPACO_ENTRE_CARACTER7x7 As Integer = 6 'em passos
Private Const ICONST_NUMERO_PASSOS_MOTOR As Integer = 48 'número de passos do motor
Public Const DCONST_MOVIMENTO_POR_PASSO As Single = (SGCONST_PASSO_ROSCA_EIXO_MOTOR / ICONST_NUMERO_PASSOS_MOTOR) 'em milimetros o quanto anda o motor por passo
Private Const ICONST_DISTANCIA_MAXIMA_EIXO_X As Integer = 200 'em milimetros
Private Const ICONST_DISTANCIA_MAXIMA_EIXO_Y As Integer = 200 'em milimetros
Private Const ICONST_PASSO_MAXIMO_EIXO_X As Integer = (ICONST_DISTANCIA_MAXIMA_EIXO_X / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR 'em passos
Private Const ICONST_PASSO_MAXIMO_EIXO_Y As Integer = (ICONST_DISTANCIA_MAXIMA_EIXO_Y / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR 'em passos
Public IVAR_PASSO_X, IVAR_PASSO_Y As Integer 'numero de passos dado pelo motor
Public SGVAR_POSICAO_X, SGVAR_POSICAO_Y As Single 'posição dos eixos em milimetros
Private BVAR_MOVE_EIXOX, BVAR_MOVE_EIXOY As Boolean   'determina se eixo terá movimento
Private IVAR_PASSODESTINO_EIXOX, IVAR_PASSODESTINO_EIXOY As Integer 'posição destino de X e Y em passos

'variáveis dos nibbles dos motores
Public BVAR_LPT1_P1, BVAR_LPT1_P2, BVAR_LPT1_P3, BVAR_LPT1_P4, BVAR_LPT1_P5, BVAR_LPT1_P6, BVAR_LPT1_P7, BVAR_LPT1_P8, BVAR_LPT1_P9, BVAR_LPT1_P10, BVAR_LPT1_P11, BVAR_LPT1_P12, BVAR_LPT1_P13, BVAR_LPT1_P14, BVAR_LPT1_P15, BVAR_LPT1_P16, BVAR_LPT1_P17 As Byte
Private BYVAR_BYTE_DADOS, BYVAR_BYTE_CONTROLE, BYVAR_BYTE_STATUS As Byte
Private SVAR_DADOS, SVAR_CONTROLE, SVAR_STATUS As String
Private Const SCONST_LPT1_DADOS As String = &H378
Private Const SCONST_LPT1_CONTROLE As String = &H37A
Private Const SCONST_LPT1_STATUS As String = &H379

'demais variáveis
Public VETOR_POSICIONAMENTO_MARCACAO As Variant
Public VETOR_CARACTER As Variant

'***************************************************************************
'                FUNCOES PARA TODOS MODELOS DE MARCADORES
'***************************************************************************

Public Sub PosicionamentoPlano_Home()
    PosicionaMarcadorChapinha 0, 0
End Sub



Private Function PegaVetorPontosMarcadosCaracter7x7(ByVal LETRA As String, ByVal POSX As Double, ByVal POSY As Double) As Variant
    Dim VETOR_MARCACAO As Variant
    Dim I, J, LINHA, COLUNA As Integer
    Dim VALOR As Integer
    VETOR_MARCACAO = Array()
    For I = 0 To UBound(VETOR_CARACTER) 'vetor com mapa de caracteres
        'procura letra dentro do vetor
        If VETOR_CARACTER(I)(0) = LETRA Then
            SVAR_TEMP = VETOR_CARACTER(I)(1)
            For J = 1 To Len(SVAR_TEMP)
                'determina número da linha e da coluna do caracter
                If J >= 1 And J <= 7 Then
                    COLUNA = J
                    LINHA = 1
                ElseIf J >= 8 And J <= 14 Then
                    COLUNA = J - 7
                    LINHA = 2
                ElseIf J >= 15 And J <= 21 Then
                    COLUNA = J - 14
                    LINHA = 3
                ElseIf J >= 22 And J <= 28 Then
                    COLUNA = J - 21
                    LINHA = 4
                ElseIf J >= 29 And J <= 35 Then
                    COLUNA = J - 28
                    LINHA = 5
                ElseIf J >= 36 And J <= 42 Then
                    COLUNA = J - 35
                    LINHA = 6
                ElseIf J >= 43 And J <= 49 Then
                    COLUNA = J - 42
                    LINHA = 7
                End If
                'verifica se o ponto deve ser marcado
                If Mid(SVAR_TEMP, J, 1) = 1 Then
                    ReDim Preserve VETOR_MARCACAO(UBound(VETOR_MARCACAO) + 1)
                    'a marcacao do ponto sera da seguinte maneira:
                    'POSX e POSY equivalem ao caracter da L1C1 (0,0), para cada nova coluna
                    'e cada nova linha será adicionado a constante de espaço entre marcaçoes
                    'o vetor criado com o ponto que deverá ser marcado: VETOR=(X,Y)
                    'I-1 e J-1 pois a C1L1 = POSX e POSY, demais caracteres sim adiciona espaço
                    VETOR_MARCACAO(UBound(VETOR_MARCACAO)) = Array((POSX + ((COLUNA - 1) * ICONST_PASSO_PONTO_CARACTER7x7)), (POSY + ((LINHA - 1) * ICONST_PASSO_PONTO_CARACTER7x7)))
                End If
            Next J
            'arruma o vetor de marcação em ordem decrescrente de X e Y
            PegaVetorPontosMarcadosCaracter7x7 = VETOR_MARCACAO
            Exit Function
        End If
    Next I
End Function

'********************************************************************************************************************************
'SCRIPT ANTIGO QUE PEGA OS CARACTERES DIRETO COM O ARQUIVO
'
'Private Function PegaVetorPontosMarcadosCaracter7x7(ByVal LETRA As String, ByVal POSX As Double, ByVal POSY As Double) As Variant
'    SVAR_ARQUIVO = App.Path & "\7x7.car"
'    Dim VETOR_MARCACAO As Variant
'    Dim I, J As Integer
'    Dim VALOR As Integer
'    VETOR_MARCACAO = Array()
'    For I = 1 To 7 'colunas dos 7 caracteres - POSX
'        For J = 1 To 7 'linhas dos 7 caracteres - POSY
'            VALOR = LeINI(SVAR_ARQUIVO, LETRA, "L" & I & "C" & J)
'            If VALOR = 1 Then 'o bit deve ser marcador, portanto adicionar vetor
'                ReDim Preserve VETOR_MARCACAO(UBound(VETOR_MARCACAO) + 1)
'                'a marcacao do ponto sera da seguinte maneira:
'                'POSX e POSY equivalem ao caracter da L1C1 (0,0), para cada nova coluna
'                'e cada nova linha será adicionado a constante de espaço entre marcaçoes
'                'o vetor criado com o ponto que deverá ser marcado: VETOR=(X,Y)
'                'I-1 e J-1 pois a C1L1 = POSX e POSY, demais caracteres sim adiciona espaço
'                VETOR_MARCACAO(UBound(VETOR_MARCACAO)) = Array((POSX + ((J - 1) * ICONST_PASSO_PONTO_CARACTER7x7)), (POSY + ((I - 1) * ICONST_PASSO_PONTO_CARACTER7x7)))
'            End If
'        Next J
'    Next I
'    'arruma o vetor de marcação em ordem decrescrente de X e Y
'    PegaVetorPontosMarcadosCaracter7x7 = VETOR_MARCACAO
'End Function
'********************************************************************************************************************************

'***************************************************************************
'           FUNCOES PARA CONTROLE DOS MOTORES E PORTAS PARALELAS
'***************************************************************************

Public Sub SetaTimers()
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Enabled = False
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Interval = ICONST_TEMPO_ESPERA_PASSO_MOTOR
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_HOME.Enabled = False
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_HOME.Interval = ICONST_TEMPO_ESPERA_PASSO_MOTOR
End Sub
Public Sub ZeraPinos()
    'inicializa as variaveis da porta paralela
    'valor inicial com todos os pinos desligados (em baixo), sendo que os pinos de controle e status trabalham com logica invertida, portanto valor =1
    BVAR_LPT1_P1 = 1
    BVAR_LPT1_P2 = 0
    BVAR_LPT1_P3 = 0
    BVAR_LPT1_P4 = 0
    BVAR_LPT1_P5 = 0
    BVAR_LPT1_P6 = 0
    BVAR_LPT1_P7 = 0
    BVAR_LPT1_P8 = 0
    BVAR_LPT1_P9 = 0
    BVAR_LPT1_P10 = 1
    BVAR_LPT1_P11 = 1
    BVAR_LPT1_P12 = 1
    BVAR_LPT1_P13 = 1
    BVAR_LPT1_P14 = 1
    BVAR_LPT1_P15 = 1
    BVAR_LPT1_P16 = 1
    BVAR_LPT1_P17 = 1
End Sub
Public Sub ZeraPortas()
    EscrevePortaDados
    EscrevePortaControle
    EscrevePortaStatus
End Sub
Private Sub EscrevePortaDados()
    BYVAR_BYTE_DADOS = Binario2Decimal(BVAR_LPT1_P9 & BVAR_LPT1_P8 & BVAR_LPT1_P7 & BVAR_LPT1_P6 & BVAR_LPT1_P5 & BVAR_LPT1_P4 & BVAR_LPT1_P3 & BVAR_LPT1_P2)
    EscrevePorta SCONST_LPT1_DADOS, BYVAR_BYTE_DADOS
End Sub
Private Sub EscrevePortaControle()
    BYVAR_BYTE_CONTROLE = Binario2Decimal(BVAR_LPT1_P17 & BVAR_LPT1_P16 & BVAR_LPT1_P14 & BVAR_LPT1_P1)
    EscrevePorta SCONST_LPT1_CONTROLE, BYVAR_BYTE_CONTROLE
End Sub
Private Sub EscrevePortaStatus()
    BYVAR_BYTE_STATUS = Binario2Decimal(BVAR_LPT1_P15 & BVAR_LPT1_P13 & BVAR_LPT1_P12 & BVAR_LPT1_P11 & BVAR_LPT1_P10)
    EscrevePorta SCONST_LPT1_STATUS, BYVAR_BYTE_STATUS
End Sub
Public Sub LePortaDados()
    SVAR_DADOS = Decimal2Binario(LePorta(SCONST_LPT1_CONTROLE))
    BVAR_LPT1_P2 = Mid(SVAR_DADOS, 8, 1)
    BVAR_LPT1_P3 = Mid(SVAR_DADOS, 7, 1)
    BVAR_LPT1_P4 = Mid(SVAR_DADOS, 6, 1)
    BVAR_LPT1_P5 = Mid(SVAR_DADOS, 5, 1)
    BVAR_LPT1_P6 = Mid(SVAR_DADOS, 4, 1)
    BVAR_LPT1_P7 = Mid(SVAR_DADOS, 3, 1)
    BVAR_LPT1_P8 = Mid(SVAR_DADOS, 2, 1)
    BVAR_LPT1_P9 = Mid(SVAR_DADOS, 1, 1)
End Sub
Public Sub LePortaControle()
    SVAR_CONTROLE = Decimal2Binario(LePorta(SCONST_LPT1_CONTROLE))
    BVAR_LPT1_P1 = Mid(SVAR_DADOS, 4, 1)
    BVAR_LPT1_P14 = Mid(SVAR_DADOS, 3, 1)
    BVAR_LPT1_P16 = Mid(SVAR_DADOS, 2, 1)
    BVAR_LPT1_P17 = Mid(SVAR_DADOS, 1, 1)
End Sub
Public Sub LePortaStatus()
    SVAR_STATUS = Decimal2Binario(LePorta(SCONST_LPT1_CONTROLE))
    BVAR_LPT1_P10 = Mid(SVAR_DADOS, 5, 1)
    BVAR_LPT1_P11 = Mid(SVAR_DADOS, 4, 1)
    BVAR_LPT1_P12 = Mid(SVAR_DADOS, 3, 1)
    BVAR_LPT1_P13 = Mid(SVAR_DADOS, 2, 1)
    BVAR_LPT1_P15 = Mid(SVAR_DADOS, 1, 1)
End Sub
Public Sub AtuaPistao_MarcadorChapinha()
    BVAR_LPT1_P16 = 0 'atua pistão para marcar
    EscrevePortaControle
    BVAR_LPT1_P16 = 1 'retorna pistão
End Sub
Public Sub MoveEixos_MarcadorChapinha()
    'Nesta função estão os códigos para mover os 2 eixos ao mesmo tempo X e Y para
    'economizar tempo e ganhar em velocidade na máquina, caso contrário cada eixo terá
    'q aguardar o movimento do outro até executar o seu, dobrando o tempo praticamente
    
    'verifica se o proximo passo ultrapassa o valor do destino, portanto, posicao proximo passo = destino para o programa nao ficar em loop
    If IVAR_PASSODESTINO_EIXOX > IVAR_PASSO_X And IVAR_PASSODESTINO_EIXOX < (IVAR_PASSO_X + 1) Then
        IVAR_PASSO_X = IVAR_PASSODESTINO_EIXOX
    ElseIf IVAR_PASSODESTINO_EIXOX < SGVAR_POSICAO_X And IVAR_PASSODESTINO_EIXOX > (IVAR_PASSO_X - 1) Then
        SGVAR_POSICAO_X = IVAR_PASSODESTINO_EIXOX
    End If
    If IVAR_PASSODESTINO_EIXOY > IVAR_PASSO_X And IVAR_PASSODESTINO_EIXOY < (IVAR_PASSO_X + 1) Then
        IVAR_PASSO_X = IVAR_PASSODESTINO_EIXOY
    ElseIf IVAR_PASSODESTINO_EIXOY < IVAR_PASSO_X And IVAR_PASSODESTINO_EIXOY > (IVAR_PASSO_X - 1) Then
        IVAR_PASSO_X = IVAR_PASSODESTINO_EIXOY
    End If
    
    '***** EIXO X *****
    If IVAR_PASSODESTINO_EIXOX <= ICONST_PASSO_MAXIMO_EIXO_X And IVAR_PASSODESTINO_EIXOX > IVAR_PASSO_X Then
        'adiciona passo e posição
        IVAR_PASSO_X = IVAR_PASSO_X + 1
        SGVAR_POSICAO_X = IVAR_PASSO_X * DCONST_MOVIMENTO_POR_PASSO
        BVAR_MOVE_EIXOX = True
        BVAR_LPT1_P1 = 0  'liga alimentacao motor eixo x
        'SENTIDO HORARIO
        If BVAR_LPT1_P2 = 1 And BVAR_LPT1_P3 = 1 And BVAR_LPT1_P4 = 0 And BVAR_LPT1_P5 = 0 Then
            BVAR_LPT1_P2 = 0
            BVAR_LPT1_P3 = 1
            BVAR_LPT1_P4 = 1
            BVAR_LPT1_P5 = 0
        ElseIf BVAR_LPT1_P2 = 0 And BVAR_LPT1_P3 = 1 And BVAR_LPT1_P4 = 1 And BVAR_LPT1_P5 = 0 Then
            BVAR_LPT1_P2 = 0
            BVAR_LPT1_P3 = 0
            BVAR_LPT1_P4 = 1
            BVAR_LPT1_P5 = 1
        ElseIf BVAR_LPT1_P2 = 0 And BVAR_LPT1_P3 = 0 And BVAR_LPT1_P4 = 1 And BVAR_LPT1_P5 = 1 Then
            BVAR_LPT1_P2 = 1
            BVAR_LPT1_P3 = 0
            BVAR_LPT1_P4 = 0
            BVAR_LPT1_P5 = 1
        ElseIf BVAR_LPT1_P2 = 1 And BVAR_LPT1_P3 = 0 And BVAR_LPT1_P4 = 0 And BVAR_LPT1_P5 = 1 Then
            BVAR_LPT1_P2 = 1
            BVAR_LPT1_P3 = 1
            BVAR_LPT1_P4 = 0
            BVAR_LPT1_P5 = 0
        Else
            BVAR_LPT1_P2 = 1
            BVAR_LPT1_P3 = 1
            BVAR_LPT1_P4 = 0
            BVAR_LPT1_P5 = 0
        End If
    ElseIf IVAR_PASSODESTINO_EIXOX >= 0 And IVAR_PASSODESTINO_EIXOX < IVAR_PASSO_X Then
        'subtrai passo e posição
        IVAR_PASSO_X = IVAR_PASSO_X - 1
        SGVAR_POSICAO_X = IVAR_PASSO_X * DCONST_MOVIMENTO_POR_PASSO
        BVAR_MOVE_EIXOX = True
        BVAR_LPT1_P1 = 0 'liga alimentacao motor eixo x
        'ANTIHORARIO
        If BVAR_LPT1_P2 = 1 And BVAR_LPT1_P3 = 1 And BVAR_LPT1_P4 = 0 And BVAR_LPT1_P5 = 0 Then
            BVAR_LPT1_P2 = 1
            BVAR_LPT1_P3 = 0
            BVAR_LPT1_P4 = 0
            BVAR_LPT1_P5 = 1
        ElseIf BVAR_LPT1_P2 = 0 And BVAR_LPT1_P3 = 1 And BVAR_LPT1_P4 = 1 And BVAR_LPT1_P5 = 0 Then
            BVAR_LPT1_P2 = 1
            BVAR_LPT1_P3 = 1
            BVAR_LPT1_P4 = 0
            BVAR_LPT1_P5 = 0
        ElseIf BVAR_LPT1_P2 = 0 And BVAR_LPT1_P3 = 0 And BVAR_LPT1_P4 = 1 And BVAR_LPT1_P5 = 1 Then
            BVAR_LPT1_P2 = 0
            BVAR_LPT1_P3 = 1
            BVAR_LPT1_P4 = 1
            BVAR_LPT1_P5 = 0
        ElseIf BVAR_LPT1_P2 = 1 And BVAR_LPT1_P3 = 0 And BVAR_LPT1_P4 = 0 And BVAR_LPT1_P5 = 1 Then
            BVAR_LPT1_P2 = 0
            BVAR_LPT1_P3 = 0
            BVAR_LPT1_P4 = 1
            BVAR_LPT1_P5 = 1
        Else
            BVAR_LPT1_P2 = 0
            BVAR_LPT1_P3 = 0
            BVAR_LPT1_P4 = 1
            BVAR_LPT1_P5 = 1
        End If
    Else
        BVAR_MOVE_EIXOX = False
        'se entrar aqui significa que o eixo já está na posição destino ou não deve ser mover, portanto deve desligar o motor
        BVAR_LPT1_P1 = 1 'desliga alimentacao eixo x
    End If

    '***** EIXO Y *****
    If IVAR_PASSODESTINO_EIXOY <= ICONST_PASSO_MAXIMO_EIXO_Y And IVAR_PASSODESTINO_EIXOY > IVAR_PASSO_Y Then
        'adiciona passo e posição
        IVAR_PASSO_Y = IVAR_PASSO_Y + 1
        SGVAR_POSICAO_Y = IVAR_PASSO_Y * DCONST_MOVIMENTO_POR_PASSO
        BVAR_MOVE_EIXOY = True
        BVAR_LPT1_P14 = 0 'liga alimentacao eixo y
        If BVAR_LPT1_P6 = 1 And BVAR_LPT1_P7 = 1 And BVAR_LPT1_P8 = 0 And BVAR_LPT1_P9 = 0 Then
            BVAR_LPT1_P6 = 0
            BVAR_LPT1_P7 = 1
            BVAR_LPT1_P8 = 1
            BVAR_LPT1_P9 = 0
        ElseIf BVAR_LPT1_P6 = 0 And BVAR_LPT1_P7 = 1 And BVAR_LPT1_P8 = 1 And BVAR_LPT1_P9 = 0 Then
            BVAR_LPT1_P6 = 0
            BVAR_LPT1_P7 = 0
            BVAR_LPT1_P8 = 1
            BVAR_LPT1_P9 = 1
        ElseIf BVAR_LPT1_P6 = 0 And BVAR_LPT1_P7 = 0 And BVAR_LPT1_P8 = 1 And BVAR_LPT1_P9 = 1 Then
            BVAR_LPT1_P6 = 1
            BVAR_LPT1_P7 = 0
            BVAR_LPT1_P8 = 0
            BVAR_LPT1_P9 = 1
        ElseIf BVAR_LPT1_P6 = 1 And BVAR_LPT1_P7 = 0 And BVAR_LPT1_P8 = 0 And BVAR_LPT1_P9 = 1 Then
            BVAR_LPT1_P6 = 1
            BVAR_LPT1_P7 = 1
            BVAR_LPT1_P8 = 0
            BVAR_LPT1_P9 = 0
        Else
            BVAR_LPT1_P6 = 1
            BVAR_LPT1_P7 = 1
            BVAR_LPT1_P8 = 0
            BVAR_LPT1_P9 = 0
        End If
    ElseIf IVAR_PASSODESTINO_EIXOY >= 0 And IVAR_PASSODESTINO_EIXOY < IVAR_PASSO_Y Then
        'subtrai passo e posição
        IVAR_PASSO_Y = IVAR_PASSO_Y - 1
        SGVAR_POSICAO_Y = IVAR_PASSO_Y * DCONST_MOVIMENTO_POR_PASSO
        BVAR_MOVE_EIXOY = True
        BVAR_LPT1_P14 = 0 'liga alimentacao eixo y
        If BVAR_LPT1_P6 = 1 And BVAR_LPT1_P7 = 1 And BVAR_LPT1_P8 = 0 And BVAR_LPT1_P9 = 0 Then
            BVAR_LPT1_P6 = 1
            BVAR_LPT1_P7 = 0
            BVAR_LPT1_P8 = 0
            BVAR_LPT1_P9 = 1
        ElseIf BVAR_LPT1_P6 = 0 And BVAR_LPT1_P7 = 1 And BVAR_LPT1_P8 = 1 And BVAR_LPT1_P9 = 0 Then
            BVAR_LPT1_P6 = 1
            BVAR_LPT1_P7 = 1
            BVAR_LPT1_P8 = 0
            BVAR_LPT1_P9 = 0
        ElseIf BVAR_LPT1_P6 = 0 And BVAR_LPT1_P7 = 0 And BVAR_LPT1_P8 = 1 And BVAR_LPT1_P9 = 1 Then
            BVAR_LPT1_P6 = 0
            BVAR_LPT1_P7 = 1
            BVAR_LPT1_P8 = 1
            BVAR_LPT1_P9 = 0
        ElseIf BVAR_LPT1_P6 = 1 And BVAR_LPT1_P7 = 0 And BVAR_LPT1_P8 = 0 And BVAR_LPT1_P9 = 1 Then
            BVAR_LPT1_P6 = 0
            BVAR_LPT1_P7 = 0
            BVAR_LPT1_P8 = 1
            BVAR_LPT1_P9 = 1
        Else
            BVAR_LPT1_P6 = 0
            BVAR_LPT1_P7 = 0
            BVAR_LPT1_P8 = 1
            BVAR_LPT1_P9 = 1
        End If
    Else
        BVAR_MOVE_EIXOY = False
        'se entrar aqui significa que o eixo já está na posição destino ou não deve ser mover, portanto deve desligar o motor
        BVAR_LPT1_P14 = 1 'desliga alimentacao eixo x
    End If
    'verifica se os 2 eixos chegaram ao fim de curso, caso contrario, envia dados para a porta
    If BVAR_MOVE_EIXOX = False And BVAR_MOVE_EIXOY = False Then
        Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Enabled = False
    Else
        'ENVIA DADOS ESCRITOS ACIMA PARA A PORTA PARALELA
        EscrevePortaControle 'envia dados para porta paralela (controle) primeira para ligar os motores
        EscrevePortaDados 'envia dados para porta paralela (dados) para controlar posição dos motores
    End If
    'Tela_Chapinha_D43.PIC_CHAP.PSet (IVAR_PASSO_X * 10 * DCONST_MOVIMENTO_POR_PASSO, IVAR_PASSO_Y * 10 * DCONST_MOVIMENTO_POR_PASSO), QBColor(2)
End Sub
Public Sub PosicionaMarcadorChapinha(ByVal X As Integer, ByVal Y As Integer)
    IVAR_PASSODESTINO_EIXOX = X
    IVAR_PASSODESTINO_EIXOY = Y
    If X <= ICONST_PASSO_MAXIMO_EIXO_X And _
       X > 0 And _
       Y <= ICONST_PASSO_MAXIMO_EIXO_Y And _
       Y > 0 Then
        Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Enabled = True
    ElseIf X = 0 And Y = 0 Then
        Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_HOME.Enabled = True
    End If
End Sub
Public Sub MoveUmPassoEixoX_Horario()
    IVAR_PASSODESTINO_EIXOX = IVAR_PASSO_X + 1
    If IVAR_PASSODESTINO_EIXOX <= ICONST_PASSO_MAXIMO_EIXO_X Then
        MoveEixos_MarcadorChapinha
    End If
End Sub
Public Sub MoveUmPassoEixoX_Antihorario()
    IVAR_PASSODESTINO_EIXOX = IVAR_PASSO_X - 1
    If IVAR_PASSODESTINO_EIXOX >= 0 Then
        MoveEixos_MarcadorChapinha
    End If
End Sub
Public Sub MoveUmPassoEixoY_Horario()
    IVAR_PASSODESTINO_EIXOY = IVAR_PASSO_Y + 1
    If IVAR_PASSODESTINO_EIXOY <= ICONST_PASSO_MAXIMO_EIXO_Y Then
        MoveEixos_MarcadorChapinha
    End If
End Sub
Public Sub MoveUmPassoEixoY_Antihorario()
    IVAR_PASSODESTINO_EIXOY = IVAR_PASSO_Y - 1
    If IVAR_PASSODESTINO_EIXOY >= 0 Then
        MoveEixos_MarcadorChapinha
    End If
End Sub
Public Sub MoveSemPararEixoX_Horario()
    IVAR_PASSODESTINO_EIXOX = ICONST_PASSO_MAXIMO_EIXO_X
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Enabled = True
End Sub
Public Sub MoveSemPararEixoX_Antihorario()
    IVAR_PASSODESTINO_EIXOX = 0
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Enabled = True
End Sub
Public Sub MoveSemPararEixoY_Horario()
    IVAR_PASSODESTINO_EIXOY = ICONST_DISTANCIA_MAXIMA_EIXO_Y
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Enabled = True
End Sub
Public Sub MoveSemPararEixoY_Antihorario()
    IVAR_PASSODESTINO_EIXOY = 0
    Tela_Posicao.TIMER_CHAPINHA_POSICIONAMENTO_EIXOS.Enabled = True
End Sub


'***************************************************************************
'              FUNCOES ESPECIFICAS PARA O MARCADOR DE CHAPINHAS
'***************************************************************************

Public Sub MarcadorChapinha(ByVal VETOR As Variant)
    Dim I, J, K, L, M As Integer
    Dim POSX, POSY, POSPONTOX, POSPONTOY, ULTIMOX, ULTIMOY As Integer
    Dim CARRO_ESQUERDA_PARA_DIREIRA As Boolean
    'inicializa variaveis
    VETOR_POSICIONAMENTO_MARCACAO = Array()
    SVAR_ARQUIVO = SVAR_CAMINHO_ARQUIVOS & "\7x7.car"
    CARRO_ESQUERDA_PARA_DIREIRA = True
    ULTIMOY = 0
    'carrega vetor
    For I = 0 To UBound(VETOR) 'cada texto do vetor
        'redimensiona a matriz e adiciona a matriz dos pontos que serão marcados de cada caracter do texto
        For J = 1 To Len(VETOR(I)(0)) 'cada caracter do texto
            'ao final de cada caracter, POSY permanece o mesmo pois o caractere ao lado
            'do texto tem a mesma altura, porém a POSX, a cada caracter, temos que adicionar
            'o comprimento de cada caracter mais o espaço entre eles até o final do texto
            POSX = VETOR(I)(1) + ((J - 1) * (7 * ICONST_PASSO_PONTO_CARACTER7x7)) + ((J - 1) * ICONST_ESPACO_ENTRE_CARACTER7x7)
            POSY = VETOR(I)(2)
            'pega mapa de pontos de cada caracter do texto
            For K = 1 To 7 'colunas dos 7 caracteres - POSX - cada ponto em X do caracter
                For L = 1 To 7 'linhas dos 7 caracteres - POSY - cada ponto em Y do caracter
                    If LeINI(SVAR_ARQUIVO, Mid(VETOR(I)(0), J, 1), "L" & K & "C" & L) = 1 Then 'ponto deve ser marcado
                        POSPONTOX = POSX + ((L - 1) * ICONST_PASSO_PONTO_CARACTER7x7) 'em passos
                        POSPONTOY = POSY + ((K - 1) * ICONST_PASSO_PONTO_CARACTER7x7)
                        'se existir será redimensionado o vetor de marcação
                        ReDim Preserve VETOR_POSICIONAMENTO_MARCACAO(UBound(VETOR_POSICIONAMENTO_MARCACAO) + 1)
                        VETOR_POSICIONAMENTO_MARCACAO(UBound(VETOR_POSICIONAMENTO_MARCACAO)) = Array(POSPONTOX, POSPONTOY)
                    End If
                Next L
            Next K
        Next J
    Next I
    'OptimizaVetorMarcacao ' ----> ESTA OPTIMIZACAO LEVA O DOBRO DO TEMPO !?!?!?!?
End Sub
Public Sub OptimizaVetorMarcacao()
    Dim VETOR, VETOR_XCOMY As Variant
    Dim X, Y, X2, Y2, TAM_VX, I As Integer, E2D, PY As Boolean
    X2 = 0
    Y2 = Fix(ICONST_DISTANCIA_MAXIMA_EIXO_Y / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR
    X = Fix(ICONST_DISTANCIA_MAXIMA_EIXO_X / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR
    Y = Y2
    E2D = True
    PY = True
    VETOR = Array()
    VETOR_XCOMY = Array()
    Do While True
    'UBound(VETOR) <> UBound(VETOR_POSICIONAMENTO_MARCACAO)
        'acha menor Y se for da E2D ou maior se D2E
        For I = 0 To UBound(VETOR_POSICIONAMENTO_MARCACAO)
            If (PY = True And VETOR_POSICIONAMENTO_MARCACAO(I)(1) < Y2 And VETOR_POSICIONAMENTO_MARCACAO(I)(1) < Y) Or (PY = False And VETOR_POSICIONAMENTO_MARCACAO(I)(1) > Y2 And VETOR_POSICIONAMENTO_MARCACAO(I)(1) < Y) Then
                Y = VETOR_POSICIONAMENTO_MARCACAO(I)(1)
            End If
        Next I
        If Y = Fix(ICONST_DISTANCIA_MAXIMA_EIXO_Y / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR Then
            Exit Do
        End If
        'acha os valores de X com o Y acima
        For I = 0 To UBound(VETOR_POSICIONAMENTO_MARCACAO)
            If VETOR_POSICIONAMENTO_MARCACAO(I)(1) = Y Then
                ReDim Preserve VETOR_XCOMY(UBound(VETOR_XCOMY) + 1)
                VETOR_XCOMY(UBound(VETOR_XCOMY)) = Array(VETOR_POSICIONAMENTO_MARCACAO(I)(0), VETOR_POSICIONAMENTO_MARCACAO(I)(1))
            End If
        Next I
        'acha menor X se for da E2D ou maior se D2E
        'e adiciona novos valores no novo vetor principal
        TAM_VX = 0
        If E2D = True Then
            Do While TAM_VX <> UBound(VETOR_XCOMY)
                'acha menor X
                For I = 0 To UBound(VETOR_XCOMY)
                    If VETOR_XCOMY(I)(0) < X And VETOR_XCOMY(I)(0) > X2 Then
                        X = VETOR_XCOMY(I)(0)
                    End If
                Next I
                ReDim Preserve VETOR(UBound(VETOR) + 1)
                VETOR(UBound(VETOR)) = Array(X, Y)
                X2 = X
                X = Fix(ICONST_DISTANCIA_MAXIMA_EIXO_X / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR
                TAM_VX = TAM_VX + 1
            Loop
        ElseIf E2D = False Then
            Do While TAM_VX <> UBound(VETOR_XCOMY)
                'acha menor X
                For I = 0 To UBound(VETOR_XCOMY)
                    If VETOR_XCOMY(I)(0) > X And VETOR_XCOMY(I)(0) < X2 Then
                        X = VETOR_XCOMY(I)(0)
                    End If
                Next I
                ReDim Preserve VETOR(UBound(VETOR) + 1)
                VETOR(UBound(VETOR)) = Array(X, Y)
                X2 = X
                X = 0
                TAM_VX = TAM_VX + 1
            Loop
        End If
        'muda direção da maquina e seta valor de x conforme direção
        If E2D = True Then
            E2D = False
            X = 0
            X2 = Fix(ICONST_DISTANCIA_MAXIMA_EIXO_X / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR
        Else
            E2D = True
            X = Fix(ICONST_DISTANCIA_MAXIMA_EIXO_X / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR
            X2 = 0
        End If
        PY = False
        VETOR_XCOMY = Array()
        Y2 = Y
        Y = Fix(ICONST_DISTANCIA_MAXIMA_EIXO_Y / SGCONST_PASSO_ROSCA_EIXO_MOTOR) * ICONST_NUMERO_PASSOS_MOTOR
    Loop
    'seta vetor de marcacao optimizado
    VETOR_POSICIONAMENTO_MARCACAO = VETOR
End Sub
