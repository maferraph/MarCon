Attribute VB_Name = "Modulo_Funcoes"
Option Explicit
'***************************************************************************
'                             VARIÁVEIS GLOBAIS
'***************************************************************************
Public SVAR_TEMP, SVAR_ARQUIVO, SVAR_ARQUIVO_TEXTO, SVAR_ARQUIVO_POSICAO, SVAR_CAMINHO_ARQUIVOS As String
Public VAR_INTEMP, IVAR_VETORMARCADOR As Integer
Public BVAR_SALVAR As Boolean
Public SVAR_MARCACAO_ATUAL As String

'***************************************************************************
'                          FUNÇOES DA API DO WINDOWS
'***************************************************************************
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Sub EscreveINI(ByVal Arquivo As String, ByVal Secao As String, ByVal Chave As String, ByVal VALOR As String)
    WritePrivateProfileString Secao, Chave, VALOR, Arquivo
End Sub
Public Function LeINI(ByVal Arquivo As String, ByVal Secao As String, ByVal Chave As String) As String
    Dim Ret As String * 256
    Dim RetLen As Long
    RetLen = GetPrivateProfileString(Secao, Chave, "NAOEXISTE", Ret, Len(Ret), Arquivo)
    LeINI = Left$(Ret, RetLen)
End Function

'***************************************************************************
'                              FUNÇÕES DIVERSAS
'***************************************************************************
Public Function Binario2Decimal(VALOR As String) As Long
    Dim VAR_DECIMAL As Long, I As Integer
    VAR_DECIMAL = 0
    For I = 1 To Len(VALOR)
        VAR_DECIMAL = VAR_DECIMAL + (Int(Mid(VALOR, Len(VALOR) - I + 1, 1)) * (2 ^ (I - 1)))
    Next I
    Binario2Decimal = VAR_DECIMAL
End Function
Public Function Decimal2Binario(VALOR As Long) As String
    Dim SVAR_BINARIO As String, LVAR_DIVIDENDO, LVAR_QUOCIENTE, LVAR_RESTO As Long
    LVAR_QUOCIENTE = 0
    LVAR_DIVIDENDO = VALOR
    SVAR_BINARIO = ""
    If VALOR > 1 Then
        Do While LVAR_QUOCIENTE <> 1
            LVAR_QUOCIENTE = Fix(LVAR_DIVIDENDO / 2)
            LVAR_RESTO = LVAR_DIVIDENDO - (LVAR_QUOCIENTE * 2)
            SVAR_BINARIO = SVAR_BINARIO & LVAR_RESTO
            LVAR_DIVIDENDO = LVAR_QUOCIENTE
        Loop
        SVAR_BINARIO = SVAR_BINARIO & "1" 'ultimo resultado do quociente
        Decimal2Binario = StrReverse(SVAR_BINARIO)
    ElseIf VALOR = 0 Then
        Decimal2Binario = "0"
    ElseIf VALOR = 1 Then
        Decimal2Binario = "1"
    Else
        Decimal2Binario = ""
    End If
End Function
