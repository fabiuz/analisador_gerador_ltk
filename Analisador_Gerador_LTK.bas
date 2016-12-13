#PBFORMS CREATED V2.00
'------------------------------------------------------------------------------
' The first line in this file is a PB/Forms metastatement.
' It should ALWAYS be the first line of the file. Other
' PB/Forms metastatements are placed at the beginning and
' end of "Named Blocks" of code that should be edited
' with PBForms only. Do not manually edit or delete these
' metastatements or PB/Forms will not be able to reread
' the file correctly.  See the PB/Forms documentation for
' more information.
' Named blocks begin like this:    #PBFORMS BEGIN ...
' Named blocks end like this:      #PBFORMS END ...
' Other PB/Forms metastatements such as:
'     #PBFORMS DECLARATIONS
' are used by PB/Forms to insert additional code.
' Feel free to make changes anywhere else in the file.
'------------------------------------------------------------------------------

' Na vers�o 5.0:
' Foram adicionados os jogos: INTRALOT_MINAS_5, INTRALOT_LOTO_MINAS, INTRALOT_KENO_MINAS
'


#COMPILE EXE
#DIM ALL
#OPTIMIZE CODE ON
#OPTIMIZE SIZE

'------------------------------------------------------------------------------
'   ** Includes **
'------------------------------------------------------------------------------
#PBFORMS BEGIN INCLUDES
#RESOURCE "Analisador_Gerador_LTK.pbr"
#INCLUDE ONCE "WIN32API.INC"
#INCLUDE ONCE "COMMCTRL.INC"
#INCLUDE ONCE "PBForms.INC"
#PBFORMS END INCLUDES
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'   ** Constants **
'------------------------------------------------------------------------------
#PBFORMS BEGIN CONSTANTS
%IDD_DIALOG1                  =  101
%IDC_FRAME1                   = 1001
%IDC_CMB_JOGO_TIPO_TODAS      = 1002
%IDC_LABEL1                   = 1003
%IDC_LABEL2                   = 1004
%IDC_CMB_APOSTAR_COM_TODAS    = 1005
%IDC_LABEL3                   = 1006
%IDC_TEXTBOX1                 = 1007    '*
%IDC_TXT_PASTA_DESTINO_TODAS  = 1008
%IDC_BTN_GERAR_TODAS          = 1009
%IDC_CHK_ATUALIZAR_TODAS      = 1010
%IDC_LBL_LOG_TODAS            = 1011
%IDC_FRAME2                   = 1012
%IDC_LABEL4                   = 1013
%IDC_TXT_ARQUIVO_ORIGEM       = 1014
%IDC_LABEL5                   = 1015
%IDC_TXT_PASTA_DESTINO_INSERT = 1016
%IDC_CHK_STATUS_INSERT        = 1017
%IDC_LOG_INSERT               = 1018
%IDC_BUTTON1                  = 1019    '*
%IDC_BTN_GERAR_INSERT         = 1020
%IDC_BTN_PARAR_TODAS          = 1021
%IDC_BTN_PARAR_INSERT         = 1022
%IDC_BTN_PAUSAR_TODAS         = 1023    '*
%IDC_CHK_PAUSAR_TODAS         = 1024
%IDC_CHK_PAUSAR_INSERT        = 1025
%IDC_FRAME3                   = 1026
%IDC_LABEL6                   = 1027
%IDC_LABEL7                   = 1030
%IDC_STATUS_LOTERIA           = 1031
%IDC_BTN_GERAR_LOTERIA        = 1032
%IDC_LOG_LOTERIA              = 1033
%IDC_BTN_PARAR_LOTERIA        = 1034
%IDC_CHK_PAUSAR_LOTERIA       = 1035
%IDC_COMMIT                   = 1036
%TXT_TESTE                    = 1037    '*
%BTN_TESTE                    = 1038    '*
%LST_TESTE                    = 1039    '*
%LBL_TESTE                    = 1040
%BTN_PARAR_TESTE              = 1041
%BTN_PAUSAR_TESTE             = 1042
%IDC_FRAME4                   = 1043    '*
%IDC_CHECKBOX1                = 1044    '*
%IDC_LOG                      = 1045
%IDC_LABEL8                   = 1046
%TXT_DIRETORIO_SQL_GERADO     = 1029
%TXT_DIRETORIO_CSV_PROCESSADO = 1047
%TXT_ARQUIVO_CSV              = 1028
#PBFORMS END CONSTANTS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'   ** Declarations **
'------------------------------------------------------------------------------
DECLARE CALLBACK FUNCTION Janela_Principal_CBK()
DECLARE FUNCTION Exibir_Janela_Principal(BYVAL hParent AS DWORD) AS LONG
#PBFORMS DECLARATIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' ** DECLARA��ES GLOBAIS **
'------------------------------------------------------------------------------
GLOBAL hJanela_Principal AS DWORD
GLOBAL uAtualizarTodas AS LONG
GLOBAL uSairTodas AS LONG
GLOBAL uPausarTodas AS LONG


GLOBAL uPausarInsert AS LONG
GLOBAL uPararInsert AS LONG
GLOBAL uAtualizarInsert AS LONG


GLOBAL uPausarTeste AS LONG
GLOBAL uPararTeste AS LONG

GLOBAL uCommit AS LONG

ENUM TIPO_DE_JOGO SINGULAR
    QUINA_COM_5_NUMEROS
    QUINA_COM_6_NUMEROS
    QUINA_COM_7_NUMEROS
    MEGASENA_COM_6_NUMEROS
    MEGASENA_COM_7_NUMEROS
    MEGASENA_COM_8_NUMEROS
    MEGASENA_COM_9_NUMEROS
    MEGASENA_COM_10_NUMEROS
    MEGASENA_COM_11_NUMEROS
    MEGASENA_COM_12_NUMEROS
    MEGASENA_COM_13_NUMEROS
    MEGASENA_COM_14_NUMEROS
    MEGASENA_COM_15_NUMEROS
    LOTOFACIL_COM_15_NUMEROS
    LOTOFACIL_COM_16_NUMEROS
    LOTOFACIL_COM_17_NUMEROS
    LOTOFACIL_COM_18_NUMEROS
END ENUM



FUNCTION Carregar_Lista(BYVAL strJogo AS STRING) AS LONG
    COMBOBOX RESET hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS

    strJogo = UCASE$(strJogo)
    SELECT CASE strJogo
        CASE "QUINA"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "5 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "6 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "7 n�meros"

        CASE "MEGASENA"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "6 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "7 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "8 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "9 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "10 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "11 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "12 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "14 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "15 n�meros"

        CASE "LOTOFACIL"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "15 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "16 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "17 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "18 n�meros"

        CASE "LOTOMANIA"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "50 n�meros

        CASE "DUPLASENA"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "6 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "7 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "8 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "9 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "10 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "11 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "12 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "14 n�meros"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "15 n�meros"
    END SELECT

    COMBOBOX SELECT hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, 1

END FUNCTION

'
' Carrega as caixas de combina��o 'TIPO DE JOGO' e 'APOSTAR COM'
'
SUB Carregar_Lista_Inicial()
    COMBOBOX RESET hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS
    COMBOBOX RESET hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS
    '
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "QUINA"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "MEGASENA"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "LOTOFACIL"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "LOTOMANIA"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "DUPLASENA"

    '
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "5 n�meros"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "6 n�meros"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "7 n�meros"

    '
    COMBOBOX SELECT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, 1
    COMBOBOX SELECT hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, 1
END SUB


'
'   Esta fun��o verifica se o tipo de jogo e a quantidade de n�meros apostados
'   � permitido na loteria
'   Retorna -1 se h� algo errado
'
FUNCTION Validar_Tipo_de_Jogo(BYVAL strJogo AS STRING, BYVAL uBolasApostadas AS LONG, BYREF strErro AS STRING) AS LONG
    ' Vamos verificar se o jogo j� foi implementado
    strJogo = UCASE$(strJogo)
    IF strJogo <> "QUINA" AND strJogo <> "MEGASENA" AND strJogo <> "LOTOFACIL" THEN
        strErro = "O jogo solicitado: '" & strJogo & "' ainda n�o foi implementado."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Verificar se a quantidade de bolas coincide para cada tipo de jogo, por exemplo,
    ' na quina, voc� pode selecionar, 5, 6 ou 7 n�meros
    SELECT CASE strJogo
        CASE "QUINA"
            ' Se a quantidade de bolas for diferente de 5, 6 ou 7
            IF ISFALSE(uBolasApostadas >= 5 AND uBolasApostadas <= 7) THEN
                strErro = "Para o jogo 'QUINA' voc� pode apostar com 5, 6 ou 7 n�meros, entretanto, voc� selecionou " & _
                          " um outro n�mero."
                FUNCTION = -1
                EXIT FUNCTION
            END IF
        CASE "MEGASENA"
            ' Se a quantidade de bolas for diferente de 6 A 15
            IF ISFALSE(uBolasApostadas >= 6 AND uBolasApostadas <= 15) THEN
                strErro = "Para o jogo 'MEGASENA' voc� pode apostar de 6 a 15 n�meros, entretanto, voc� selecionou " & _
                          " um outro n�mero."
                FUNCTION = -1
                EXIT FUNCTION
            END IF

        CASE "LOTOFACIL"
            ' Se a quantidade de bolas for diferente de 6 A 15
            IF ISFALSE(uBolasApostadas >= 15 AND uBolasApostadas <= 18) THEN
                strErro = "Para o jogo 'LOTOFACIL' voc� pode apostar de 15 a 18 n�meros, entretanto, voc� selecionou " & _
                          " um outro n�mero."
                FUNCTION = -1
                EXIT FUNCTION
            END IF

        CASE "DUPLASENA"
            ' Se a quantidade de bolas apostador for diferente de 6 a 15
            IF ISFALSE(uBolasApostadas >= 6 AND uBolasApostadas <= 15) THEN
                strErro = "Para o jogo 'DuplaSena' voc� pode apostar de 6 a 15 n�meros, entretanto, voc� selecionou " & _
                          " um outro n�mero."
                FUNCTION = -1
                EXIT FUNCTION
            END IF

        CASE "LOTOMANIA"
            ' Se a quantidade for diferente de 50
            IF uBolasApostadas <> 50 THEN
                strErro = "Para o jogo 'LOTOMANIA' voc� deve apostar 50 n�meros."
                FUNCTION = -1
                EXIT FUNCTION
            END IF
        END SELECT
    '
    FUNCTION = 0
END FUNCTION

'
'   Esta fun��o gera o nome do arquivo de acordo com o tipo do jogo e quantidade de n�meros apostados
'   A fun��o retorna um string vazio se algo ocorreu de errado
'
FUNCTION Gerar_Nome_do_Arquivo(BYVAL strPastaDestino AS STRING, BYVAL strJogo AS STRING, _
                                BYVAL uBolasApostadas AS LONG, BYREF strErro AS STRING) AS STRING
    ' Vamos verificar se a pasta existe
    IF DIR$(strPastaDestino, %SUBDIR) = "" THEN
        strErro = "A pasta de destino especificada n�o existe: '" & strPastaDestino & "'."
        FUNCTION = ""
        EXIT FUNCTION
    END IF

    ' Vamos verificar se a pasta tem o �ltimo caractere � direita, o caractere '\'
    IF RIGHT$(strPastaDestino, 1) <> "\" THEN
        strPastaDestino &= "\"
    END IF

    ' Vamos o nome do arquivo com a nomenclatura abaixo:
    ' <JOGO>_COM_<uBolasApostadas>_N�MEROS_DATA_HORA
    DIM objDataHora AS IPOWERTIME
    LET objDataHora = CLASS "POWERTIME"
    IF ISNOTHING(objDataHora) THEN
        strErro = "N�o foi poss�vel instanciar a classe 'POWERTIME', interface 'IPOWERTIME'."
        FUNCTION = ""
        EXIT FUNCTION
    END IF

    DIM strArquivoDestino AS STRING
    strArquivoDestino = strJogo & "_COM_" & FORMAT$(uBolasApostadas) & "_N�MEROS_"

    objDataHora.Now()
    strArquivoDestino &= FORMAT$(objDataHora.Year(), "0000") & "_" & FORMAT$(objDataHora.Month(), "00") & "_"
    strArquivoDestino &= FORMAT$(objDataHora.Day(), "00") & "_" & FORMAT$(objDataHora.Hour(), "00") & "-"
    strArquivoDestino &= FORMAT$(objDataHora.Minute(), "00") & "-" & FORMAT$(objDataHora.Second(), "00")
    strArquivoDestino = strPastaDestino & strArquivoDestino & ".txt"

    FUNCTION = strArquivoDestino
END FUNCTION

FUNCTION Carregar_Arranjo(BYVAL strJogo AS STRING, BYREF strBolas() AS STRING, BYREF strErro AS STRING) AS LONG
    SELECT CASE strJogo
        CASE "QUINA"
            REDIM strBolas(1 TO 80)
        CASE "MEGASENA"
            REDIM strBolas(1 TO 60)
        CASE "LOTOFACIL"
            REDIM strBolas(1 TO 25)
        CASE "DUPLASENA"
            REDIM strBolas(1 TO 50)
        CASE "LOTOMANIA"
            REDIM strBolas(0 TO 99)

        ' **************************
        ' ACRESCENTADO NA VERS�O 5.0
        ' **************************
        ' INICIO
        CASE "INTRALOT_MINAS_5"
            REDIM strBolas(1 TO 34)

        CASE "INTRALOT_LOTO_MINAS"
            REDIM strBolas(1 TO 38)

        CASE "INTRALOT_KENO_MINAS"
            REDIM strBolas(1 TO 80)
        ' FIM
        ' **************************


    END SELECT


    DIM iA AS LONG
    DIM uLimiteInferior AS LONG
    DIM uLimiteSuperior AS LONG
    '
    ERRCLEAR
    uLimiteInferior = LBOUND(strBolas, 1)
    uLimiteSuperior = UBOUND(strBolas, 1)

    FOR iA = uLimiteInferior TO uLimiteSuperior
        strBolas(iA) = FORMAT$(iA, "00")
    NEXT iA

    IF ERR <> 0 THEN
        strErro = "Erro: '" & FORMAT$(ERR) & "', " & ERROR$(ERR)
        FUNCTION = -1
    ELSE
        FUNCTION = 0
    END IF
END FUNCTION


FUNCTION Gerar_Todas_Combinacoes ALIAS "Gerar_Todas_Combinacoes" (BYVAL strJogo AS STRING, _
    BYVAL uBolasApostadas AS LONG, _       ' Quantidade de bolas apostadas
    BYVAL strPastaDestino AS STRING, _
     BYREF strErro AS STRING) EXPORT AS LONG

     ' Validar dados informados pelo usu�rio
     ' Se h� erro, retorna -1
     IF Validar_Tipo_de_Jogo(strJogo, uBolasApostadas, strErro) = -1 THEN
         FUNCTION = -1
         EXIT FUNCTION
     END IF

     ' Vamos gerar nome do arquivo
     DIM strArquivoDestino AS STRING
     strArquivoDestino = Gerar_Nome_do_Arquivo(strPastaDestino, strJogo, uBolasApostadas, strErro)
     ' Se n�o conseguir criar o nome do arquivo, sair ent�o
     IF strArquivoDestino = "" THEN
         FUNCTION = -1
         EXIT FUNCTION
     END IF

    ' Vamos tentar abrir o arquivo, se n�o conseguirmos sair da fun��o
    DIM uArquivoDestino AS LONG
    uArquivoDestino = FREEFILE
    OPEN strArquivoDestino FOR OUTPUT AS uArquivoDestino
    IF ERR <> 0 THEN
        strErro = "Erro: " & ERROR$(ERR)
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos carregar o arranjo que criamos
    DIM strBolas() AS STRING
    IF Carregar_Arranjo(strJogo, strBolas(), strErro) = -1 THEN
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    IF Gerar_Combinacao_Por_Jogo(strJogo, uBolasApostadas, strBolas(), uArquivoDestino, strErro) = -1 THEN
        EXIT FUNCTION
    END IF


END FUNCTION







FUNCTION Gerar_Combinacao_Por_Jogo(BYVAL strJogo AS STRING, BYVAL uBolasApostadas AS LONG, _
    BYREF strBolas() AS STRING, BYVAL uArquivoDestino AS LONG, BYREF strErro AS STRING) AS LONG
        ' Vari�veis serve para guardar o tipo do jogo
    DIM uJogoTipo AS LONG
    SELECT CASE strJogo
        CASE "QUINA"
        '
            SELECT CASE uBolasApostadas
                CASE 5
                    uJogoTipo = %QUINA_COM_5_NUMEROS
                CASE 6
                    uJogoTipo = %QUINA_COM_6_NUMEROS
                CASE 7
                    uJogoTipo = %QUINA_COM_7_NUMEROS
            END SELECT
        '
        CASE "MEGASENA"
        '
            SELECT CASE uBolasApostadas
                CASE 6
                    uJogoTipo = %MEGASENA_COM_6_NUMEROS
                CASE 7
                    uJogoTipo = %MEGASENA_COM_7_NUMEROS
                CASE 8
                    uJogoTipo = %MEGASENA_COM_8_NUMEROS
                CASE 9
                    uJogoTipo = %MEGASENA_COM_9_NUMEROS
                CASE 10
                    uJogoTipo = %MEGASENA_COM_10_NUMEROS
                CASE 11
                    uJogoTipo = %MEGASENA_COM_11_NUMEROS
                CASE 12
                    uJogoTipo = %MEGASENA_COM_12_NUMEROS
                CASE 13
                    uJogoTipo = %MEGASENA_COM_13_NUMEROS
                CASE 14
                    uJogoTipo = %MEGASENA_COM_14_NUMEROS
                CASE 15
                    uJogoTipo = %MEGASENA_COM_15_NUMEROS
            END SELECT
        '
        CASE "LOTOFACIL"
            SELECT CASE uBolasApostadas
                CASE 15
                    uJogoTipo = %LOTOFACIL_COM_15_NUMEROS
                CASE 16
                    uJogoTipo = %LOTOFACIL_COM_16_NUMEROS
                CASE 17
                    uJogoTipo = %LOTOFACIL_COM_17_NUMEROS
                CASE 18
                    uJogoTipo = %LOTOFACIL_COM_18_NUMEROS
            END SELECT
    END SELECT


    REGISTER uColuna1 AS LONG, uColuna2 AS LONG, uColuna3 AS LONG, uColuna4 AS LONG
    REGISTER uColuna5 AS LONG, uColuna6 AS LONG, uColuna7 AS LONG, uColuna8 AS LONG
    REGISTER uColuna9 AS LONG, uColuna10 AS LONG, uColuna11 AS LONG, uColuna12 AS LONG
    REGISTER uColuna13 AS LONG, uColuna14 AS LONG, uColuna15 AS LONG

    ' Para os jogos 'MEGASENA' e 'LOTOFACIL', o limite inferior � sempre '1' (um)
    ' entretanto, para o jogo, 'LOTOMANIA', o limite inferior � sempre '0' (zero)
    ' ent�o, devemos definir uColuna1 como o limite inferior do jogo solicitado
    DIM uLimiteInferior AS LONG
    DIM uLimiteSuperior AS LONG
    uLimiteInferior = LBOUND(strBolas(), 1)
    uLimiteSuperior = UBOUND(strBolas(), 1)


    DIM strTexto AS STRING
    strTexto = ""

    DIM uContador AS DWORD
    uContador = 0

    ' Para otimizar cada jogo, irei fazer o loop separado para cada tipo de jogo
    strErro = ""
    SELECT CASE uJogoTipo
        CASE %QUINA_COM_5_NUMEROS
            GOSUB GERAR_QUINA_COM_5_NUMEROS
        '
        CASE %QUINA_COM_6_NUMEROS
            GOSUB GERAR_QUINA_COM_6_NUMEROS
        '
        CASE %QUINA_COM_7_NUMEROS
            GOSUB GERAR_QUINA_COM_7_NUMEROS
        '
        CASE %MEGASENA_COM_6_NUMEROS
            GOSUB GERAR_MEGASENA_COM_6_NUMEROS
        '
        CASE %MEGASENA_COM_7_NUMEROS
            GOSUB GERAR_MEGASENA_COM_7_NUMEROS
        '
        CASE %MEGASENA_COM_8_NUMEROS
            GOSUB GERAR_MEGASENA_COM_8_NUMEROS
        '
        CASE %MEGASENA_COM_9_NUMEROS
            GOSUB GERAR_MEGASENA_COM_9_NUMEROS
        '
        CASE %MEGASENA_COM_10_NUMEROS
            GOSUB GERAR_MEGASENA_COM_10_NUMEROS
        '
        CASE %MEGASENA_COM_11_NUMEROS
            GOSUB GERAR_MEGASENA_COM_11_NUMEROS
        '
        CASE %MEGASENA_COM_12_NUMEROS
            GOSUB GERAR_MEGASENA_COM_12_NUMEROS
        '
        CASE %MEGASENA_COM_13_NUMEROS
            GOSUB GERAR_MEGASENA_COM_13_NUMEROS
        '
        CASE %MEGASENA_COM_14_NUMEROS
            GOSUB GERAR_MEGASENA_COM_14_NUMEROS
        '
        CASE %MEGASENA_COM_15_NUMEROS
            GOSUB GERAR_MEGASENA_COM_15_NUMEROS
    END SELECT

    uSairTodas = 0


    IF strErro <> "" THEN
        MSGBOX strErro, %MB_ICONERROR
    ELSE
        MSGBOX "Executado com sucesso!!!"
    END IF

    EXIT FUNCTION


GERAR_QUINA_COM_5_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        DIALOG DOEVENTS
                        strTexto = "QUINA_COM_5_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                        strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5)
                        PRINT #uArquivoDestino, strTexto
                        IF ERR <> 0 THEN
                            strErro = "Erro: " & ERROR$(ERR)
                            FUNCTION = -1
                            RETURN
                        END IF
                        uContador += 1
                        IF uAtualizarTodas = 1 THEN
                            strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                            CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                        END IF

                        ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                        ' para podermos sair do loop
                        IF uSairTodas = 1 THEN
                            strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                            CLOSE
                            RETURN
                        END IF

                        ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                        ' ou o usu�rio clicou no bot�o 'Parar'
                        DO WHILE uPausarTodas = 1
                            DIALOG DOEVENTS
                        LOOP
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_QUINA_COM_6_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            DIALOG DOEVENTS
                            strTexto = "QUINA_COM_6_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                            strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6)
                            PRINT #uArquivoDestino, strTexto
                            IF ERR <> 0 THEN
                                strErro = "Erro: " & ERROR$(ERR)
                                FUNCTION = -1
                                RETURN
                            END IF
                            uContador += 1
                            IF uAtualizarTodas = 1 THEN
                                strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                            END IF

                            ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                            ' para podermos sair do loop
                            IF uSairTodas = 1 THEN
                                strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                CLOSE
                                RETURN
                            END IF

                            ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                            ' ou o usu�rio clicou no bot�o 'Parar'
                            DO WHILE uPausarTodas = 1
                                DIALOG DOEVENTS
                            LOOP
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_QUINA_COM_7_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                DIALOG DOEVENTS
                                strTexto = "QUINA_COM_7_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7)
                                PRINT #uArquivoDestino, strTexto
                                IF ERR <> 0 THEN
                                    strErro = "Erro: " & ERROR$(ERR)
                                    FUNCTION = -1
                                    RETURN
                                END IF
                                uContador += 1
                                IF uAtualizarTodas = 1 THEN
                                    strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                    CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                END IF

                                ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                ' para podermos sair do loop
                                IF uSairTodas = 1 THEN
                                    strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                    CLOSE
                                    RETURN
                                END IF

                                ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                ' ou o usu�rio clicou no bot�o 'Parar'
                                DO WHILE uPausarTodas = 1
                                    DIALOG DOEVENTS
                                LOOP
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_6_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            DIALOG DOEVENTS
                            strTexto = "MEGASENA_COM_6_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                            strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6)
                            PRINT #uArquivoDestino, strTexto
                            IF ERR <> 0 THEN
                                strErro = "Erro: " & ERROR$(ERR)
                                FUNCTION = -1
                                RETURN
                            END IF
                            uContador += 1
                            IF uAtualizarTodas = 1 THEN
                                strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                            END IF

                            ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                            ' para podermos sair do loop
                            IF uSairTodas = 1 THEN
                                strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                CLOSE
                                RETURN
                            END IF

                            ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                            ' ou o usu�rio clicou no bot�o 'Parar'
                            DO WHILE uPausarTodas = 1
                                DIALOG DOEVENTS
                            LOOP
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_7_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                DIALOG DOEVENTS
                                strTexto = "MEGASENA_COM_7_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7)
                                PRINT #uArquivoDestino, strTexto
                                IF ERR <> 0 THEN
                                    strErro = "Erro: " & ERROR$(ERR)
                                    FUNCTION = -1
                                    RETURN
                                END IF
                                uContador += 1
                                IF uAtualizarTodas = 1 THEN
                                    strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                    CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                END IF

                                                                    ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                    ' para podermos sair do loop
                                    IF uSairTodas = 1 THEN
                                        strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                        CLOSE
                                        RETURN
                                    END IF


                                ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                ' ou o usu�rio clicou no bot�o 'Parar'
                                DO WHILE uPausarTodas = 1
                                    DIALOG DOEVENTS
                                LOOP
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_8_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    DIALOG DOEVENTS
                                    strTexto = "MEGASENA_COM_8_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                    strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                    strTexto = strTexto & strBolas(uColuna8)
                                    PRINT #uArquivoDestino, strTexto
                                    IF ERR <> 0 THEN
                                        strErro = "Erro: " & ERROR$(ERR)
                                        FUNCTION = -1
                                        RETURN
                                    END IF
                                    uContador += 1
                                    IF uAtualizarTodas = 1 THEN
                                        strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                        CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                    END IF

                                    ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                    ' para podermos sair do loop
                                    IF uSairTodas = 1 THEN
                                        strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                        CLOSE
                                        RETURN
                                    END IF

                                    ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                    ' ou o usu�rio clicou no bot�o 'Parar'
                                    DO WHILE uPausarTodas = 1
                                        DIALOG DOEVENTS
                                    LOOP
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_9_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    FOR uColuna9 = uColuna8 + 1 TO uLimiteSuperior
                                        DIALOG DOEVENTS
                                        strTexto = "MEGASENA_COM_9_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                        strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                        strTexto = strTexto & strBolas(uColuna8) & ";" & strBolas(uColuna9)
                                        PRINT #uArquivoDestino, strTexto
                                        IF ERR <> 0 THEN
                                            strErro = "Erro: " & ERROR$(ERR)
                                            FUNCTION = -1
                                            RETURN
                                        END IF
                                        uContador += 1
                                        IF uAtualizarTodas = 1 THEN
                                            strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                            CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                        END IF
                                        ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                        ' para podermos sair do loop
                                        IF uSairTodas = 1 THEN
                                            strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                            CLOSE
                                            RETURN
                                        END IF

                                        ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                        ' ou o usu�rio clicou no bot�o 'Parar'
                                        DO WHILE uPausarTodas = 1
                                            DIALOG DOEVENTS
                                        LOOP
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_10_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    FOR uColuna9 = uColuna8 + 1 TO uLimiteSuperior
                                        FOR uColuna10 = uColuna9 + 1 TO uLimiteSuperior
                                            DIALOG DOEVENTS
                                            strTexto = "MEGASENA_COM_10_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                            strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                            strTexto = strTexto & strBolas(uColuna8) & ";" & strBolas(uColuna9) & ";" & strBolas(uColuna10)
                                            PRINT #uArquivoDestino, strTexto
                                            IF ERR <> 0 THEN
                                                strErro = "Erro: " & ERROR$(ERR)
                                                FUNCTION = -1
                                                RETURN
                                            END IF
                                            uContador += 1
                                            IF uAtualizarTodas = 1 THEN
                                                strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                                CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                            END IF
                                            ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                            ' para podermos sair do loop
                                            IF uSairTodas = 1 THEN
                                                strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                                RETURN
                                            END IF

                                            ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                            ' ou o usu�rio clicou no bot�o 'Parar'
                                            DO WHILE uPausarTodas = 1
                                                DIALOG DOEVENTS
                                            LOOP
                                        NEXT uColuna10
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_11_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    FOR uColuna9 = uColuna8 + 1 TO uLimiteSuperior
                                        FOR uColuna10 = uColuna9 + 1 TO uLimiteSuperior
                                            FOR uColuna11 = uColuna10 + 1 TO uLimiteSuperior
                                                DIALOG DOEVENTS
                                                strTexto = "MEGASENA_COM_11_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                                strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                                strTexto = strTexto & strBolas(uColuna8) & ";" & strBolas(uColuna9) & ";" & strBolas(uColuna10) & ";" & strBolas(uColuna11) & ";"
                                                PRINT #uArquivoDestino, strTexto
                                                IF ERR <> 0 THEN
                                                    strErro = "Erro: " & ERROR$(ERR)
                                                    FUNCTION = -1
                                                    RETURN
                                                END IF
                                                uContador += 1
                                                IF uAtualizarTodas = 1 THEN
                                                    strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                                    CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                                END IF

                                                ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                                ' para podermos sair do loop
                                                IF uSairTodas = 1 THEN
                                                    strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                                    RETURN
                                                END IF

                                                ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                                ' ou o usu�rio clicou no bot�o 'Parar'
                                                DO WHILE uPausarTodas = 1
                                                    DIALOG DOEVENTS
                                                LOOP
                                            NEXT uColuna11
                                        NEXT uColuna10
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_12_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    FOR uColuna9 = uColuna8 + 1 TO uLimiteSuperior
                                        FOR uColuna10 = uColuna9 + 1 TO uLimiteSuperior
                                            FOR uColuna11 = uColuna10 + 1 TO uLimiteSuperior
                                                FOR uColuna12 = uColuna11 + 1 TO uLimiteSuperior
                                                    DIALOG DOEVENTS
                                                    strTexto = "MEGASENA_COM_12_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                                    strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                                    strTexto = strTexto & strBolas(uColuna8) & ";" & strBolas(uColuna9) & ";" & strBolas(uColuna10) & ";" & strBolas(uColuna11) & ";"
                                                    strTexto = strTexto & strBolas(uColuna12)
                                                    PRINT #uArquivoDestino, strTexto
                                                    IF ERR <> 0 THEN
                                                        strErro = "Erro: " & ERROR$(ERR)
                                                        FUNCTION = -1
                                                        RETURN
                                                    END IF
                                                    uContador += 1
                                                    IF uAtualizarTodas = 1 THEN
                                                        strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                                        CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                                    END IF

                                                    ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                                    ' para podermos sair do loop
                                                    IF uSairTodas = 1 THEN
                                                        strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                                        RETURN
                                                    END IF

                                                    ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                                    ' ou o usu�rio clicou no bot�o 'Parar'
                                                    DO WHILE uPausarTodas = 1
                                                        DIALOG DOEVENTS
                                                    LOOP
                                                NEXT uColuna12
                                            NEXT uColuna11
                                        NEXT uColuna10
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_13_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    FOR uColuna9 = uColuna8 + 1 TO uLimiteSuperior
                                        FOR uColuna10 = uColuna9 + 1 TO uLimiteSuperior
                                            FOR uColuna11 = uColuna10 + 1 TO uLimiteSuperior
                                                FOR uColuna12 = uColuna11 + 1 TO uLimiteSuperior
                                                    FOR uColuna13 = uColuna12 + 1 TO uLimiteSuperior
                                                        DIALOG DOEVENTS
                                                        strTexto = "MEGASENA_COM_13_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                                        strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                                        strTexto = strTexto & strBolas(uColuna8) & ";" & strBolas(uColuna9) & ";" & strBolas(uColuna10) & ";" & strBolas(uColuna11) & ";"
                                                        strTexto = strTexto & strBolas(uColuna12) & ";" & strBolas(uColuna13)
                                                        PRINT #uArquivoDestino, strTexto
                                                        IF ERR <> 0 THEN
                                                            strErro = "Erro: " & ERROR$(ERR)
                                                            FUNCTION = -1
                                                            RETURN
                                                        END IF
                                                        uContador += 1
                                                        IF uAtualizarTodas = 1 THEN
                                                            strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                                            CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                                        END IF

                                                        ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                                        ' para podermos sair do loop
                                                        IF uSairTodas = 1 THEN
                                                            strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                                            RETURN
                                                        END IF

                                                        ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                                        ' ou o usu�rio clicou no bot�o 'Parar'
                                                        DO WHILE uPausarTodas = 1
                                                            DIALOG DOEVENTS
                                                        LOOP
                                                    NEXT uColuna13
                                                NEXT uColuna12
                                            NEXT uColuna11
                                        NEXT uColuna10
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_14_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    FOR uColuna9 = uColuna8 + 1 TO uLimiteSuperior
                                        FOR uColuna10 = uColuna9 + 1 TO uLimiteSuperior
                                            FOR uColuna11 = uColuna10 + 1 TO uLimiteSuperior
                                                FOR uColuna12 = uColuna11 + 1 TO uLimiteSuperior
                                                    FOR uColuna13 = uColuna12 + 1 TO uLimiteSuperior
                                                        FOR uColuna14 = uColuna13 + 1 TO uLimiteSuperior
                                                            DIALOG DOEVENTS
                                                            strTexto = "MEGASENA_COM_14_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                                            strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                                            strTexto = strTexto & strBolas(uColuna8) & ";" & strBolas(uColuna9) & ";" & strBolas(uColuna10) & ";" & strBolas(uColuna11) & ";"
                                                            strTexto = strTexto & strBolas(uColuna12) & ";" & strBolas(uColuna13) & ";" & strBolas(uColuna14)
                                                            PRINT #uArquivoDestino, strTexto
                                                            IF ERR <> 0 THEN
                                                                strErro = "Erro: " & ERROR$(ERR)
                                                                FUNCTION = -1
                                                                RETURN
                                                            END IF
                                                            uContador += 1
                                                            IF uAtualizarTodas = 1 THEN
                                                                strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                                                CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                                            END IF
                                                            ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                                            ' para podermos sair do loop
                                                            IF uSairTodas = 1 THEN
                                                                strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                                                RETURN
                                                            END IF

                                                            ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                                            ' ou o usu�rio clicou no bot�o 'Parar'
                                                            DO WHILE uPausarTodas = 1
                                                                DIALOG DOEVENTS
                                                            LOOP
                                                        NEXT uColuna14
                                                    NEXT uColuna13
                                                NEXT uColuna12
                                            NEXT uColuna11
                                        NEXT uColuna10
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN

GERAR_MEGASENA_COM_15_NUMEROS:
    FOR uColuna1 = uLimiteInferior TO uLimiteSuperior
        FOR uColuna2 = uColuna1 + 1 TO uLimiteSuperior
            FOR uColuna3 = uColuna2 + 1 TO uLimiteSuperior
                FOR uColuna4 = uColuna3 + 1 TO uLimiteSuperior
                    FOR uColuna5 = uColuna4 + 1 TO uLimiteSuperior
                        FOR uColuna6 = uColuna5 + 1 TO uLimiteSuperior
                            FOR uColuna7 = uColuna6 + 1 TO uLimiteSuperior
                                FOR uColuna8 = uColuna7 + 1 TO uLimiteSuperior
                                    FOR uColuna9 = uColuna8 + 1 TO uLimiteSuperior
                                        FOR uColuna10 = uColuna9 + 1 TO uLimiteSuperior
                                            FOR uColuna11 = uColuna10 + 1 TO uLimiteSuperior
                                                FOR uColuna12 = uColuna11 + 1 TO uLimiteSuperior
                                                    FOR uColuna13 = uColuna12 + 1 TO uLimiteSuperior
                                                        FOR uColuna14 = uColuna13 + 1 TO uLimiteSuperior
                                                            FOR uColuna15 = uColuna14 + 1 TO uLimiteSuperior
                                                                DIALOG DOEVENTS
                                                                strTexto = "MEGASENA_COM_15_NUMEROS;" & strBolas(uColuna1) & ";" & strBolas(uColuna2) & ";" & strBolas(uColuna3) & ";"
                                                                strTexto = strTexto & strBolas(uColuna4) & ";" & strBolas(uColuna5) & ";" & strBolas(uColuna6) & ";" & strBolas(uColuna7) & ";"
                                                                strTexto = strTexto & strBolas(uColuna8) & ";" & strBolas(uColuna9) & ";" & strBolas(uColuna10) & ";" & strBolas(uColuna11) & ";"
                                                                strTexto = strTexto & strBolas(uColuna12) & ";" & strBolas(uColuna13) & ";" & strBolas(uColuna14) & ";" & strBolas(uColuna15)
                                                                PRINT #uArquivoDestino, strTexto
                                                                IF ERR <> 0 THEN
                                                                    strErro = "Erro: " & ERROR$(ERR)
                                                                    FUNCTION = -1
                                                                    RETURN
                                                                END IF
                                                                '
                                                                uContador += 1
                                                                IF uAtualizarTodas = 1 THEN
                                                                    strTexto = "Qt: " & FORMAT$(uContador, "###,#") & ", " & strTexto
                                                                    CONTROL SET TEXT hJanela_Principal, %IDC_LBL_LOG_TODAS, strTexto
                                                                END IF
                                                                ' Se o usu�rio clicou em 'Parar' ou  'Fechou a janela' a vari�vel abaixo deve estar com o valor 1
                                                                ' para podermos sair do loop
                                                                IF uSairTodas = 1 THEN
                                                                    strErro = "A gera��o foi interrompida por solicita��o do usu�rio."
                                                                    RETURN
                                                                END IF

                                                                ' Se o usu�rio clicar no bot�o 'Pausar' ficar no loop at� que 'Pausar' seja igual a '0' zero
                                                                ' ou o usu�rio clicou no bot�o 'Parar'
                                                                DO WHILE uPausarTodas = 1
                                                                    DIALOG DOEVENTS
                                                                LOOP
                                                            NEXT uColuna15
                                                        NEXT uColuna14
                                                    NEXT uColuna13
                                                NEXT uColuna12
                                            NEXT uColuna11
                                        NEXT uColuna10
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1
    '
    RETURN










END FUNCTION

'------------------------------------------------------------------------------
'   ** Main Application Entry Point **
'------------------------------------------------------------------------------
FUNCTION PBMAIN()
    PBFormsInitComCtls (%ICC_WIN95_CLASSES OR %ICC_DATE_CLASSES OR _
        %ICC_INTERNET_CLASSES)

    Exibir_Janela_Principal %HWND_DESKTOP
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'   ** CallBacks **
'------------------------------------------------------------------------------
CALLBACK FUNCTION Janela_Principal_CBK()
    STATIC strErro AS STRING

    SELECT CASE AS LONG CB.MSG
        CASE %WM_INITDIALOG
            ' Initialization handler
            ' Preencher as caixas de combina��o 'JOGO_TIPO' e 'APOSTAR COM'
            Carregar_Lista_Inicial()


        CASE %WM_NCACTIVATE
            STATIC hWndSaveFocus AS DWORD
            IF ISFALSE CB.WPARAM THEN
                ' Save control focus
                hWndSaveFocus = GetFocus()
            ELSEIF hWndSaveFocus THEN
                ' Restore control focus
                SetFocus(hWndSaveFocus)
                hWndSaveFocus = 0
            END IF

        CASE %WM_CLOSE
            IF MSGBOX("Deseja realmente sair, pois se houver gera��o de combina��es, o arquivo ficar� incompleto???", %MB_YESNO, "SAIR???") = %IDYES THEN
                uSairTodas = 1
            END IF
            FUNCTION = 1

        CASE %WM_COMMAND
            ' Process control notifications
            SELECT CASE AS LONG CB.CTL
                CASE %BTN_PARAR_TESTE
                    uPararTeste = 1

                CASE %BTN_PAUSAR_TESTE
                    uPausarTeste = IIF(uPausarTeste = 0, 1, 0)

                ' Se o usu�rio clicar no tipo do jogo, alterar, a caixa de combina��o "QT DE APOSTAS"
                CASE %IDC_CMB_JOGO_TIPO_TODAS
                    IF CB.CTLMSG = %CBN_SELCHANGE THEN
                        DIM strTexto AS STRING
                        DIM uIndice AS LONG

                        COMBOBOX GET SELECT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS TO uIndice
                        COMBOBOX GET TEXT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, uIndice TO strTexto
                        Carregar_Lista(strTexto)
                    END IF

                ' Se o usu�rio pressionou o bot�o Gerar, verificar ent�o qual
                ' jogo o usu�rio selecionou e a quantidade de bolas apostadas
                CASE %IDC_BTN_GERAR_TODAS
                    IF CB.CTLMSG = %BN_CLICKED THEN


                        ' Verificar qual jogo foi selecionado
                        DIM strJogoSelecionado AS STRING
                        DIM uJogoIndiceSelecionado AS LONG

                        COMBOBOX GET SELECT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS TO uJogoIndiceSelecionado
                        COMBOBOX GET TEXT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, uIndice TO strJogoSelecionado
                        '
                        ' Verificar qual o tipo de aposta foi selecionado
                        DIM strApostaSelecionado AS STRING
                        DIM uApostaIndiceSelecionado AS LONG

                        COMBOBOX GET SELECT hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS TO uApostaIndiceSelecionado
                        COMBOBOX GET TEXT hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, uIndice TO strApostaSelecionado

                        ' A caixa de combina��o come�a com n�meros e depois h� sequ�ncia alfab�ticas, por exemplo
                        ' '5 n�meros', '6 n�meros', ent�o devemos pegar a parte mais a esquerda composta somente de
                        ' n�meros
                        DIM uEspaco AS LONG
                        uEspaco = INSTR(strApostaSelecionado, ANY " ")
                        strApostaSelecionado = LEFT$(strApostaSelecionado, uEspaco - 1)

                        ' Verificar se h� algum pasta v�lida
                        DIM strPastaDestino AS STRING

                        CONTROL GET TEXT hJanela_Principal, %IDC_TXT_PASTA_DESTINO_TODAS TO strPastaDestino
                        IF ISFALSE(ISFOLDER(strPastaDestino)) THEN
                            MSGBOX "A pasta '" & strPastaDestino & "' n�o existe!!!", %MB_ICONERROR
                            EXIT FUNCTION
                        END IF

                        ' Desativar controle
                        CONTROL DISABLE hJanela_Principal, %IDC_BTN_GERAR_TODAS

                        ' Ativar Controle de Parada
                        CONTROL ENABLE hJanela_Principal, %IDC_BTN_PARAR_TODAS

                        IF Gerar_Todas_Combinacoes(strJogoSelecionado, VAL(strApostaSelecionado), strPastaDestino, strErro) = -1 THEN
                            MSGBOX "Um erro ocorreu: '" & strErro & "'.", %MB_ICONERROR
                        END IF

                        ' Ativar controle
                        CONTROL ENABLE hJanela_Principal, %IDC_BTN_GERAR_TODAS

                        ' Desativar Controle de Parada
                        CONTROL DISABLE hJanela_Principal, %IDC_BTN_PARAR_TODAS

                    END IF

                ' Se o usu�rio digita na caixa de texto, ativa ou desativar bot�o gerar
                CASE %IDC_TXT_PASTA_DESTINO_TODAS
                    IF CB.CTLMSG = %EN_CHANGE THEN
                        DIM strTextoDestino AS STRING
                        CONTROL GET TEXT hJanela_Principal, %IDC_TXT_PASTA_DESTINO_TODAS TO strTextoDestino
                        IF strTextoDestino = "" THEN
                            CONTROL DISABLE hJanela_Principal, %IDC_BTN_GERAR_TODAS
                        ELSE
                            CONTROL ENABLE hJanela_Principal, %IDC_BTN_GERAR_TODAS
                        END IF
                    END IF


                ' Se o usu�rio clica no bot�o 'Verificar status da gera��o'
                CASE %IDC_CHK_ATUALIZAR_TODAS
                    IF CB.CTLMSG = %BN_CLICKED THEN
                        CONTROL GET CHECK hJanela_Principal, %IDC_CHK_ATUALIZAR_TODAS TO uAtualizarTodas
                    END IF

                ' Se o usu�rio clicar no bot�o 'PARAR', definir a vari�vel uSairTodas = 1
                CASE %IDC_BTN_PARAR_TODAS
                    IF CB.CTLMSG = %BN_CLICKED THEN
                        uSairTodas = 1
                    END IF

                ' Se o usu�rio clicar no bot�o 'PAUSAR', definir a vari�vel para o valor do checkBox
                CASE %IDC_CHK_PAUSAR_TODAS
                    IF CB.CTLMSG = %BN_CLICKED THEN
                        CONTROL GET CHECK hJanela_Principal, %IDC_CHK_PAUSAR_TODAS TO uPausarTodas
                    END IF

               ' ---------------------------------------------------------------------------------
               ' Vamos analisar a parte referente ao Insert e Where que queremos processar
               ' ---------------------------------------------------------------------------------

               ' Se o usu�rio digita na caixa de texto, verificar se a caixa de texto est�
               ' vazia, se sim, desabilitar controles
               CASE %IDC_TXT_ARQUIVO_ORIGEM, %IDC_TXT_PASTA_DESTINO_INSERT
                   IF CB.CTLMSG = %EN_CHANGE THEN
                       LOCAL strArquivoOrigem AS STRING
                       DIM strPastaDestinoInsert AS STRING

                       CONTROL GET TEXT hJanela_Principal, %IDC_TXT_ARQUIVO_ORIGEM TO strArquivoOrigem
                       CONTROL GET TEXT hJanela_Principal, %IDC_TXT_PASTA_DESTINO_INSERT TO strPastaDestinoInsert

                       ' Se um ou ambos controles que est�o vazios desabilitar outros controles
                       IF strArquivoOrigem = "" OR strPastaDestinoInsert = "" THEN
                           CONTROL DISABLE hJanela_Principal, %IDC_CHK_STATUS_INSERT
                           CONTROL DISABLE hJanela_Principal, %IDC_CHK_PAUSAR_INSERT
                           CONTROL DISABLE hJanela_Principal, %IDC_BTN_PARAR_INSERT
                           CONTROL DISABLE hJanela_Principal, %IDC_BTN_GERAR_INSERT
                       ELSE
                           CONTROL ENABLE hJanela_Principal, %IDC_CHK_STATUS_INSERT
                           CONTROL ENABLE hJanela_Principal, %IDC_CHK_PAUSAR_INSERT
                           CONTROL ENABLE hJanela_Principal, %IDC_BTN_PARAR_INSERT
                           CONTROL ENABLE hJanela_Principal, %IDC_BTN_GERAR_INSERT
                       END IF
                    END IF


                ' Se o usu�rio clicar no bot�o 'Verificar status da Gera��o' definir a vari�vel global
                ' uPausarInsert
                CASE %IDC_CHK_STATUS_INSERT
                    CONTROL GET CHECK hJanela_Principal, %IDC_CHK_STATUS_INSERT TO uAtualizarInsert

                CASE %IDC_CHK_PAUSAR_INSERT
                    CONTROL GET CHECK hJanela_Principal, %IDC_CHK_PAUSAR_INSERT TO uPausarInsert

                ' Se o usu�rio clicar no bot�o 'Inserir senten�a 'commit' ap�s sql', definir a vari�vel global igual a 1

                CASE %IDC_COMMIT
                    CONTROL GET CHECK hJanela_PRINCIPAL, %IDC_COMMIT TO uCommit

                CASE %IDC_BTN_GERAR_INSERT
                    LOCAL strArquivoOrigemInsert AS STRING
                    'local strPastaDestinoInsert as string
                    IF CB.CTLMSG = %BN_CLICKED THEN
                        CONTROL GET TEXT hJanela_Principal, %IDC_TXT_ARQUIVO_ORIGEM TO strArquivoOrigemInsert
                        CONTROL GET TEXT hJanela_Principal, %IDC_TXT_PASTA_DESTINO_INSERT TO strPastaDestinoInsert
                    END IF

                    IF Gerar_Insert_Where_Todas_Combinacoes(strArquivoOrigemInsert, strPastaDestinoInsert, strErro) = -1 THEN
                        MSGBOX strErro, %MB_ICONERROR
                        EXIT FUNCTION
                    ELSE
                        MSGBOX "Executado com sucesso!!!"
                        EXIT FUNCTION
                    END IF
                    CLOSE

                CASE %IDC_BTN_PARAR_INSERT
                    IF MSGBOX("Voc� desejar parar o processamento dos n�meros???", %MB_YESNO, "PARAR") = %IDYES THEN
                        uPararInsert = 1
                    ELSE
                        uPararInsert = 0
                    END IF

                CASE %IDC_BTN_GERAR_LOTERIA
                    DIM strArquivo_Loteria AS STRING
                    'DIM strDestino_Loteria AS STRING

                    ' Indica o diret�rio csv que foi processado.
                    'DIM strDiretorio_Csv_Processado AS STRING

                    LOCAL strSucesso AS STRING

                    'CONTROL GET TEXT hJanela_Principal, %IDC_TXT_ARQUIVO_LOTERIA TO strArquivo_Loteria
                    'CONTROL GET TEXT hJanela_Principal, %IDC_TXT_DESTINO_LOTERIA TO strDestino_Loteria
                    'CONTROL GET TEXT hJanela_Principal, %DESTINO_CSV_PROCESSADO TO strDiretorio_csv_processado

                    ' Vai guardar o string indicando se houve um erro.
                    'LOCAL strErro AS STRING
                    strErro = ""

                    ' Vamos verificar se o diret�rio onde ser�o colocados os arquivos *.sql gerado existe.
                    LOCAL strdiretorio_sql_gerado AS STRING

                    CONTROL GET TEXT hJanela_Principal, %TXT_DIRETORIO_SQL_GERADO TO strdiretorio_sql_gerado

                    IF NOT ISFOLDER(strdiretorio_sql_gerado) THEN
                        strErro += "Diretorio '" + strDiretorio_sql_gerado + "', onde ficar�o os arquivos *.sql gerados, n�o existe."
                    END IF


                    ' Vamos verificar se o diret�rio onde ser�o colocados os arquivos *.csv processados existe.
                    LOCAL strDiretorio_csv_processado AS STRING

                    CONTROL GET TEXT hJanela_Principal, %TXT_DIRETORIO_CSV_PROCESSADO TO strDiretorio_csv_processado

                    IF NOT ISFOLDER(strdiretorio_csv_processado) THEN
                        strErro += IIF$(strErro <> "", $CRLF, "") + "Diretorio '" + strDiretorio_csv_processado + "', onde ficar�o os arquivos *.csv processados, n�o existe."
                    END IF

                    LOCAL strArquivo_CSV AS STRING
                    CONTROL GET TEXT hJanela_Principal, %TXT_ARQUIVO_CSV TO strArquivo_CSV

                    ' Vamos verificar se o diret�rio onde ser�o lidos os arquivos *.csv e onde ser�o movidos os arquivos *.csv quando processados
                    ' n�o s�o iguais.
                    IF PATHSCAN$(PATH, strArquivo_csv) = PATHSCAN$(PATH, strdiretorio_csv_processado) THEN
                        strErro += IIF$(strErro <> "", $CRLF, "") + "Diret�rio onde ser�o lidos os arquivos *.csv e " + $CRLF
                        strErro += "diretorio onde ser�o movidos os arquivos *.csv ap�s processados, s�o iguais."
                    END IF

                    ' Se o nome do arquivo fornecido pelo usu�rio � uma pasta, indicar um erro.
                    IF ISTRUE(ISFOLDER(strArquivo_csv)) THEN
                        strErro += IIF$(strErro <> "", $CRLF, "") + "'" + strArquivo_csv + "' n�o � um arquivo, � um diret�rio."
                    END IF

                    ' Se o arquivo fornecido pelo usu�rio, n�o termina em *.csv, indicar tamb�m um erro.
                    IF UCASE$(PATHNAME$(EXTN, strArquivo_csv)) <> ".CSV" THEN
                        strErro += IIF$(strErro <> "", $CRLF, "") + "'" + strArquivo_csv + "' n�o termina com a extens�o '.csv'."
                    END IF

                    ' Se foi encontrado algum erro, strErro � um string n�o vazio.
                    IF strErro <> "" THEN
                        MSGBOX strErro, %MB_ICONERROR OR %MB_TASKMODAL
                        EXIT FUNCTION
                    END IF

                    ' Se o usu�rio fornece um filtro ou n�o, iremos utilizar o loop e utilizar a senten�a 'DIR'
                    ' para verificar se h� mais arquivos.
                    LOCAL strArquivo_csv_localizado AS STRING
                    LOCAL strDiretorio_do_arquivo_csv AS STRING

                    strArquivo_csv_localizado = DIR$(strArquivo_csv)
                    strDiretorio_do_arquivo_csv = PATHNAME$(PATH, strArquivo_csv)

                    ' Coloca a barra final, se n�o houver.
                    IF RIGHT$(strDiretorio_csv_processado, 1) <> "\" THEN
                        strDiretorio_csv_processado += "\"
                    END IF

                    IF strArquivo_csv_localizado = "" THEN
                        MSGBOX "Nenhum arquivo localizado.", %MB_ICONERROR OR %MB_TASKMODAL
                        EXIT FUNCTION
                    END IF

                    ' Se chegarmos aqui, quer dizer que foi encontrado algum arquivo
                    ' Desativar ent�o o controle, enquanto executarmos a fun��o.
                    CONTROL DISABLE hJanela_Principal, %IDC_BTN_GERAR_LOTERIA

                    ' Vamos cronometrar o tempo que o arquivo foi executado.
                    LOCAL objTempo_Inicial, objTempo_Final AS IPOWERTIME
                    LET objTempo_Inicial = CLASS "POWERTIME"
                    LET objTempo_Final = CLASS "POWERTIME"

                    LOCAL uContador_de_Arquivos AS LONG
                    uContador_de_Arquivos = 1



                    objTempo_Inicial.Now
                    DO WHILE strArquivo_csv_Localizado <> ""
                        strArquivo_csv_localizado = strdiretorio_do_arquivo_csv + strArquivo_csv_localizado





                         IF Gerar_Insert_Where_Resultado_Loteria(strArquivo_csv_Localizado, strdiretorio_sql_gerado, strErro, strSucesso) = -1 THEN
                            MSGBOX strErro, %MB_ICONERROR OR %MB_TASKMODAL
                            CONTROL ENABLE hJanela_Principal, %IDC_BTN_GERAR_LOTERIA
                            EXIT FUNCTION
                         END IF
                         IF uSairTodas = 1 THEN
                             CONTROL ENABLE hJanela_Principal, %IDC_BTN_GERAR_LOTERIA
                             EXIT DO
                         END IF

                        strSucesso = "Qt de arquivos: " + FORMAT$(uContador_de_Arquivos, "0000000000") + $CRLF + strSucesso
                        CONTROL SET TEXT hJanela_Principal, %IDC_LOG, strSucesso
                        CONTROL REDRAW hJanela_Principal, %IDC_LOG
                        INCR uContador_de_Arquivos

                        ' Vamos mover o arquivo
                        TRY
                            LOCAL strArquivo_csv_movido AS STRING
                            strArquivo_Csv_Movido = strDiretorio_csv_processado + PATHSCAN$(NAMEX, strArquivo_csv_localizado)

                            ' Move o arquivo para o destino
                            NAME strArquivo_csv_localizado AS strArquivo_csv_movido
                        CATCH
                            MSGBOX ERROR$(ERR)
                            CLOSE
                            EXIT FUNCTION
                        END TRY

                        strArquivo_csv_Localizado = DIR$
                        DIALOG DOEVENTS
                    LOOP

                    strSucesso = "Arquivos lidos: " + FORMAT$(uContador_de_arquivos, "000000000") + $CRLF
                    strSucesso += tempo_decorrido(objTempo_Inicial, objTempo_Final)

                    CONTROL SET TEXT hJanela_Principal, %IDC_LOG, strSucesso + $CRLF + "EXECUTADO COM SUCESSO!!!"
                    MSGBOX "Arquivos processados com sucesso.", %MB_ICONINFORMATION
                    CONTROL ENABLE hJanela_Principal, %IDC_BTN_GERAR_LOTERIA

            END SELECT
    END SELECT
END FUNCTION

FUNCTION Tempo_Decorrido (BYVAL objTempo_Inicial AS IPOWERTIME, BYVAL objTempo_Final AS IPOWERTIME)  AS STRING
    DIM strTempo AS STRING
    DIM uSinal AS LONG
    DIM uDia AS LONG
    DIM uHora AS LONG
    DIM uMinuto AS LONG
    DIM uSegundo AS LONG
    DIM umiSegundo AS QUAD
    DIM uTick AS QUAD

    objTempo_Final.Now()
    objTempo_Final.TimeDiff(objTempo_Inicial, uSinal, uDia, uHora, uMinuto, uSegundo, uMiSegundo, uTick)

    strTempo += "Data_Hora_Inicial: "
    strTempo += FORMAT$(objTempo_Inicial.Day, "00") + "/"
    strTempo += FORMAT$(objTempo_Inicial.Month, "00") + "/"
    strTempo += FORMAT$(objTempo_Inicial.Year, "0000") + " "
    strTempo += FORMAT$(objTempo_Inicial.Hour, "00") + "h,"
    strTempo += FORMAT$(objTempo_Inicial.Minute, "00") + "min,"
    strTempo += FORMAT$(objTempo_Inicial.Second, "00") + "sec,"
    strTempo += FORMAT$(objTempo_Inicial.mSecond, "000") + "msec,"
    strTempo += FORMAT$(objTempo_Inicial.Tick, "000") + "t." + $CRLF
    strTempo += "Data_Hora_Final: "
    strTempo += FORMAT$(objTempo_Final.Day, "00") + "/"
    strTempo += FORMAT$(objTempo_Final.Month, "00") + "/"
    strTempo += FORMAT$(objTempo_Final.Year, "0000") + " "
    strTempo += FORMAT$(objTempo_Final.Hour, "00") + "h,"
    strTempo += FORMAT$(objTempo_Final.Minute, "00") + "min,"
    strTempo += FORMAT$(objTempo_Final.Second, "00") + "sec,"
    strTempo += FORMAT$(objTempo_Final.mSecond, "000") + "msec,"
    strTempo += FORMAT$(objTempo_Final.Tick, "000") + "t." + $CRLF + $CRLF
    strTempo += "Tempo Decorrido: "
    strTempo += FORMAT$(uDia, "00") + "d, "
    strTempo += FORMAT$(uHora, "00") + "h, "
    strTempo += FORMAT$(uMinuto, "00") + "min, "
    strTempo += FORMAT$(uSegundo, "00") + "sec, "
    strTempo += FORMAT$(uMiSegundo, "00") + "msec, "
    strTempo += FORMAT$(uTick, "00") + "tick."

    Tempo_Decorrido = strTempo

END FUNCTION

'------------------------------------------------------------------------------

' Esta fun��o valida todos os campos, e se tiver todas as linhas com o layout correto, retorna o arranjo
' Esta fun��o, tamb�m, verifica se cada valor em cada campo, corresponde ao intervalo para o jogo pretendido.
FUNCTION Validar_Campos_Loteria (BYVAL strArquivoOrigem AS STRING, BYREF strLinhas() AS STRING, BYREF strErro AS STRING) AS LONG
    DIM uArquivoOrigem AS LONG
    uArquivoOrigem = FREEFILE
    OPEN strArquivoOrigem FOR INPUT AS #uArquivoOrigem
    IF ERR <> 0 THEN
        strErro = ERROR$(ERR)
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos colocar no in�cio da vari�vel, o nome do arquivo.
    ' Na vers�o 5.0 deste arqui, colocamos o nome do arquivo no �nicio do string que indica
    ' erro, para ficar f�cil identificar o arquivo.
    strErro = "Erro no arquivo: '" + FILENAME$(uArquivoOrigem) & "' " & $CRLF
    '
    IF ISTRUE(EOF(uArquivoOrigem)) THEN
        strErro += "Erro, arquivo est� vazio."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos contar quantas linhas h�
    DIM uLinhas AS LONG
    FILESCAN #uArquivoOrigem, RECORDS TO uLinhas

    ' Gravar todas as linhas no arranjo
    DIM uLinhasGravadas AS LONG
    LOCAL strLinhas() AS STRING

    REDIM strLinhas(1 TO uLinhas)
    LINE INPUT #uArquivoOrigem, strLinhas() RECORDS uLinhas TO uLinhasGravadas
    CLOSE uArquivoOrigem



    IF uLinhasGravadas <> uLinhas THEN
        strErro += "N�o foi poss�vel ler todas as linhas."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    LOCAL uCampos_Quantidade AS LONG
    uCampos_Quantidade = PARSECOUNT(strlinhas(1), ANY ";")

    ' A primeira linha deve ter este layout
    ' JOGO_TIPO;CONCURSO;DATA SORTEIO;BOLA1;BOLA2;...
    DIM strCampos() AS STRING
    REDIM strCampos(1 TO uCampos_Quantidade)
    PARSE UCASE$(strLinhas(1)), strCampos(), ANY ";"

    ' Validar Campos
    IF TRIM$(strCampos(1)) <> "JOGO_TIPO" THEN
        strErro += "Erro, Linha 1, Campo 1, deve ter a palavra 'JOGO_TIPO'."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    IF TRIM$(strCampos(2)) <> "CONCURSO" THEN
        strErro += "Erro, linha 1, Campo 2, deve ter a palavra 'CONCURSO'."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    IF TRIM$(strCampos(3)) <> "DATA SORTEIO" THEN
        strErro += "Erro, linha 1, Campo 3, deve ter a palavra 'DATA SORTEIO'."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    DIM uB AS LONG
    FOR uB = 4 TO uCampos_Quantidade
        ' Vamos verificar se existe os campos para indicar as bolas
        IF TRIM$(strCampos(uB)) <> "BOLA" & FORMAT$(uB-3) THEN
            strErro += "Erro, linha 1, Campo " & FORMAT$(uB) & " deve ter a palavra 'BOLA" & FORMAT$(uB-3) & "'."
            FUNCTION = -1
            EXIT FUNCTION
        END IF
    NEXT uB

    ' Se h� somente uma linha, ent�o devemos retornar com erro
    IF uLinhas = 1 THEN
        strErro += "Erro, s� h� uma linha com cabecalho."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos verificar se todas as linhas tem a mesma quantidade de campos campo
    DIM uA AS LONG
    FOR uA = 1 TO uLinhas
        IF uCampos_Quantidade <> PARSECOUNT(strLinhas(uA), ANY ";") THEN
            ERASE strLinhas()
            strErro += "Erro, linha '" & FORMAT$(uA) & "' n�o � igual � quantidade de campos da linha anterior."
            FUNCTION = -1
            EXIT FUNCTION
        END IF
    NEXT uA

    ' <INCLUIDO EM 24/12/2014>
    ' Vamos verificar se todas as linhas corresponde ao mesmo jogo e se, os n�meros nos campos que correspondem
    ' �s bolas, coincidem com o intervalo corresponde ao jogo espec�ficio

    ' Vamos o nome do jogo, que est� na segunda linha, este dever� est� em todas as linhas seguintes.

    ' O campo 1, da linha 2, indica qual � o jogo, que est� sendo processado.
    ' Ent�o, todas as linhas seguintes, no campo 1, deve ter o mesmo jogo.
    LOCAL strJogoTipo AS STRING
    PARSE UCASE$(strLinhas(2)), strCampos(), ANY ";"

    strJogoTipo = strCampos(1)
    IF  strJogoTipo <> "QUINA" AND strJogoTipo <> "MEGASENA" AND _
        strJogoTipo <> "LOTOFACIL" AND strJogoTipo <> "LOTOMANIA" AND _
        strJogoTipo <> "DUPLASENA" AND strJogoTipo <> "INTRALOT_MINAS_5" AND _
        strJogoTipo <> "INTRALOT_LOTO_MINAS" AND strJogoTipo <> "INTRALOT_KENO_MINAS" THEN
            strErro += "O jogo: '" + strJogoTipo + "' n�o existe."
            FUNCTION = -1
            EXIT FUNCTION
    END IF



    ' A quantidade de bolas � definido pela quantidade de campos menos 3.
    ' Pois, os tr�s primeiros campos, indica o:
    ' tipo do jogo, concurso e a data.
    LOCAL uBolasQuantidade AS LONG
    uBolasQuantidade = uCampos_Quantidade - 3


    FOR uA = 2 TO uLinhas
        PARSE UCASE$(strLinhas(uA)), strCampos(), ANY ";"

        '
        ' Validar o tipo de jogo, todas as linhas deve ter o mesmo jogo.
        '

        IF strJogoTipo <> strCampos(1) THEN
            strErro += "O tipo do jogo na linha: " + FORMAT$(uA) + " no campo 1, n�o corresponde ao " & _
                        " ao jogo: '" & strJogoTipo & "' definido nas linhas anteriores."
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        ' ********** VALIDAR CAMPO DE DATA *****************
        ' A data deve est� neste formato: dd/mm/yyyy ou d/m/yyyy, exemplo: 1/1/2015 ou 01/1/2015 ou 1/01/2015 ou 01/01/2015
        ' Verificar se cont�m os caracteres v�lidos.
        ' O terceiro campo indica a data

        IF VERIFY(strCampos(3), "0123456789/") <> 0 THEN
            strErro += "Erro, linha " + FORMAT$(uA) + ", campo 3, formato de data inv�lido: '" + strCampos(3) + "'"
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        ' Vamos separar os campos de data, pelo caractere '/'
        LOCAL strData() AS STRING
        REDIM strData(1 TO 3)

        ' Vamos verificar se h� somente tr�s campos
        IF PARSECOUNT(strCampos(3), ANY "/") <> 3 THEN
            strErro += "Erro, linha " + FORMAT$(uA) + ", campo 3, formato de data inv�lido: '" + strCampos(3) + "'"
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        PARSE strCampos(3), strData(), ANY "/"
        ' Se chegamos at� aqui, quer dizer que h� 3 campos, vamos verificar se os campos de data e m�s est�o dentro
        ' da faixa v�lida.
        DIM uData_Dia AS LONG
        uData_Dia = VAL(strData(1))

        IF uData_Dia < 1 OR uData_Dia > 31 THEN
            strErro += "Erro, linha " + FORMAT$(uA) + ", campo 3, formato de data inv�lido: '" + strCampos(3) + "'." + $CRLF
            strErro += "O intervalo v�lido � maior ou igual a 1 e menor ou igual a 31."
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        DIM uData_Mes AS LONG
        uData_Mes = VAL(strData(2))

        ' Se fora da faixa v�lida para o m�s, indicar um erro.
        IF uData_Mes < 1 OR uData_Mes > 12 THEN
            strErro += "Erro, linha " + FORMAT$(uA) + ", campo 3, formato de data inv�lido: '" + strCampos(3) + "'." + $CRLF
            strErro += "O intervalo v�lido para o m�s deve estar entre 1 e 12."
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        ' Cada m�s, excetuando fevereiro, termina em 30 ou 31.
        ' Para o m�s de fevereiro deveremos ver se � bissexto ou n�o para considerar, se � 29 ou 28.

        ' Se o dia termina em 31, deveremos verificar se o m�s �
        ' Janeiro, Mar�o, Maio, Julho, Agosto, Outubro, Dezembro, se n�o �, � um erro.
        IF uData_Dia = 31 THEN
            SELECT CASE AS LONG uData_Mes
                CASE 1, 3, 5, 7, 8, 10, 12
                    ' N�o faz nada, rsrs.

                CASE ELSE
                    strErro += "Erro, linha " + FORMAT$(uA) + ", campo 3, formato de data inv�lido: '" + strCampos(3) + "'." + $CRLF
                    strErro += "Dia 31 � inv�lido para o m�s: " & FORMAT$(uData_Mes) & "."
                    FUNCTION = -1
                    EXIT FUNCTION
            END SELECT

        ' Devemos colocar os operandos do operador 'OR', entre par�nteses
        ' Por que a preced�ncia do operador 'OR' � menor que do operador 'AND'

        ELSEIF uData_Dia = 29 AND uData_Mes = 2 THEN
            ' Vamos verificar se o ano � bissexto, se n�o for, uData_Dia deve ser diferente de 29.
            ' Um n�mero � bissexto, se for
            ' Divis�vel por 400;
            ' Divis�vel por 4 e n�o divis�vel por 100;

            LOCAL uData_Ano AS LONG
            uData_Ano = VAL(strData(3))

            ' Se o ano n�o � bissexto, ent�o o dia 29 n�o pode existir.
            IF ISFALSE((uData_Ano MOD 4 = 0 AND uData_Ano MOD 100 <> 0) OR uData_Ano MOD 400 = 0) THEN
                strErro += "Erro, linha " + FORMAT$(uA) + ", campo 3, formato de data inv�lido: '" + strCampos(3) + "'." + $CRLF
                strErro += "Dia 29 para o m�s de fevereiro somente � v�lido para ano bissexto, entretanto, " + $CRLF
                strErro += "o ano " + FORMAT$(uData_Ano) + " n�o � bissexto."
                FUNCTION = -1
                EXIT FUNCTION
            END IF
        END IF


        ' **************************** VERIFICAR SE H� BOLAS REPETIDAS ********************************
        '
        IF uA = 10 THEN
            '
        END IF


        LOCAL uX AS LONG
        LOCAL uY AS LONG
        FOR uX = 3 TO uCampos_Quantidade
            FOR uY = uX + 1 TO uCampos_Quantidade
                ' Se jogo � DUPLASENA, haver� 12 bolas, onde as seis primeiras correspondem ao primeiro sorteio
                ' e, as seis �ltimas bolas, correspondem ao segundo sorteio.
                ' ent�o, n�o podemos comparar o primeiro sorteio com o segundo.
                ' O primeiro sorteio come�a no �ndice 3 e vai at� 8.
                ' O segundo sorteio come�a em 9 e vai at� 14.
                ' Ent�o, se uX � menor ou igual a 8, estamos no primeiro sorteio, ent�o
                ' n�o devemos comparar com o segundo sorteio que come�a em uY = 9
                IF uA = 10 THEN
                    '
                END IF

                IF uX <= 8 AND uY > 8 THEN
                    EXIT FOR
                END IF


                IF strCampos(uX) = strCampos(uY) THEN
                    strErro += "Erro, linha " + FORMAT$(uA) + ", campos " + FORMAT$(uX) + " e " + FORMAT$(uY) + "." + $CRLF
                    strErro += "Valores iguais em tais campos." + $CRLF
                    strErro += "N�o pode haver bolas id�nticas."
                    FUNCTION = -1
                    EXIT FUNCTION
                END IF
            NEXT uY
        NEXT uX















        ' Vamos verificar se a quantidade de campos que correspondem �s bolas,
        ' correspondem igualmente � quantidade de bolas sorteadas para o jogo respectivo

        IF strJogoTipo = "QUINA" AND uBolasQuantidade <> 5 THEN
            strErro += "No jogo 'QUINA', a quantidade de bolas sorteadas � '5', entretanto, no arquivo, " & _
                        " h� somente " + FORMAT$(uBolasQuantidade) + " campos que correspondem �s bolas."
            FUNCTION = -1
            EXIT FUNCTION
        ELSEIF strJogoTipo = "MEGASENA" AND uBolasQuantidade <> 6 THEN
            strErro += "No jogo 'MEGASENA', a quantidade de bolas sorteadas � '6', entretanto, no arquivo, " & _
                        " h� somente " + FORMAT$(uBolasQuantidade) + " campos que correspondem �s bolas."
            FUNCTION = -1
            EXIT FUNCTION
        ELSEIF strJogoTipo = "LOTOMANIA" AND uBolasQuantidade <> 20 THEN
            strErro += "No jogo 'LOTOMANIA', a quantidade de bolas sorteadas � '20', entretanto, no arquivo, " & _
                        " h� somente " + FORMAT$(uBolasQuantidade) + " campos que correspondem �s bolas."
            FUNCTION = -1
            EXIT FUNCTION
        ELSEIF strJogoTipo = "LOTOFACIL" AND uBolasQuantidade <> 15 THEN
            strErro += "No jogo 'LOTOFACIL', a quantidade de bolas sorteadas � '15', entretanto, no arquivo, " & _
                        " h� somente " + FORMAT$(uBolasQuantidade) + " campos que correspondem �s bolas."
            FUNCTION = -1
            EXIT FUNCTION

        ' No caso do jogo 'DUPLASENA' em um mesmo concurso s�o sorteados 6 bolas, duas vezes,
        ' No resultado que coletamos do site da caixa, todas as bolas do mesmo concurso s�o dispostas todas na mesma linha, separando,
        ' em primeiro lugar, as bolas do primeiro sorteio e em seguida, as bolas do segundo sorteio
        ' por isso, para n�o interferir, no layout, que j� colocamos aqui em c�digo,
        ' vamos dispor tais bolas na mesma sequencia, onde as primeiras 6 bolas correspondem ao primeiro
        ' sorteio e as 6 bolas restantes, correspondem ao segundo sorteio.
        ELSEIF strJogoTipo = "DUPLASENA" AND uBolasQuantidade <> 12 THEN
            strErro += "No jogo 'DUPLASENA', s�o sorteados, em um �nico concurso, 2 sorteios, onde cada sorteio " & _
                    " corresponde a 6 n�meros, entretanto, no arquivo h� somente " + FORMAT$(uBolasQuantidade) & _
                    " campos que correspondem �s bolas.
            FUNCTION = -1
            EXIT FUNCTION


        ELSEIF strJogoTipo = "INTRALOT_MINAS_5" AND uBolasQuantidade <> 5 THEN
            strErro += "No jogo 'INTRALOT_MINAS_5' s�o sorteados 5 n�meros, entretanto, no arquivo, " & _
                        " h� somente " + FORMAT$(uBolasQuantidade) + " campos que correspondem �s bolas."
            FUNCTION = -1
            EXIT FUNCTION

        ' Inclu�do na vers�o 5.0
        ELSEIF strJogoTipo = "INTRALOT_LOTO_MINAS" AND uBolasQuantidade <> 6 THEN
            strErro += "No jogo 'INTRALOT_LOTO_MINAS' s�o sorteados 6 n�meros, entretanto, no arquivo, " & _
                        "h� somente " + FORMAT$(uBolasQuantidade) + " campos que correspondem �s bolas."
            FUNCTION = -1
            EXIT FUNCTION

        ' Inclu�do na vers�o 5.0
        ELSEIF strJogoTipo = "INTRALOT_KENO_MINAS" AND uBolasQuantidade <> 20 THEN
            strErro += "No jogo 'INTRALOT_KENO_MINAS' s�o sorteados 20 n�meros, entretanto, no arquivo, " & _
                        "h� somente " + FORMAT$(uBolasQuantidade) + " campos que correspondem �s bolas."
            FUNCTION = -1
            EXIT FUNCTION
        END IF


    NEXT


END FUNCTION

FUNCTION Abrir_Arquivo_Para_Gravar_Loteria(BYVAL strJogo AS STRING, BYVAL strPastaDestino AS STRING, BYREF uInsert_Jogo_Bolas AS LONG, _
                                            BYREF uInsert_Jogo_Combinacoes AS LONG, BYREF uWhere_Jogo_Bolas AS LONG, _
                                            BYREF uWhere_Jogo_Combinacoes AS LONG, BYREF strErro AS STRING) AS LONG
    DIM objTempo AS IPOWERTIME
    LET objTempo = CLASS "POWERTIME"
    objTempo.Now()
    DIM strTempo AS STRING
    strTempo =  FORMAT$(objTempo.Day, "00") & "_" & FORMAT$(objTempo.Month, "00") & "_" & FORMAT$(objTempo.Year, "00") & _
                "_" & FORMAT$(objTempo.Hour, "00") & "_" & FORMAT$(objTempo.Minute, "00") & "_" & FORMAT$(objTempo.Second, "00") & _
                "_" & FORMAT$(objTempo.mSecond, "000")

    LOCAL strInsert_Jogo_Bolas, strInsert_Jogo_Combinacoes AS STRING

    strInsert_Jogo_Bolas = strPastaDestino & strJogo & "_JOGO_BOLAS_INSERT_" & strTempo & "_000001.sql"

    uInsert_Jogo_Bolas = FREEFILE
    TRY
        OPEN strInsert_Jogo_Bolas FOR OUTPUT AS uInsert_Jogo_Bolas
    CATCH
        strErro = ERROR$(ERR)
        FUNCTION = -1
    END TRY
END FUNCTION

'
'   Esta fun��o retorna todas as combina��es poss�veis para um conjunto
'
' Par�metros:
'   strTexto:           Um string onde os elementos que fazem parte do conjunto � separado por um delimitador, por exemplo:
'                       strTexto --> 1;2;3;4;5
'   strDelimitador:     O caractere delimitador a ser o separador
'   strCombinacoes():   Retorna um arranjo, indicando todas as combina��es geradas.
'   strErro:            Se ocorrer algum erro, indicar nesta vari�vel.
'
' Obs.:                 Na fun��o, por exemplo, o conjunto, 1;2;3;4, retornaria, alguns exemplos, 1;2, 1;3, 1;2;3, 2;3, 2;3;4.
'                       Ou seja, a fun��o retorna somente o conjunto que n�o se repetem. Conclui-se:
'                       Na fun��o: 2;3 e 3;2 s�o a mesma coisa.


FUNCTION Combinacao(BYVAL strTexto AS STRING, BYVAL strDelimitador AS STRING, BYREF strCombinacoes() AS STRING, BYREF strErro AS STRING) AS LONG
    ' Vamos verificar se os argumentos est�o corretos.
    IF strTexto = "" THEN
        strErro = "Erro: 'strTexto' est� com o conte�do vazio."
        FUNCTION = -1
        EXIT FUNCTION
    END IF
    '
    '
    IF LEN(strDelimitador) > 1 THEN
        strErro = "Erro: O comprimento de 'strDelimitador' deve ser igual a 1."
        FUNCTION = -1
        EXIT FUNCTION
    END IF


    DIM combinados_Colecao AS ILINKLISTCOLLECTION
    LET combinados_Colecao = CLASS "LinkListCollection"
    IF ISFALSE(ISOBJECT(combinados_Colecao)) THEN
        strErro = "Erro: N�o foi poss�vel criar o objecto ILinkListCollection"
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos separar strTexto em campos, utilizando o 'strDelimitador'
    LOCAL uQuantidade_de_Campos AS LONG
    uQuantidade_de_Campos = PARSECOUNT(strTexto, ANY strDelimitador)
    IF uQuantidade_de_Campos = 1 THEN
        REDIM strCombinacoes(0 TO 0)
        strCombinacoes(0) = strTexto
    END IF

    ' Vamos criar os arranjos principais do programa
    LOCAL strCampos() AS STRING
    LOCAL uIndice() AS LONG
    LOCAL strCampos_Combinados AS STRING

    ' Vamos redimensionar os arranjos
    REDIM uIndice(1 TO uQuantidade_de_Campos)
    REDIM strCampos(1 TO uQuantidade_de_Campos)
    PARSE strTexto, strCampos(), ANY strDelimitador

    ' Vamos ordenar
    ARRAY SORT strCampos()

    ' Declara as vari�veis que ser�o utilizados no loop for
    LOCAL uColuna1, uColuna2, uColuna3, uColuna4, uColuna5, uColuna6, uColuna7 AS LONG
    LOCAL uColuna8, uColuna9, uColuna10, uColuna11, uColuna12, uColuna13, uColuna14 AS LONG
    LOCAL uColuna15 AS LONG

    LOCAL uLimite AS LONG

  ' Vamos executar todas as combina��es poss�veis.
    FOR uColuna1 = 1 TO uQuantidade_de_Campos
        uIndice(1) = uColuna1
        uLimite = 1
        GOSUB GERAR_COMBINACAO
        '
        FOR uColuna2 = uColuna1 + 1 TO uQuantidade_de_Campos
            uIndice(2) = uColuna2
            uLimite = 2
            GOSUB GERAR_COMBINACAO
            '
            FOR uColuna3 = uColuna2 + 1 TO uQuantidade_de_Campos
                uIndice(3) = uColuna3
                uLimite = 3
                GOSUB GERAR_COMBINACAO
                '
                FOR uColuna4 = uColuna3 + 1 TO uQuantidade_de_Campos
                    uIndice(4) = uColuna4
                    uLimite = 4
                    GOSUB GERAR_COMBINACAO
                    '
                    FOR uColuna5 = uColuna4 + 1 TO uQuantidade_de_Campos
                        uIndice(5) = uColuna5
                        uLimite = 5
                        GOSUB GERAR_COMBINACAO
                        '
                        FOR uColuna6 = uColuna5 + 1 TO uQuantidade_de_Campos
                            uIndice(6) = uColuna6
                            uLimite = 6
                            GOSUB GERAR_COMBINACAO
                            '                              '
                            FOR uColuna7 = uColuna6 + 1 TO uQuantidade_de_Campos
                                uIndice(7) = uColuna7
                                uLimite = 7
                                GOSUB GERAR_COMBINACAO
                                '
                                FOR uColuna8 = uColuna7 + 1 TO uQuantidade_de_Campos
                                    uIndice(8) = uColuna8
                                    uLimite = 8
                                    GOSUB GERAR_COMBINACAO
                                    '                                         '
                                    FOR uColuna9 = uColuna8 + 1 TO uQuantidade_de_Campos
                                        uIndice(9) = uColuna9
                                        uLimite = 9
                                        GOSUB GERAR_COMBINACAO
                                        '
                                        FOR uColuna10 = uColuna9 + 1 TO uQuantidade_de_Campos
                                            uIndice(10) = uColuna10
                                            uLimite = 10
                                            GOSUB GERAR_COMBINACAO
                                            '
                                            FOR uColuna11 = uColuna10 + 1 TO uQuantidade_de_Campos
                                                uIndice(11) = uColuna11
                                                uLimite = 11
                                                GOSUB GERAR_COMBINACAO
                                                '
                                                FOR uColuna12 = uColuna11 + 1 TO uQuantidade_de_Campos
                                                    uIndice(12) = uColuna12
                                                    uLimite = 12
                                                    GOSUB GERAR_COMBINACAO
                                                    '
                                                    FOR uColuna13 = uColuna12 + 1 TO uQuantidade_de_Campos
                                                        uIndice(13) = uColuna13
                                                        uLimite = 13
                                                        GOSUB GERAR_COMBINACAO
                                                        '
                                                        FOR uColuna14 = uColuna13 + 1 TO uQuantidade_de_Campos
                                                            uIndice(14) = uColuna14
                                                            uLimite = 14
                                                            GOSUB GERAR_COMBINACAO
                                                            '
                                                            FOR uColuna15 = uColuna14 + 1 TO uQuantidade_de_Campos
                                                                uIndice(15) = uColuna15
                                                                uLimite = 15
                                                                GOSUB GERAR_COMBINACAO
                                                                '
                                                            NEXT uColuna15
                                                        NEXT uColuna14
                                                    NEXT uColuna13
                                                NEXT uColuna12
                                            NEXT uColuna11
                                        NEXT uColuna10
                                    NEXT uColuna9
                                NEXT uColuna8
                            NEXT uColuna7
                        NEXT uColuna6
                    NEXT uColuna5
                NEXT uColuna4
            NEXT uColuna3
        NEXT uColuna2
    NEXT uColuna1

    ' Vamos verificar quantas combina�oes foram criados
    DIM uQuantidade AS LONG
    uQuantidade = combinados_colecao.count

    IF uQuantidade > 0 THEN
        REDIM strCombinacoes(1 TO uQuantidade)

        DO WHILE uQuantidade <> 0
            strCombinacoes(uQuantidade) = VARIANT$(combinados_colecao.Item(uQuantidade))
            uQuantidade -= 1
        LOOP
    END IF



    FUNCTION = 0


    EXIT FUNCTION

GERAR_COMBINACAO:
    strCampos_Combinados = ""

    DO WHILE uPausarTeste = 1
        DIALOG DOEVENTS
        CONTROL SET TEXT hJanela_Principal, %LBL_TESTE, "<< EM PAUSA >>"

        IF uPararTeste = 1 THEN
            FUNCTION = -1
            EXIT FUNCTION
        END IF
    LOOP

    IF uPararTeste = 1 THEN
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    LOCAL lngA AS LONG
    FOR lngA = 1 TO uLimite
        strCampos_Combinados = strCampos_Combinados & "_" & strCampos(uIndice(lngA))
    NEXT
    combinados_Colecao.add(strCampos_Combinados)

    CONTROL SET TEXT hJanela_Principal, %LBL_TESTE, strCampos_Combinados
    DIALOG DOEVENTS


    RETURN




















END FUNCTION



'
'   Esta fun��o retorna todas as combina��es poss�veis para um conjunto
'
' Par�metros:
'   strTexto:           Um string onde os elementos que fazem parte do conjunto � separado por um delimitador, por exemplo:
'                       strTexto --> 1;2;3;4;5
'   strDelimitador:     O caractere delimitador a ser o separador
'   strCombinacoes():   Retorna um arranjo, indicando todas as combina��es geradas.
'   strErro:            Se ocorrer algum erro, indicar nesta vari�vel.
'
' Obs.:                 Na fun��o, por exemplo, o conjunto, 1;2;3;4, retornaria, alguns exemplos, 1;2, 1;3, 1;2;3, 2;3, 2;3;4.
'                       Ou seja, a fun��o retorna somente o conjunto que n�o se repetem. Conclui-se:
'                       Na fun��o: 2;3 e 3;2 s�o a mesma coisa.



' ****************************
' ALTERAR ESTE
' ****************************


GLOBAL uContar_Grupo AS LONG
FUNCTION Gerar_Insert_Where_Resultado_Loteria(BYVAL strArquivoOrigem AS STRING, BYVAL strPastaDestino AS STRING, BYREF strErro AS STRING, BYREF strSucesso AS STRING) AS LONG
    ' Vamos verificar se arquivo existe
    IF ISFALSE(ISFILE(strArquivoOrigem)) THEN
        strErro = "Erro, arquivo '" & strArquivoOrigem & "' n�o existe."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos verificar se a pasta existe
    IF ISFALSE(ISFOLDER(strPastaDestino)) THEN
        strErro = "Erro, pasta de destino '" & strPastaDestino & "' n�o existe."
        FUNCTION = -1
        EXIT FUNCTION
    END IF
    ' Verificar se tem o caractere '\'
    IF RIGHT$(strPastaDestino, 1) <> "\" THEN
        strPastaDestino &= "\"
    END IF

    LOCAL strLinhas() AS STRING
    ' Vamos Validar se todas as linhas tem a mesma quantidade de campos
    IF Validar_Campos_Loteria(strArquivoOrigem, strLinhas(), strErro) = -1 THEN
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    DIM strCampos() AS STRING
    DIM uQuantidade_de_Bolas AS LONG
    DIM uQuantidade_de_Campos AS LONG

    ' Vamos verificar a quantidade de campos
    ' A primeira linha � o cabe�alho, devemos ler a segunda linha
    uQuantidade_de_Campos = PARSECOUNT(strLinhas(2), ANY ";")
    REDIM strCampos(1 TO uQuantidade_de_Campos)
    PARSE strLinhas(2), strCampos(), ANY ";"


    DIM strJogo AS STRING
    strJogo = UCASE$(strCampos(1))
    IF  strJogo <> "QUINA" AND _
        strJogo <> "MEGASENA" AND _
        strJogo <> "LOTOFACIL" AND _
        strJogo <> "LOTOMANIA" AND _
        strJogo <> "DUPLASENA" AND _
        strJogo <> "INTRALOT_MINAS_5" AND _
        strJogo <> "INTRALOT_LOTO_MINAS" AND _
        strJogo <> "INTRALOT_KENO_MINAS" THEN
            strErro = "Erro, jogo desconhecido '" & strJogo & "'."
            FUNCTION = -1
         EXIT FUNCTION
    END IF

    ' Abrir arquivos para gravacao
    LOCAL uInsert_Jogo_Bolas, uInsert_Jogo_Combinacoes AS LONG
    LOCAL uWhere_Jogo_Bolas, uWhere_Jogo_Combinacoes AS LONG

    IF Abrir_Arquivo_Para_Gravar_Loteria(strJogo, strPastaDestino, uInsert_Jogo_Bolas, uInsert_Jogo_Combinacoes, _
                                        uWhere_Jogo_Bolas, uWhere_Jogo_Combinacoes, strErro) = -1 THEN
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    DIM uBolas() AS LONG
    DIM uIndice() AS LONG

    ' O layout come�a assim: JOGO_TIPO;CONCURSO;DATA e depois vem as bolas, por isso, iremos substrair 3
    uQuantidade_de_Bolas = uQuantidade_de_Campos - 3

    REDIM uBolas(1 TO uQuantidade_de_Bolas)
    REDIM uIndice(1 TO uQuantidade_de_Bolas)

    DIM strBolas_Conjunto AS STRING
    DIM uA AS LONG
    DIM uB AS LONG
    DIM uLinhas AS LONG

    LOCAL objTempo_Inicial, objTempo_Final AS IPOWERTIME
    LET objTempo_Inicial = CLASS "POWERTIME"
    LET objTempo_Final = CLASS "POWERTIME"

    LOCAL quantidade_de_combinacoes AS QUAD
    quantidade_de_combinacoes = 0
    uPararInsert = 0
    uPausarInsert = 0

    uContar_Grupo = 1
    uLInhas = UBOUND(strLinhas())
    objTempo_Inicial.Now()
    FOR uA = 2 TO uLinhas
        CONTROL SET TEXT hJanela_Principal, %IDC_LOG_LOTERIA, "Arquivo: " & $CRLF & strArquivoOrigem & $CRLF & _
                                                                "Linha: " & FORMAT$(uA, "000000") & $CRLF & _
                                                                "combina��es: " & FORMAT$(quantidade_de_combinacoes, "000000000")
        'DIALOG DOEVENTS
        '
        ' Carregar arranjo
        PARSE strLinhas(uA), strCampos(), ANY ";"

        FOR uB = 4 TO uQuantidade_de_Campos
            uBolas(uB-3) = VAL(strCampos(uB))
        NEXT uB


        ' Aqui iremos fazer uma altera��o pois para o jogo 'DUPLASENA' para o mesmo concurso
        ' h� 2 sorteios
        ' Classifica o arranjo
        ' Como h� um validador de quantidade de campos, n�o precisamos nos preocupar, que est� tudo ok
        ' Sempre que for, 'DUPLASENA', sempre haver� 12 campos para as bolas.

        IF strCampos(1) = "DUPLASENA" THEN
            ARRAY SORT uBolas() FOR 6
            ARRAY SORT uBolas(7) FOR 6
        ELSE
            ARRAY SORT uBolas()
        END IF


        ' Declara as vari�veis que ser�o utilizados no loop for
        LOCAL uColuna1, uColuna2, uColuna3, uColuna4, uColuna5, uColuna6, uColuna7 AS LONG
        LOCAL uColuna8, uColuna9, uColuna10, uColuna11, uColuna12, uColuna13, uColuna14 AS LONG
        LOCAL uColuna15, uColuna16, uColuna17, uColuna18, uColuna19, uColuna20 AS LONG



        ' A vari�vel uLimite, indica a quantidade de bolas a combinar.
        LOCAL uLimite AS LONG

        IF uPararInsert = 1 THEN
            EXIT FOR
        END IF

        ' Aqui iremos fazer uma gambiarra, pois para o jogo 'DUPLASENA' s�o sorteados 2 jogos.
        ' No arquivo, as bolas est�o disposta conforme descrito abaixo:
        ' as 6 primeiras bolas, correspondem ao sorteio 1, as 6 �ltimas bolas, correspondem ao sorteio 2
        ' Ent�o, iremos fazer um loop
        LOCAL uSorteio AS LONG
        uSorteio = 1
        IF UCASE$(strCampos(1)) = "DUPLASENA" THEN
            uSorteio = 2
        ELSE
            uSorteio = 1
        END IF

        uQuantidade_de_Bolas = uQuantidade_de_Campos - 3

        LOCAL uColunaInicial AS LONG
        DO WHILE uSorteio <> 0
            IF uSorteio = 2 THEN
                uColunaInicial = 7
                ' Cria um string com todos os n�meros concatenados
                strBolas_Conjunto = ""
                ' Neste caso, somente esta situa��o ocorrer� para a duplasena
                FOR uB = 7 TO 12
                    strBolas_Conjunto += "_" + FORMAT$(uBolas(uB), "00")
                NEXT uB
            ELSE
                uColunaInicial = 1

                ' Cria um string com todos os n�meros concatenados
                strBolas_Conjunto = ""
                IF UCASE$(strCampos(1)) = "DUPLASENA" THEN
                    ' Se o jogo � 'DUPLASENA' devemos considerar a quantidade de bolas quando o sorteio for '1'
                    ' igual a 6, pois ser� ir� indicar 12
                    uQuantidade_de_Bolas = 6

                    FOR uB = 1 TO 6
                        strBolas_Conjunto += "_" + FORMAT$(uBolas(uB), "00")
                    NEXT uB
                ELSE
                    FOR uB = 1 TO uQuantidade_de_Bolas
                        strBolas_Conjunto += "_" + FORMAT$(uBolas(uB), "00")
                    NEXT uB
                END IF

            END IF


            FOR uColuna1 = uColunaInicial TO uQuantidade_de_Bolas
                uIndice(1) = uColuna1
                uLimite = 1
                GOSUB INSERIR_COMBINACAO_LOTERIA
                '
                FOR uColuna2 = uColuna1 + 1 TO uQuantidade_de_Bolas
                    uIndice(2) = uColuna2
                    uLimite = 2
                    GOSUB INSERIR_COMBINACAO_LOTERIA
                    '
                    FOR uColuna3 = uColuna2 + 1 TO uQuantidade_de_Bolas
                        uIndice(3) = uColuna3
                        uLimite = 3
                        GOSUB INSERIR_COMBINACAO_LOTERIA
                        '
                        FOR uColuna4 = uColuna3 + 1 TO uQuantidade_de_Bolas
                            uIndice(4) = uColuna4
                            uLimite = 4
                            GOSUB INSERIR_COMBINACAO_LOTERIA
                            '
                            FOR uColuna5 = uColuna4 + 1 TO uQuantidade_de_Bolas
                                uIndice(5) = uColuna5
                                uLimite = 5
                                GOSUB INSERIR_COMBINACAO_LOTERIA
                                '
                                FOR uColuna6 = uColuna5 + 1 TO uQuantidade_de_Bolas
                                    uIndice(6) = uColuna6
                                    uLimite = 6
                                    GOSUB INSERIR_COMBINACAO_LOTERIA
                                    '                              '
                                    FOR uColuna7 = uColuna6 + 1 TO uQuantidade_de_Bolas
                                        uIndice(7) = uColuna7
                                        uLimite = 7
                                        GOSUB INSERIR_COMBINACAO_LOTERIA
                                        '
                                        FOR uColuna8 = uColuna7 + 1 TO uQuantidade_de_Bolas
                                            uIndice(8) = uColuna8
                                            uLimite = 8
                                            GOSUB INSERIR_COMBINACAO_LOTERIA
                                            '                                         '
                                            FOR uColuna9 = uColuna8 + 1 TO uQuantidade_de_Bolas
                                                uIndice(9) = uColuna9
                                                uLimite = 9
                                                GOSUB INSERIR_COMBINACAO_LOTERIA
                                                '
                                                FOR uColuna10 = uColuna9 + 1 TO uQuantidade_de_Bolas
                                                    uIndice(10) = uColuna10
                                                    uLimite = 10
                                                    GOSUB INSERIR_COMBINACAO_LOTERIA
                                                    '
                                                    FOR uColuna11 = uColuna10 + 1 TO uQuantidade_de_Bolas
                                                        uIndice(11) = uColuna11
                                                        uLimite = 11
                                                        GOSUB INSERIR_COMBINACAO_LOTERIA
                                                        '
                                                        FOR uColuna12 = uColuna11 + 1 TO uQuantidade_de_Bolas
                                                            uIndice(12) = uColuna12
                                                            uLimite = 12
                                                            GOSUB INSERIR_COMBINACAO_LOTERIA
                                                            '
                                                            FOR uColuna13 = uColuna12 + 1 TO uQuantidade_de_Bolas
                                                                uIndice(13) = uColuna13
                                                                uLimite = 13
                                                                GOSUB INSERIR_COMBINACAO_LOTERIA
                                                                '
                                                                FOR uColuna14 = uColuna13 + 1 TO uQuantidade_de_Bolas
                                                                    uIndice(14) = uColuna14
                                                                    uLimite = 14
                                                                    GOSUB INSERIR_COMBINACAO_LOTERIA
                                                                    '
                                                                    FOR uColuna15 = uColuna14 + 1 TO uQuantidade_de_Bolas
                                                                        uIndice(15) = uColuna15
                                                                        uLimite = 15
                                                                        GOSUB INSERIR_COMBINACAO_LOTERIA

                                                                        FOR uColuna16 = uColuna15 + 1 TO uQuantidade_de_Bolas
                                                                            uIndice(16) = uColuna16
                                                                            uLimite = 16
                                                                            GOSUB INSERIR_COMBINACAO_LOTERIA

                                                                            FOR uColuna17 = uColuna16 + 1 TO uQuantidade_de_Bolas
                                                                                uIndice(17) = uColuna17
                                                                                uLimite = 17
                                                                                GOSUB INSERIR_COMBINACAO_LOTERIA

                                                                                FOR uColuna18 = uColuna17 + 1 TO uQuantidade_de_Bolas
                                                                                    uIndice(18) = uColuna18
                                                                                    uLimite = 18
                                                                                    GOSUB INSERIR_COMBINACAO_LOTERIA


                                                                                    FOR uColuna19 = uColuna18 + 1 TO uQuantidade_de_Bolas
                                                                                        uIndice(19) = uColuna19
                                                                                        uLimite = 19
                                                                                        GOSUB INSERIR_COMBINACAO_LOTERIA

                                                                                        FOR uColuna20 = uColuna19 + 1 TO uQuantidade_de_Bolas
                                                                                            uIndice(20) = uColuna20
                                                                                            uLimite = 20
                                                                                            GOSUB INSERIR_COMBINACAO_LOTERIA

                                                                                        NEXT uColuna20
                                                                                    NEXT uColuna19
                                                                                NEXT uColuna18
                                                                            NEXT uColuna17
                                                                        NEXT uColuna16
                                                                    NEXT uColuna15
                                                                NEXT uColuna14
                                                            NEXT uColuna13
                                                        NEXT uColuna12
                                                    NEXT uColuna11
                                                NEXT uColuna10
                                            NEXT uColuna9
                                        NEXT uColuna8
                                    NEXT uColuna7
                                NEXT uColuna6
                            NEXT uColuna5
                        NEXT uColuna4
                    NEXT uColuna3
                NEXT uColuna2
            NEXT uColuna1

            uSorteio -= 1
        LOOP

    NEXT


    FLUSH
    CLOSE

    DIM strTempo AS STRING
    DIM uSinal AS LONG
    DIM uDia AS LONG
    DIM uHora AS LONG
    DIM uMinuto AS LONG
    DIM uSegundo AS LONG
    DIM umiSegundo AS QUAD
    DIM uTick AS QUAD

    objTempo_Final.Now()
    objTempo_Final.TimeDiff(objTempo_Inicial, uSinal, uDia, uHora, uMinuto, uSegundo, uMiSegundo, uTick)

    strTempo  = "Executado com sucesso!!!" + $CRLF
    strTempo += FORMAT$(quantidade_de_combinacoes, "###,#")  + " combina��es geradas." + $CRLF
    strTempo += FORMAT$(quantidade_de_combinacoes / (uLinhas-1), "###,#") + " combina��es por jogo." + $CRLF
    strTempo += "Data_Hora_Inicial: "
    strTempo += FORMAT$(objTempo_Inicial.Day, "00") + "/"
    strTempo += FORMAT$(objTempo_Inicial.Month, "00") + "/"
    strTempo += FORMAT$(objTempo_Inicial.Year, "0000") + " "
    strTempo += FORMAT$(objTempo_Inicial.Hour, "00") + "h,"
    strTempo += FORMAT$(objTempo_Inicial.Minute, "00") + "min,"
    strTempo += FORMAT$(objTempo_Inicial.Second, "00") + "sec,"
    strTempo += FORMAT$(objTempo_Inicial.mSecond, "000") + "msec,"
    strTempo += FORMAT$(objTempo_Inicial.Tick, "000") + "t." + $CRLF
    strTempo += "Data_Hora_Final: "
    strTempo += FORMAT$(objTempo_Final.Day, "00") + "/"
    strTempo += FORMAT$(objTempo_Final.Month, "00") + "/"
    strTempo += FORMAT$(objTempo_Final.Year, "0000") + " "
    strTempo += FORMAT$(objTempo_Final.Hour, "00") + "h,"
    strTempo += FORMAT$(objTempo_Final.Minute, "00") + "min,"
    strTempo += FORMAT$(objTempo_Final.Second, "00") + "sec,"
    strTempo += FORMAT$(objTempo_Final.mSecond, "000") + "msec,"
    strTempo += FORMAT$(objTempo_Final.Tick, "000") + "t." + $CRLF + $CRLF
    strTempo += "Tempo Decorrido: "
    strTempo += FORMAT$(uDia, "00") + "d, "
    strTempo += FORMAT$(uHora, "00") + "h, "
    strTempo += FORMAT$(uMinuto, "00") + "min, "
    strTempo += FORMAT$(uSegundo, "00") + "sec, "
    strTempo += FORMAT$(uMiSegundo, "00") + "msec, "
    strTempo += FORMAT$(uTick, "00") + "tick."

    strSucesso = strTempo


    EXIT FUNCTION

INSERIR_COMBINACAO_LOTERIA:

    ' Para o banco n�o ficar muito pesado evitaremos, pegar combina��es maiores que 5 bolas para o jogo Lotofacil.


    ' Tamb�m evitaremos que LOTOMANIA, fique muito grande, colocaremos o limite de at� 3.
    IF strCampos(1) = "LOTOMANIA" AND uLimite > 2 THEN
        RETURN
    END IF

    ' Vamos evitar que INTRALOT_KENO_MINAS fique muito grande, vamos limitar somente at� 1 bola.
    IF strCampos(1) = "INTRALOT_KENO_MINAS" AND uLimite > 1 THEN
        RETURN
    END IF


    ' Vamos contar quantas combina��es foram executadas
    quantidade_de_combinacoes += 1
    'Dialog doevents



    ' A cada 7 mil combina��es iremos gerar, um novo arquivo de insert
    IF ((quantidade_de_combinacoes MOD 229370 = 0) AND TRIM$(strCampos(1)) = "LOTOFACIL") OR _
        (quantidade_de_combinacoes MOD 25201 = 0 AND TRIM$(strCampos(1)) = "DUPLASENA")   OR _
        (quantidade_de_combinacoes MOD 337500 = 0 AND TRIM$(strCampos(1)) = "LOTOMANIA") OR _
        (quantidade_de_combinacoes MOD 500000 = 0 AND TRIM$(strCampos(1)) = "INTRALOT_MINAS_5") THEN

        ' Acima como o limite � de at� 3, � poss�vel gerar 1350 combina��es.
        ' ent�o, 337500 � igual a 250 * 1350




        ' Na DuplaSena, cada 6 bolas, em ordem crescente, gera 63 combina��es, como h� 2 sorteios, em um �nico concurso
        ' conclue-se que h� 126 combina��es, ent�o para evitarmos que o arquivo fique gigante, dividimos as informa��es
        ' em mais de um arquivo, para a DuplaSena, iremos considerar 200 concursos completos para gerar um novo arquivo.
        ' Ent�o, 200 concurso vezes 126 combina��es � igual a 25200 linhas, ent�o quando se chega na linha 25201 � um novo concurso
        ' devemos gerar um novo arquivo, e come�ar neste.


        ' Na lotofacil � gerado 32767 combina��es por jogo, ent�o a cada 7 jogos, iremos fazer um 'commit' para inserir no banco
        'PRINT #uInsert_Jogo_Bolas, "commit;" & $CRLF
        FLUSH

        DIM strArquivo AS STRING

        strArquivo = FILENAME$(uInsert_Jogo_Bolas)
        IF strArquivo <> "" THEN
            DIM uPosicao AS LONG
            'MsgBox strArquivo
            uPosicao = INSTR(strArquivo, FORMAT$(uContar_Grupo, "000000") & ".sql")
            IF uPosicao = 0 THEN
                MSGBOX "Arquivo n�o termina em '000001.sql'.", %MB_ICONERROR
                uPararInsert = 1
                CLOSE
                EXIT FUNCTION
            END IF
            uContar_Grupo += 1
            strArquivo = MID$(strArquivo, 1, uPosicao - 1) & FORMAT$(uContar_Grupo, "000000") & ".sql"
            CLOSE uInsert_Jogo_Bolas
            OPEN strArquivo FOR APPEND AS uInsert_Jogo_Bolas
        END IF
    END IF


    ' Incluso em 04/07/2015
    ' A vari�ve abaixo serve para indicar quais os campos na tabela devemos inserir o valor '1'
    ' Se por exemplo, a combina��o for dos n�meros _05_07_12
    ' A vari�vel ter� este valor: ', num_5, num_7, num_12', pois assim, estaremos indicando quais campos
    ' ser�o atualizados.
    DIM strNumeros_Insert AS STRING
    strNumeros_Insert = ""


    DIM strBolas_Combinadas AS STRING

    strBolas_Combinadas = ""

    LOCAL lngA AS LONG
    FOR lngA = 1 TO uLimite
        strBolas_Combinadas = strBolas_Combinadas & "_" & FORMAT$(uBolas(uIndice(lngA)), "00")

        strNumeros_Insert += ", num_" + FORMAT$(uBolas(uIndice(lngA)))
    NEXT

    ' Gerar arquivo para inserir na tabela 'JOGO_BOLAS'
    DIM strInsert_Jogo_Bolas AS STRING

    strInsert_Jogo_Bolas = "Insert Into ltk.jogo_bolas (jogo_tipo, concurso, sorteio, data, bolas_combinadas, bolas_combinadas_qt, bolas_combinadas_base, bolas_combinadas_base_qt"

    ' Acrescentado em 04/07/2015
    ' Abaixo, indicaremos quais os campos que ser�o atualizados.
    strInsert_Jogo_Bolas += strNumeros_Insert

    ' Devemos fechar os campos com ')' par�nteses.
    strInsert_Jogo_Bolas += ")"


    strInsert_Jogo_Bolas += " VALUES ("

    strInsert_Jogo_Bolas += "'" & strCampos(1) + "', "
    strInsert_Jogo_Bolas += "'" & strCampos(2) + "', "
    strInsert_Jogo_Bolas += "'" & FORMAT$(uSorteio) & "', "
    strInsert_Jogo_Bolas += "'" & strCampos(3) & "', "
    strInsert_Jogo_Bolas += "'" & strBolas_Combinadas & "', "
    strInsert_Jogo_Bolas += "'" & FORMAT$(uLimite) & "', "
    strInsert_Jogo_Bolas += "'" & strBolas_Conjunto & "', "
    strInsert_Jogo_Bolas += "'" & FORMAT$(uQuantidade_de_Bolas)  & "'"      ' Observe que nesta linha n�o h� v�rgula no final, pois, a v�rgula est� inserida
                                                                            ' na vari�vel 'strNumeros_Insert'

    ' Aqui, iremos repetir v�rios '1' conforme a o valor de 'uLimite'.
    strInsert_Jogo_Bolas += REPEAT$(uLimite, ", '1'")

    ' Devemos, inserir o par�nteses referente ao final do insert ')' + um CARRIAGE RETURN + LINEFEED
    strInsert_Jogo_Bolas += "); "



    IF uCommit = 1 THEN
        strInsert_Jogo_Bolas += "commit;" + $CRLF
    END IF



    TRY
        PRINT #uInsert_Jogo_Bolas, strInsert_Jogo_Bolas

    CATCH
        strErro = ERROR$(ERR)
        CLOSE
        FUNCTION = -1
        EXIT FUNCTION
    END TRY
    RETURN

END FUNCTION

GLOBAL uLinhaInsert AS QUAD
GLOBAL uCombinacoesInsert AS LONG

GLOBAL uMinutoInsert, uSegundoInsert AS LONG
GLOBAL uMiSegundoInsert AS QUAD

GLOBAL uHoraInsertDecorrido, uMinutoInsertDecorrido, uSegundoInsertDecorrido AS LONG
GLOBAL uMiSegundoInsertDecorrido, uTickInsertDecorrido AS QUAD

GLOBAL objRelatorioInicio AS IPOWERTIME


FUNCTION Gerar_Insert_Where_Todas_Combinacoes(BYVAL strArquivoOrigem AS STRING, BYVAL strPastaDestino AS STRING, BYREF strErro AS STRING) AS LONG
    ' Vamos verificar se arquivo de origem existe
    IF ISFALSE(ISFILE(strArquivoOrigem)) THEN
        strErro = "Arquivo '" & strArquivoOrigem & "' n�o existe."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos verificar se pasta existe
    IF ISFALSE(ISFOLDER(strPastaDestino)) THEN
        strErro = "A pasta '" & strPastaDestino & "' n�o existe."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    DIM uArquivoOrigem AS LONG
    DIM uArquivoInsert AS LONG
    DIM uArquivoWhere AS LONG

    ' Vamos tentar detectar que tipo de jogo �, abrindo-o e lendo o layout da primeira linha
    uArquivoOrigem = FREEFILE
    OPEN strArquivoOrigem FOR INPUT AS #uArquivoOrigem BASE = 1
    IF ERR <> 0 THEN
        strErro = ERROR$(ERR)
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Se o arquivo est� vazio quer dizer que o arquivo j� est� no fim
    IF ISTRUE(EOF(uArquivoOrigem)) THEN
        strErro = "Arquivo est� vazio."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos ler a primeira linha para definir as vari�veis que ser�o utilizadas no jogo
    DIM strLinha AS STRING
    LINE INPUT #uArquivoOrigem, strLinha

    ' Vari�vel arranjo para guardar todos os campos
    DIM strCampos() AS STRING
    DIM uCampos AS LONG

    uCampos = PARSECOUNT(strLinha, ANY "_;")
    REDIM strCampos(1 TO uCampos)

    ' Verificar se uCampos � maior que 4 campos
    IF uCampos < 4 THEN
        strErro = "Layout de arquivo inv�lido
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos preencher o arranjo
    PARSE strLinha, strCampos(), ANY "_;"

    ' Verificar se jogo � v�lido
    IF  strCampos(1) <> "QUINA" AND _
        strCampos(1) <> "MEGASENA" AND _
        strCampos(1) <> "LOTOFACIL" AND _
        strCampos(1) <> "LOTOMANIA" THEN
            strErro = "Jogo '" & strCampos(1) & "' inv�lido."
            FUNCTION = -1
            EXIT FUNCTION
    END IF

    ' Verificar se o campo 2 e 4 contenha, respectivamente, as palavras
    ' 'COM' e 'NUMEROS'
    IF (strCampos(2) + strCampos(4)) <> "COMNUMEROS" THEN
        strErro = "Campo 2, deve ter a palavra 'COM' e campo 4, a palavra 'NUMEROS'."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Vamos verificar a quantidade de bolas apostadas
    ' que fica no campo 3
    DIM uBolasApostadas AS LONG
    uBolasApostadas = VAL(strCampos(3))

    ' De acordo com o jogo, vamos verificar a quantidade de bolas apostadas
    SELECT CASE strCampos(1)
        CASE "QUINA"
            IF uBolasApostadas < 5 OR uBolasApostadas > 7 THEN
                strErro = "No campo 3, a quantidade informada '" & strCampos(3) & _
                "' n�o coincide com a quantidade de n�meros a apostar para o jogo 'QUINA'."
                FUNCTION = -1
                EXIT FUNCTION
            END IF

        CASE "MEGASENA"
            IF uBolasApostadas < 6 OR uBolasApostadas > 15 THEN
                strErro = "No campo 3, a quantidade informada '" & strCampos(3) & _
                "' n�o coincide com a quantidade de n�meros a apostar para o jogo 'MEGASENA'."
                FUNCTION = -1
                EXIT FUNCTION
            END IF

        CASE "LOTOFACIL
            IF uBolasApostadas < 15 OR uBolasApostadas > 18 THEN
                strErro = "No campo 3, a quantidade informada '" & strCampos(3) & _
                "' n�o coincide com a quantidadde de n�meros a apostar para o jogo 'LOTOFACIL'."
                FUNCTION = -1
                EXIT FUNCTION
            END IF

        CASE "LOTOMANIA"
            IF uBolasApostadas <> 50 THEN
                strErro = "No campo 3, a quantidade informada '" & strCampos(3) & _
                "' n�o coincide com a quantidade de n�meros a apostar para o jogo 'LOTOMANIA'."
                FUNCTION = -1
                EXIT FUNCTION
            END IF
    END SELECT

    ' Vamos verificar se a quantidade de campos coincide com a quantidade de bolas apostadas
    IF uCampos - 4 <> uBolasApostadas THEN
        strErro = "A quantidade de bolas n�o corresponde � quantidade de n�meros apostadas."
        FUNCTION = -1
        EXIT FUNCTION
    END IF

    ' Para cada tipo de jogo, h� n�meros diferentes para o limite inferior e superior
    DIM uLimiteInferior  AS LONG
    DIM uLimiteSuperior AS LONG

    SELECT CASE strCampos(1)
        CASE "QUINA"
            uLimiteInferior = 1
            uLimiteSuperior = 80
        CASE "MEGASENA"
            uLimiteInferior = 1
            uLimiteSuperior = 60
        CASE "LOTOFACIL"
            uLimiteInferior = 1
            uLimiteSuperior = 25
        CASE "LOTOMANIA"
            uLimiteInferior = 0
            uLimiteSuperior = 99
        CASE "DUPLASENA"
            uLimiteInferior = 1
            uLimiteSuperior = 50
    END SELECT

    '
    ' Vamos abrir o arquivo para grava��o posteriormente, mas primeiro criar um arquivo
    '
    DIM objTempo AS IPOWERTIME
    LET objTempo = CLASS "POWERTIME"
    objTempo.Now()

    DIM strTempo AS STRING
    strTempo =  "_" & FORMAT$(objTempo.Day, "00") & "_" & FORMAT$(objTempo.Month, "00") & "_" & FORMAT$(objTempo.Year, "00") & _
                "_" & FORMAT$(objTempo.Hour, "00") & "_" & FORMAT$(objTempo.Minute, "00") & "_" & FORMAT$(objTempo.Second, "00") & _
                "_" & FORMAT$(objTempo.mSecond, "000")

    DIM uWhere AS LONG
    DIM uInsert AS LONG
    DIM strWhere AS STRING
    DIM strInsert AS STRING

    ' Vamos verificar a pasta de destino, se tem o caractere '\'
    IF RIGHT$(strPastaDestino, 1) <> "\" THEN
        strPastaDestino &= "\"
    END IF

    'uWhere = FREEFILE
    uInsert = FREEFILE
    'strWhere = strPastaDestino &  "WHERE_" & strCampos(1) & "_COM_" & strCampos(3) & "_NUMEROS_" & strTempo & ".txt"

    strInsert = strPastaDestino &  "INSERT_" & strCampos(1) & "_COM_" & strCampos(3) & "_NUMEROS_" & strTempo & ".txt"

    'OPEN strWhere FOR OUTPUT AS #uWhere
    'IF ERR <> 0 THEN
    '   strErro = ERROR$(ERR)
    '    FUNCTION = -1
    '    EXIT FUNCTION
    'END IF

    OPEN strInsert FOR OUTPUT AS #uInsert
    IF ERR <> 0 THEN
        strErro = ERROR$(ERR)
        FUNCTION = -1
        EXIT FUNCTION
    END IF
    '
    '---------------------------------------------------------------------------------------------------------------------------
    '

    ' Como verificamos todos os campos, vamos definir as vari�veis
    DIM strJogo AS STRING           ' Por exemplo, QUINA
    DIM strJogoSubTipo AS STRING    ' Por exemplo, QUINA_COM_5_NUMEROS
    DIM uLimite AS LONG             ' Quantidade de bolas apostadas

    DIM uIndice() AS LONG           ' Guarda as posi��es dos �ndice do arranjo
    DIM uBolas() AS LONG            ' Guarda todas as bolas

    uCombinacoesInsert = 0

    strJogo = strCampos(1)
    strJogoSubTipo = strCampos(1) & "_COM_" & strCampos(3) & "_NUMEROS"
    uBolasApostadas = VAL(strCampos(3))
    REDIM uIndice(1 TO uBolasApostadas)
    REDIM uBolas(1 TO uBolasApostadas)


    LOCAL uColuna1, uColuna2, uColuna3, uColuna4, uColuna5, uColuna6, uColuna7 AS LONG
    LOCAL uColuna8, uColuna9, uColuna10, uColuna11, uColuna12, uColuna13, uColuna14 AS LONG
    LOCAL uColuna15 AS LONG


    ' Vamos definir vari�vel que guardar� o hor�rio de in�cio
    'Dim objRelatorioInicio as ipowertime
    LET objRelatorioInicio = CLASS "POWERTIME"
    objRelatorioInicio.Now()

    uLinhaInsert = 0
    ' Vamos para o in�cio do arquivo
    SEEK #uArquivoOrigem, 1
    DO WHILE ISFALSE EOF(uArquivoOrigem)
        ' Ler linha
        LINE INPUT #uArquivoOrigem, strLinha
        IF ERR <> 0 THEN
            strErro = ERROR$(ERR)
            CLOSE
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        uLinhaInsert += 1

        ' Cada campo � separado por '_' e ';'
        ' por exemplo, QUINA_COM_5_NUMEROS;1;2;3;4;5
        ' Vamos validar sempre cada linha lida
        ' Como abrirmos o arquivo anterior, todas os 4 primeiros campos tem que sempre
        ' coincidir, pois isto quer dizer, que refere-se ao mesmo jogo
        DIM uCamposQuantidade AS LONG
        uCamposQuantidade = PARSECOUNT(strLinha, ANY "_;")
        IF uCamposQuantidade <> uCampos THEN
            strErro = "Erro, Linha: " & FORMAT$(uLinhaInsert) & ", quantidade de campos n�o coincide com as quantidade de campos anteriores
            CLOSE
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        ' Colocar campos na vari�vel de arranjo
        PARSE strLinha, strCampos(), ANY "_;"

        ' Verificar se refere-se ao mesmo jogo
        IF ISFALSE(strCampos(1) & "_" & strCampos(2) & "_" & strCampos(3)  & "_NUMEROS" = strJogoSubTipo) THEN
            strErro =   "Erro, linha: " & FORMAT$(uLinhaInsert) & ", o jogo '" & strCampos(1) & "_" & strCampos(2) & "_" & strCampos(3) & _
                        "' n�o se refere ao jogo '" & strJogoSubTipo
            CLOSE
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        ' Verificar se a quantidade de bolas � igual � quantidade de n�meros apostados
        ' Quantidade de n�meros apostados, fica no campo 3 do arranjo 'strCampos'
        IF VAL(strCampos(3)) <> uCamposQuantidade - 4 THEN
            strErro =   "Erro, linha: " & FORMAT$(uLinhaInsert) & ", a quantidade de n�meros apostados: '" & strCampos(3) & _
                        "' n�o coincide com a quantidade de n�meros no campos correspondente desta linha."
            CLOSE
            FUNCTION = -1
            EXIT FUNCTION
        END IF


        DIM strBolas_Combinadas_Conjunto AS STRING
        strBolas_Combinadas_Conjunto = ""

        ' Se chegamos at� aqui, quer dizer, que todos os campos est�o ok
        ' Ent�o vamos atribuir os valores dos campos � vari�vel uBolas()
        DIM uA AS LONG
        FOR uA = 1 TO uBolasApostadas
            uBolas(uA) = VAL(strCampos(4 + uA))
            strBolas_Combinadas_Conjunto += "_" & FORMAT$(uBolas(uA), "00")
        NEXT uA

        ' Vamos verificar se as bolas est�o dentro da faixa
        DIM uVerificarFaixa AS LONG
        ARRAY SCAN uBolas(), < uLimiteInferior , TO uVerificarFaixa
        '
        ' Se encontramos algo, fora da faixa, uVerificarFaixa ser� diferente de zero '0'
        IF uVerificarFaixa <> 0 THEN
            strErro =   "Erro, linha: " & FORMAT$(uLinhaInsert) & ", o valor " & FORMAT$(uBolas(uVerificarFaixa)) & " do campo " & _
                        FORMAT$(uVerificarFaixa + 4) & " est� fora da faixa, a faixa permitida �: " & FORMAT$(uLimiteInferior) & _
                        "-" & FORMAT$(uLimiteSuperior)
            FUNCTION = -1
            EXIT FUNCTION
        END IF

        ARRAY SCAN uBolas(), > uLimiteSuperior , TO uVerificarFaixa
        '
        ' Se encontramos algo, fora da faixa, uVerificarFaixa ser� diferente de zero '0'
        IF uVerificarFaixa <> 0 THEN
            strErro =   "Erro, linha: " & FORMAT$(uLinhaInsert) & ", o valor " & FORMAT$(uBolas(uVerificarFaixa)) & " do campo " & _
                        FORMAT$(uVerificarFaixa + 4) & " est� fora da faixa, a faixa permitida �: " & FORMAT$(uLimiteInferior) & _
                        "-" & FORMAT$(uLimiteSuperior)
            FUNCTION = -1
            EXIT FUNCTION
        END IF


        ' Vamos come�ar a gerar todas as combina��es poss�veis para o jogo especificado
        ' Primeiro vamos classificar o arranjo
        ARRAY SORT uBolas()

        ' O loop abaixo gera todas as combina��es para um conjunto de n�meros, por exemplo, para os n�meros:
        '               *********************
        ' INDICE -->    * 1   *     2    *  3 *  4  *
        ' uBolas -->    * 5     6    *  7 *  8  *
        ' Ent�o fica
        '               * 5                 *
        '               * 5 6
        '               * 5 6 7
        '               * 5 6 7 8
        '               * 5 6 8
        '               * 5 7
        '               * 5 7 8
        '               * 6
        '               * 6 7
        '               * 6 7 8
        '               * 6 8
        '               * 7
        '               * 7 8
        '               * 8

        IF uPararInsert = 1 THEN
            strErro = "Processamento cancelado por solicitacao do usu�rio"
            CLOSE
            FUNCTION = -1
            EXIT FUNCTION
        END IF


        DIM objAgora AS IPOWERTIME
        LET objAgora = CLASS "POWERTIME"
        objAgora.now()

        IF uLinhaInsert = 2 THEN
            MSGBOX "uCombinacoesInsert = " & FORMAT$(uCombinacoesInsert, "###,")
            EXIT FUNCTION
        END IF

        FOR uColuna1 = 1 TO uBolasApostadas
            uIndice(1) = uColuna1
            Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 1)
            '
            FOR uColuna2 = uColuna1 + 1 TO uBolasApostadas
                uIndice(2) = uColuna2
                Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 2)
                '
                FOR uColuna3 = uColuna2 + 1 TO uBolasApostadas
                    uIndice(3) = uColuna3
                    Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 3)
                    '
                    FOR uColuna4 = uColuna3 + 1 TO uBolasApostadas
                        uIndice(4) = uColuna4
                        Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 4)
                        '
                        FOR uColuna5 = uColuna4 + 1 TO uBolasApostadas
                            uIndice(5) = uColuna5
                            Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 5)
                            '
                            FOR uColuna6 = uColuna5 + 1 TO uBolasApostadas
                                uIndice(6) = uColuna6
                                Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 6)
                                '
                                FOR uColuna7 = uColuna6 + 1 TO uBolasApostadas
                                    uIndice(7) = uColuna7
                                    Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 7)
                                    '
                                    FOR uColuna8 = uColuna7 + 1 TO uBolasApostadas
                                        uIndice(8) = uColuna8
                                        Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 8)
                                        '
                                        FOR uColuna9 = uColuna8 + 1 TO uBolasApostadas
                                            uIndice(9) = uColuna9
                                            Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 9)
                                            '
                                            FOR uColuna10 = uColuna9 + 1 TO uBolasApostadas
                                                uIndice(10) = uColuna10
                                                Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 10)
                                                '
                                                FOR uColuna11 = uColuna10 + 1 TO uBolasApostadas
                                                    uIndice(11) = uColuna11
                                                    Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 11)
                                                    '
                                                    FOR uColuna12 = uColuna11 + 1 TO uBolasApostadas
                                                        uIndice(12) = uColuna12
                                                        Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 12)
                                                        '
                                                        FOR uColuna13 = uColuna12 + 1 TO uBolasApostadas
                                                            uIndice(13) = uColuna13
                                                            Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 13)
                                                            '
                                                            FOR uColuna14 = uColuna13 + 1 TO uBolasApostadas
                                                                uIndice(14) = uColuna14
                                                                Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 14)
                                                                '
                                                                FOR uColuna15 = uColuna14 + 1 TO uBolasApostadas
                                                                    uIndice(15) = uColuna15
                                                                    Gerar_Arquivo_Insert_Where_Todas_Combinacoes(strJogo, strJogoSubTipo, strBolas_Combinadas_Conjunto, uBolasApostadas, uInsert, uWhere, uBolas(), uIndice(), 8)
                                                                NEXT uColuna15
                                                            NEXT uColuna14
                                                        NEXT uColuna13
                                                    NEXT uColuna12
                                                NEXT uColuna11
                                            NEXT uColuna10
                                        NEXT uColuna9
                                    NEXT uColuna8
                                NEXT uColuna7
                            NEXT uColuna6
                        NEXT uColuna5
                    NEXT uColuna4
                NEXT uColuna3
            NEXT uColuna2
        NEXT uColuna1

        LOCAL uSinal AS LONG
        LOCAL objDepois AS IPOWERTIME
        objDepois = CLASS "POWERTIME"
        objDepois.now()
        objDepois.TimeDiff(objAgora, uSinal, BYVAL 0, BYVAL 0, uMinutoInsert, uSegundoInsert, uMiSegundoInsert)
    LOOP

    LOCAL objRelatorioFinal AS IPOWERTIME
    LET objRelatorioFinal = CLASS "POWERTIME"
    LOCAL uSinalFinal AS LONG

    objRelatorioFinal.Now()
    objRelatorioFinal.TimeDiff(objRelatorioInicio, uSinalFinal, BYVAL 0, uHoraInsertDecorrido, uMinutoInsertDecorrido, _
                                uSegundoInsertDecorrido, uMiSegundoInsertDecorrido, uTickInsertDecorrido)

    DIM strFinal AS STRING
    strFinal =      "Tempo_Decorrido: " + FORMAT$(uHoraInsertDecorrido, "00") + "h," + FORMAT$(uMinutoInsertDecorrido, "00") + "min," + _
                    FORMAT$(uSegundoInsertDecorrido, "00") + "sec," + FORMAT$(uMiSegundoInsertDecorrido, "000") + "msec," + _
                    FORMAT$(uTickInsertDecorrido, "000") + "tick, "

    MSGBOX strFinal





END FUNCTION

'
' Grava no arquivo em um formato que podemos utilizar futuramente em um banco de dados
' os par�metros da fun��o s�o:
'
'   strJogo:                            O nome do jogo, por exemplo, "QUINA"
'   strBolas_Combinadas_Conjunto:       Um string que representa todas as bolas do conjunto
'   uInsert:                            Identificador utilizado pela senten�a OPEN para gravar o arquivo
'   uWhere:                             Identificador utilizado pela senten�a OPEN para gravar o arquivo
'   uBolas():                           Vari�vel arranjo que guarda todas as bolas
'   uIndice():                          Est� vari�vel indica em qual �ndice deve pegar no arranjo uBolas
'   uLimite:                            Vari�vel que indica quantos �tens ser� composto o subconjunto dos n�meros
'
SUB Gerar_Arquivo_Insert_Where_Todas_Combinacoes(BYVAL strJogo AS STRING,  BYVAL strJogoSubTipo AS STRING, _
                                                BYVAL strBolas_Combinadas_Conjunto AS STRING, BYVAL uBolasApostadas AS LONG, _
                                                BYVAL uInsert AS LONG, BYVAL uWhere AS LONG, _
                                                BYREF uBolas() AS LONG, BYREF uIndice() AS LONG, _
                                                BYVAL uLimite AS LONG)
    DIALOG DOEVENTS

    DIM lngA AS LONG
    DIM strBolas_Combinadas AS STRING

    uCombinacoesInsert += 1



    strBolas_Combinadas = ""
    FOR lngA = 1 TO uLimite
        strBolas_Combinadas = strBolas_Combinadas & "_" & FORMAT$(uBolas(uIndice(lngA)), "00")
    NEXT


    ' Vamos contabilizar quantidade de n�meros pares e impares
    LOCAL uQuantidade_de_Pares, uQuantidade_de_Impares AS LONG
    uQuantidade_de_Pares = 0
    uQuantidade_de_Impares = 0
    FOR lngA = 1 TO uLimite
        IF uBolas(uIndice(lngA)) MOD 2 = 0 THEN
            uQuantidade_de_Pares += 1
        ELSE
            uQuantidade_de_Impares += 1
        END IF
    NEXT

    ' Vamos contabilizar quantidade de quadrantes
    LOCAL uQuadrante1, uQuadrante2, uQuadrante3, uQuadrante4 AS LONG

    ' A faixa de quadrantes s�o:
    ' JOGO      QUADRANTE 1         QUADRANTE 2     QUADRANTE 3         QUADRANTE 4
    ' QUINA     DE 1 A 20           DE 21 A 40      DE 41 A 60          DE 61 A 80
    ' MEGASENA  DE 1 A 15           DE 16 A 30      DE 31 A 45          DE 46 A 60
    ' LOTOFACIL DE 1 A 6            DE 7 A 12       DE 13 A 18          DE 19 A 25

    FOR lngA = 1 TO uLimite
        LOCAL uNumero AS LONG
        uNumero = uBolas(uIndice(lngA))

        uQuadrante1 = 0
        uQuadrante2 = 0
        uQuadrante3 = 0
        uQuadrante4 = 0
        SELECT CASE strJogo
            CASE "QUINA"
                SELECT CASE uNumero
                    CASE 1 TO 20
                        uQuadrante1 = 1
                    CASE 21 TO 40
                        uQuadrante2 = 1
                    CASE 41 TO 60
                        uQuadrante3 = 1
                    CASE 61 TO 80
                        uQuadrante4 = 1
                END SELECT

            CASE "MEGASENA"
                SELECT CASE uNumero
                    CASE 1 TO 15
                        uQuadrante1 = 1
                    CASE 16 TO 30
                        uQuadrante2 = 1
                    CASE 31 TO 45
                        uQuadrante3 = 1
                    CASE 46 TO 60
                        uQuadrante4 = 1
                END SELECT

            CASE "LOTOFACIL"
                SELECT CASE uNumero
                    CASE 1 TO 6
                        uQuadrante1 = 1
                    CASE 7 TO 12
                        uQuadrante2 = 1
                    CASE 13 TO 18
                        uQuadrante3 = 1
                    CASE 19 TO 25
                        uQuadrante4 = 1
                END SELECT
        END SELECT
    NEXT

    ' Gerar Insert para colocar na tabela
    DIM strInsert AS STRING
    strInsert = "Insert into JOGO_BOLAS_COMBINADAS (JOGO_TIPO, JOGO_SUB_TIPO, BOLAS_COMBINADAS_BASE, BOLAS_COMBINADAS_BASE_QT, BOLAS_COMBINADAS, BOLAS_QUANTIDADE) VALUES ("
    strInsert &= "'" & strJogo & "', "
    strInsert &= "'" & strJogoSubTipo & "', "
    strInsert &= "'" & strBolas_Combinadas_Conjunto & "', "
    strInsert &= FORMAT$(uBolasApostadas) & ", "
    strInsert &= "'" & strBolas_Combinadas & "', "
    strInsert &= FORMAT$(uLimite) & ")"

    PRINT #uInsert, strInsert

    ' Gerar Select para verificar se j� foi inserido
    DIM strWhere AS STRING
    'strWhere =  "Select * from JOGO_BOLAS_COMBINADAS WHERE JOGO_TIPO = '" & strJogo & "' "
    'strWhere &= "AND JOGO_SUB_TIPO = '" & strJogoSubTipo & "' "
    'strWhere &= "AND BOLAS_COMBINADAS_BASE = '" & strBolas_Combinadas_Conjunto & "' "
    'strWhere &= "AND BOLAS_COMBINADAS = '" & strBolas_Combinadas & "'"
    'PRINT #uWhere, strWhere

    DIM strTexto AS STRING
    IF uAtualizarInsert = 1 THEN
        ' Se atualizaremos a tela, ent�o devemos, verificar o tempo decorrido
        DIM uSinal AS LONG
        DIM objRelatorioDecorrido AS IPOWERTIME
        objRelatorioDecorrido = CLASS "POWERTIME"
        objRelatorioDecorrido.Now()
        objRelatorioDecorrido.TimeDiff(objRelatorioInicio, uSinal, BYVAL 0, uHoraInsertDecorrido, uMinutoInsertDecorrido, _
                                        uSegundoInsertDecorrido, uMiSegundoInsertDecorrido, uTickInsertDecorrido)


        strTexto =  "Linha: " & FORMAT$(uLinhaInsert, "000000000,") & $CRLF & _
                    FORMAT$(uCombinacoesInsert, "000000000,") & " j� processadas. " & $CRLF & _
                    "Tempo_An�lise_Linha: " + FORMAT$(uMinutoInsert, "00") & "min," & FORMAT$(uSegundoInsert, "00") & "sec," & _
                    FORMAT$(uMiSegundoInsert, "00") & "msec " & $CRLF & _
                    "Tempo_Decorrido: " + FORMAT$(uHoraInsertDecorrido, "00") + "h," + FORMAT$(uMinutoInsertDecorrido, "00") + "min," + _
                    FORMAT$(uSegundoInsertDecorrido, "00") + "sec," + FORMAT$(uMiSegundoInsertDecorrido, "000") + "msec," + _
                    FORMAT$(uTickInsertDecorrido, "000") + "tick, " + _
                    "BOLAS_CONJUNTO: '" & $CRLF & _
                    strBolas_Combinadas_Conjunto & $CRLF & _
                    "', BOLAS_COMBINADAS: '" & $CRLF & _
                    strBolas_Combinadas & "'."
        CONTROL SET TEXT hJanela_Principal, %IDC_LOG_INSERT, strTexto
    END IF

END SUB

'------------------------------------------------------------------------------
'   ** Dialogs **
'------------------------------------------------------------------------------
FUNCTION Exibir_Janela_Principal(BYVAL hParent AS DWORD) AS LONG
    LOCAL lRslt AS LONG

#PBFORMS BEGIN DIALOG %IDD_DIALOG1->->
    LOCAL hDlg  AS DWORD

    DIALOG NEW hParent, "Analisador e Gerador - LTK", 209, 4, 563, 422, _
        %WS_POPUP OR %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR _
        %WS_MINIMIZEBOX OR %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME _
        OR %DS_3DLOOK OR %DS_NOFAILCREATE OR %DS_SETFONT, _
        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR, TO hDlg
    CONTROL ADD FRAME,    hDlg, %IDC_FRAME1, "Gerar todas as combina��es", 5, _
        5, 275, 95
    CONTROL ADD COMBOBOX, hDlg, %IDC_CMB_JOGO_TIPO_TODAS, , 10, 25, 95, 40, _
        %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR _
        %CBS_DROPDOWNLIST, %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR
    CONTROL ADD COMBOBOX, hDlg, %IDC_CMB_APOSTAR_COM_TODAS, , 115, 25, 100, _
        40, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR _
        %CBS_DROPDOWNLIST, %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR
    CONTROL ADD TEXTBOX,  hDlg, %IDC_TXT_PASTA_DESTINO_TODAS, _
        ".\PRM_GERADAS", 10, 50, 265, 15
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHK_ATUALIZAR_TODAS, "Verificar status " + _
        "da gera��o", 10, 85, 100, 10
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_GERAR_TODAS, "Gerar", 225, 80, 50, _
        15
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL1, "Jogo:", 10, 15, 75, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL2, "&Apostas com:", 115, 15, 75, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL3, "Pasta de destino (nome do " + _
        "arquivo gerado automaticamente):", 10, 40, 200, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LBL_LOG_TODAS, "", 10, 65, 265, 15
    CONTROL ADD FRAME,    hDlg, %IDC_FRAME2, "Gerar Insert e Where para todas " + _
        "as combina��es", 280, 5, 271, 115
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL4, "Nome do arquivo:", 286, 15, _
        100, 10
    CONTROL ADD TEXTBOX,  hDlg, %IDC_TXT_ARQUIVO_ORIGEM, "", 286, 25, 260, 13
    CONTROL ADD TEXTBOX,  hDlg, %IDC_TXT_PASTA_DESTINO_INSERT, "C:\TMP\", _
        286, 50, 260, 15
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL5, "&Pasta de destino", 286, 40, _
        100, 10
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHK_STATUS_INSERT, "Verificar status da " + _
        "gera��o", 285, 105, 100, 10
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_GERAR_INSERT, "Gerar", 495, 100, 50, _
        15
    CONTROL ADD LABEL,    hDlg, %IDC_LOG_INSERT, "", 285, 65, 260, 35
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_PARAR_TODAS, "&Parar", 170, 80, 50, _
        15, %WS_CHILD OR %WS_VISIBLE OR %WS_DISABLED OR %WS_TABSTOP OR _
        %BS_TEXT OR %BS_PUSHBUTTON OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT _
        OR %WS_EX_LTRREADING
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_PARAR_INSERT, "&Parar", 435, 100, _
        50, 15, %WS_CHILD OR %WS_VISIBLE OR %WS_DISABLED OR %WS_TABSTOP OR _
        %BS_TEXT OR %BS_PUSHBUTTON OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT _
        OR %WS_EX_LTRREADING
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHK_PAUSAR_TODAS, "PAUSA&R", 115, 84, _
        50, 10
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHK_PAUSAR_INSERT, "PAUSA&R", 385, 105, _
        50, 10
    CONTROL ADD FRAME,    hDlg, %IDC_FRAME3, "Gerar Insert de resultado de " + _
        "loteria", 5, 120, 545, 175
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL6, "Nome do arquivo ou filtro de " + _
        "arquivo (*.csv):", 10, 135, 260, 10
    CONTROL ADD TEXTBOX,  hDlg, %TXT_ARQUIVO_CSV, "", 10, 145, 260, 13
    CONTROL ADD TEXTBOX,  hDlg, %TXT_DIRETORIO_SQL_GERADO, "C:\TMP\", 10, _
        170, 260, 15
    CONTROL ADD TEXTBOX,  hDlg, %TXT_DIRETORIO_CSV_PROCESSADO, "C:\TMP\", 10, _
        200, 260, 15
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_GERAR_LOTERIA, "Gerar", 275, 200, _
        45, 15
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL7, "&Pasta de destino dos arquivos " + _
        "*.sql gerado", 10, 160, 260, 10
    CONTROL ADD CHECKBOX, hDlg, %IDC_STATUS_LOTERIA, "Verificar status da " + _
        "gera��o", 275, 160, 100, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LOG_LOTERIA, "", 10, 225, 545, 65, _
        %WS_CHILD OR %WS_VISIBLE OR %SS_LEFT OR %SS_SUNKEN, %WS_EX_LEFT OR _
        %WS_EX_LTRREADING
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_PARAR_LOTERIA, "&Parar", 450, 200, _
        50, 15, %WS_CHILD OR %WS_VISIBLE OR %WS_DISABLED OR %WS_TABSTOP OR _
        %BS_TEXT OR %BS_PUSHBUTTON OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT _
        OR %WS_EX_LTRREADING
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHK_PAUSAR_LOTERIA, "PAUSA&R", 275, 175, _
        50, 10
    CONTROL ADD CHECKBOX, hDlg, %IDC_COMMIT, "Senten�a 'commit' ap�s cada " + _
        "senten�a sql.", 275, 145, 100, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LOG, "", 5, 300, 550, 115, %WS_CHILD OR _
        %WS_VISIBLE OR %SS_LEFT OR %SS_SUNKEN, %WS_EX_LEFT OR _
        %WS_EX_LTRREADING
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL8, "&Pasta de destino dos arquivos " + _
        "*.csv processados.", 10, 190, 260, 10
#PBFORMS END DIALOG

        hJanela_Principal = hDlg

    DIALOG SHOW MODAL hDlg, CALL Janela_Principal_CBK TO lRslt


#PBFORMS BEGIN CLEANUP %IDD_DIALOG1
#PBFORMS END CLEANUP

    FUNCTION = lRslt
END FUNCTION
'------------------------------------------------------------------------------