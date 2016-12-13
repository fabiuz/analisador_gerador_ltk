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
%IDC_BTN_GERAR_INSERT         = 1020
%IDC_BTN_PARA_INSERT          = 1022
%IDC_BTN_PARA_TODAS           = 1021
%IDC_BUTTON1                  = 1019    '*
%IDC_CHECKBOX1                = 1017
%IDC_CHK_ATUALIZAR_TODAS      = 1010
%IDC_CMB_APOSTAR_COM_TODAS    = 1005
%IDC_CMB_JOGO_TIPO_TODAS      = 1002
%IDC_FRAME1                   = 1001
%IDC_FRAME2                   = 1012
%IDC_GERAR_TODAS              = 1009
%IDC_LABEL1                   = 1003
%IDC_LABEL2                   = 1004
%IDC_LABEL3                   = 1006
%IDC_LABEL4                   = 1013
%IDC_LABEL5                   = 1015
%IDC_LABEL6                   = 1018
%IDC_LBL_LOG_TODAS            = 1011
%IDC_TEXTBOX1                 = 1007    '*
%IDC_TXT_ARQUIVO_ORIGEM       = 1014
%IDC_TXT_PASTA_DESTINO_INSERT = 1016
%IDC_TXT_PASTA_DESTINO_TODAS  = 1008
%IDD_DIALOG1                  =  101
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
' ** DECLARAÇÕES GLOBAIS **
'------------------------------------------------------------------------------
GLOBAL hJanela_Principal AS DWORD



FUNCTION Carregar_Lista(BYVAL strJogo AS STRING) AS LONG
    COMBOBOX RESET hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS

    strJogo = UCASE$(strJogo)
    SELECT CASE strJogo
        CASE "QUINA"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "5 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "6 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "7 números"

        CASE "MEGASENA
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "6 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "7 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "8 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "9 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "10 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "11 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "12 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "14 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "15 números"

        CASE "LOTOFACIL"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "15 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "16 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "17 números"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "18 números"

        CASE "LOTOMANIA"
            COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "50 números"
    END SELECT

    COMBOBOX SELECT hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, 1

END FUNCTION

'
' Carrega as caixas de combinação 'TIPO DE JOGO' e 'APOSTAR COM'
'
SUB Carregar_Lista_Inicial()
    COMBOBOX RESET hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS
    COMBOBOX RESET hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS
    '
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "QUINA"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "MEGASENA"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "LOTOFACIL"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, "LOTOMANIA"

    '
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "5 números"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "6 números"
    COMBOBOX ADD hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, "7 números"

    '
    COMBOBOX SELECT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, 1
    COMBOBOX SELECT hJanela_Principal, %IDC_CMB_APOSTAR_COM_TODAS, 1
END SUB

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

    SELECT CASE AS LONG CB.MSG
        CASE %WM_INITDIALOG
            ' Initialization handler
            ' Preencher as caixas de combinação 'JOGO_TIPO' e 'APOSTAR COM'
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

        CASE %WM_COMMAND
            ' Process control notifications
            SELECT CASE AS LONG CB.CTL
                ' Se o usuário clicar no tipo do jogo, alterar, a caixa de combinação "QT DE APOSTAS"
                CASE %IDC_CMB_JOGO_TIPO_TODAS
                    IF CB.CTLMSG = %CBN_SELCHANGE THEN
                        DIM strTexto AS STRING
                        DIM uIndice AS LONG

                        COMBOBOX GET SELECT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS TO uIndice
                        COMBOBOX GET TEXT hJanela_Principal, %IDC_CMB_JOGO_TIPO_TODAS, uIndice TO strTexto
                        Carregar_Lista(strTexto)
                    END IF

            END SELECT
    END SELECT
END FUNCTION
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
'   ** Dialogs **
'------------------------------------------------------------------------------
FUNCTION Exibir_Janela_Principal(BYVAL hParent AS DWORD) AS LONG
    LOCAL lRslt AS LONG

#PBFORMS BEGIN DIALOG %IDD_DIALOG1->->
    LOCAL hDlg  AS DWORD

    DIALOG NEW hParent, "Analisador e Gerador - LTK", 213, 182, 458, 283, _
        %WS_POPUP OR %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR _
        %WS_MINIMIZEBOX OR %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME _
        OR %DS_3DLOOK OR %DS_NOFAILCREATE OR %DS_SETFONT, _
        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR, TO hDlg
    CONTROL ADD FRAME,    hDlg, %IDC_FRAME1, "Gerar todas as combinações", 5, _
        5, 215, 95
    CONTROL ADD COMBOBOX, hDlg, %IDC_CMB_JOGO_TIPO_TODAS, , 10, 25, 95, 40, _
        %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR _
        %CBS_DROPDOWNLIST, %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR
    CONTROL ADD COMBOBOX, hDlg, %IDC_CMB_APOSTAR_COM_TODAS, , 115, 25, 100, _
        40, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR _
        %CBS_DROPDOWNLIST, %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR
    CONTROL ADD TEXTBOX,  hDlg, %IDC_TXT_PASTA_DESTINO_TODAS, "C:\TMP\", 10, _
        50, 205, 15
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHK_ATUALIZAR_TODAS, "Verificar status " + _
        "da geração", 10, 85, 100, 10
    CONTROL ADD BUTTON,   hDlg, %IDC_GERAR_TODAS, "Gerar", 165, 80, 50, 15
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL1, "Jogo:", 10, 15, 75, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL2, "&Apostas com:", 115, 15, 75, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL3, "Pasta de destino (nome do " + _
        "arquivo gerado automaticamente):", 10, 40, 200, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LBL_LOG_TODAS, "", 10, 65, 205, 10
    CONTROL ADD FRAME,    hDlg, %IDC_FRAME2, "Gerar Insert e Where para todas " + _
        "as combinações", 225, 5, 230, 95
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL4, "Nome do arquivo:", 230, 15, _
        100, 10
    CONTROL ADD TEXTBOX,  hDlg, %IDC_TXT_ARQUIVO_ORIGEM, "", 230, 25, 220, 13
    CONTROL ADD TEXTBOX,  hDlg, %IDC_TXT_PASTA_DESTINO_INSERT, "C:\TMP\", _
        230, 50, 220, 15
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL5, "&Pasta de destino", 230, 40, _
        100, 10
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHECKBOX1, "Verificar status da " + _
        "geração", 230, 85, 100, 10
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_GERAR_INSERT, "Gerar", 400, 80, 50, _
        15
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL6, "", 230, 65, 220, 10
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_PARA_TODAS, "&Parar", 110, 80, 50, _
        15, %WS_CHILD OR %WS_VISIBLE OR %WS_DISABLED OR %WS_TABSTOP OR _
        %BS_TEXT OR %BS_PUSHBUTTON OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT _
        OR %WS_EX_LTRREADING
    CONTROL ADD BUTTON,   hDlg, %IDC_BTN_PARA_INSERT, "&Parar", 345, 80, 50, _
        15, %WS_CHILD OR %WS_VISIBLE OR %WS_DISABLED OR %WS_TABSTOP OR _
        %BS_TEXT OR %BS_PUSHBUTTON OR %BS_CENTER OR %BS_VCENTER, %WS_EX_LEFT _
        OR %WS_EX_LTRREADING
#PBFORMS END DIALOG

        hJanela_Principal = hDlg

    DIALOG SHOW MODAL hDlg, CALL Janela_Principal_CBK TO lRslt


#PBFORMS BEGIN CLEANUP %IDD_DIALOG1
#PBFORMS END CLEANUP

    FUNCTION = lRslt
END FUNCTION
'------------------------------------------------------------------------------
