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
#OPTIMIZE CODE OFF

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
%IDC_CHK_ATUALIZAR_TODAS     = 1010
%IDC_CMB_APOSTAR_COM_TODAS   = 1005
%IDC_CMB_JOGO_TIPO_TODAS     = 1002
%IDC_FRAME1                  = 1001
%IDC_GERAR_TODAS             = 1009
%IDC_LABEL1                  = 1003
%IDC_LABEL2                  = 1004
%IDC_LABEL3                  = 1006
%IDC_LBL_LOG_TODAS           = 1011
%IDC_TEXTBOX1                = 1007 '*
%IDC_TXT_PASTA_DESTINO_TODAS = 1008
%IDD_DIALOG1                 =  101
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

    DIALOG NEW hParent, "Analisar e Gerador - LTK", 117, 118, 695, 373, _
        %WS_POPUP OR %WS_BORDER OR %WS_DLGFRAME OR %WS_SYSMENU OR _
        %WS_MINIMIZEBOX OR %WS_CLIPSIBLINGS OR %WS_VISIBLE OR %DS_MODALFRAME _
        OR %DS_3DLOOK OR %DS_NOFAILCREATE OR %DS_SETFONT, _
        %WS_EX_CONTROLPARENT OR %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR, TO hDlg
    CONTROL ADD FRAME,    hDlg, %IDC_FRAME1, "Gerar todas as combinações", 5, _
        5, 270, 95
    CONTROL ADD COMBOBOX, hDlg, %IDC_CMB_JOGO_TIPO_TODAS, , 10, 25, 95, 40, _
        %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR _
        %CBS_DROPDOWNLIST, %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR
    CONTROL ADD COMBOBOX, hDlg, %IDC_CMB_APOSTAR_COM_TODAS, , 115, 25, 100, _
        40, %WS_CHILD OR %WS_VISIBLE OR %WS_TABSTOP OR %WS_VSCROLL OR _
        %CBS_DROPDOWNLIST, %WS_EX_LEFT OR %WS_EX_LTRREADING OR _
        %WS_EX_RIGHTSCROLLBAR
    CONTROL ADD TEXTBOX,  hDlg, %IDC_TXT_PASTA_DESTINO_TODAS, "", 10, 50, _
        205, 15
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL1, "Jogo:", 10, 15, 75, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL2, "&Apostas com:", 115, 15, 75, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LABEL3, "Pasta de destino (nome do " + _
        "arquivo gerado automaticamente):", 10, 40, 200, 10
    CONTROL ADD BUTTON,   hDlg, %IDC_GERAR_TODAS, "Gerar", 220, 50, 50, 15
    CONTROL ADD CHECKBOX, hDlg, %IDC_CHK_ATUALIZAR_TODAS, "Verificar status " + _
        "da geração", 10, 85, 100, 10
    CONTROL ADD LABEL,    hDlg, %IDC_LBL_LOG_TODAS, "", 10, 70, 260, 10
#PBFORMS END DIALOG

    DIALOG SHOW MODAL hDlg, CALL Janela_Principal_CBK TO lRslt

#PBFORMS BEGIN CLEANUP %IDD_DIALOG1
#PBFORMS END CLEANUP

    FUNCTION = lRslt
END FUNCTION
'------------------------------------------------------------------------------
