*&---------------------------------------------------------------------*
*& Report ZOCN_EMAIL_NF - Relatório de e-mail 'Receber NF'
*&---------------------------------------------------------------------*
*& Autor Ocean: Marcus Vinícius 16/06/2024
*&---------------------------------------------------------------------*
REPORT zocn_email_nf.

INCLUDE zocn_email_nf_top. 
*&---------------------------------------------------------------------*
*&  Include           ZOCN_EMAIL_NF_TOP                         inicio *
*&---------------------------------------------------------------------*
TABLES: kna1, adr6, adrt, vbak, vbap, vbrk, vbfa, j_1bnfdoc, j_1bnflin.

TYPE-POOLS slis.

TYPES: BEGIN OF ty_saida,
  kunnr     TYPE kna1-kunnr,
*  adrnr     TYPE kna1-adrnr,
  name      TYPE kna1-name1,
  nf        TYPE j_1bnfdoc-nfenum,
  data      TYPE vbrk-fkdat,
  confer    TYPE vbrk-ernam,
  smtp_addr TYPE adr6-smtp_addr,
*  remark    TYPE adrt-remark,
END OF ty_saida.

DATA: gt_saida TYPE TABLE OF ty_saida,
      wa_saida TYPE ty_saida.

DATA: lt_fieldcat TYPE slis_t_fieldcat_alv,
      wa_fieldcat TYPE slis_fieldcat_alv.
*&---------------------------------------------------------------------*
*&  Include           ZOCN_EMAIL_NF_TOP                          final *
*&---------------------------------------------------------------------*
INCLUDE zocn_email_nf_src.
*&---------------------------------------------------------------------*
*&  Include           ZOCN_EMAIL_NF_SRC                         inicio *
*&---------------------------------------------------------------------*
SELECTION-SCREEN BEGIN OF BLOCK bll WITH FRAME TITLE TEXT-t01.

  SELECT-OPTIONS: so_fkdat FOR vbrk-fkdat OBLIGATORY,
                  so_kunnr FOR kna1-kunnr,
                  so_aufnr FOR vbap-aufnr,
                  so_nfe   FOR j_1bnfdoc-nfenum.

  PARAMETERS: p_email TYPE adr6-smtp_addr.
*              p_nfe   TYPE j_1bnfdoc-nfenum.

SELECTION-SCREEN END OF BLOCK bll.
*&---------------------------------------------------------------------*
*&  Include           ZOCN_EMAIL_NF_SRC                          final *
*&---------------------------------------------------------------------*

START-OF-SELECTION.

PERFORM f_seleciona_dados.

INCLUDE zocn_email_nf_form.
*&---------------------------------------------------------------------*
*&  Include           ZOCN_EMAIL_NF_FORM                        inicio *
*&---------------------------------------------------------------------*
* Preenchimento do Field Catalog
wa_fieldcat-fieldname = 'KUNNR'.
wa_fieldcat-seltext_m = 'Número Cliente'.
wa_fieldcat-just      = 'C'.
APPEND wa_fieldcat TO lt_fieldcat.

wa_fieldcat-fieldname = 'NAME'.
wa_fieldcat-seltext_m = 'Nome do Cliente'.
wa_fieldcat-just      = 'L'.
wa_fieldcat-outputlen = 20.
APPEND wa_fieldcat TO lt_fieldcat.

wa_fieldcat-fieldname = 'NF'.
wa_fieldcat-seltext_m = 'Nota Fiscal'.
wa_fieldcat-just      = 'C'.
wa_fieldcat-outputlen = 10.
APPEND wa_fieldcat TO lt_fieldcat.

wa_fieldcat-fieldname = 'DATA'.
wa_fieldcat-seltext_s = 'Dt. Fat.'.
wa_fieldcat-seltext_m = 'Data Faturamento'.
wa_fieldcat-just      = 'C'.
wa_fieldcat-outputlen = 10.
APPEND wa_fieldcat TO lt_fieldcat.

wa_fieldcat-fieldname = 'CONFER'.
wa_fieldcat-seltext_m = 'Criador da fatura'.
wa_fieldcat-just      = 'L'.
wa_fieldcat-outputlen = 15.
APPEND wa_fieldcat TO lt_fieldcat.

wa_fieldcat-fieldname = 'SMTP_ADDR'.
wa_fieldcat-seltext_s = 'E-mail'.
wa_fieldcat-seltext_m = 'E-mail'.
wa_fieldcat-just      = 'L'.
wa_fieldcat-outputlen = 20.
APPEND wa_fieldcat TO lt_fieldcat.

*wa_fieldcat-fieldname = 'REMARK'.
*wa_fieldcat-seltext_m = 'Observação'.
*APPEND wa_fieldcat TO lt_fieldcat.

SORT gt_saida BY kunnr nf smtp_addr ASCENDING.
DELETE ADJACENT DUPLICATES FROM gt_saida COMPARING kunnr name nf data confer smtp_addr.

IF gt_saida IS NOT INITIAL.
  CALL FUNCTION 'REUSE_ALV_GRID_DISPLAY'
  EXPORTING
    it_fieldcat      = lt_fieldcat
  TABLES
    t_outtab         = gt_saida
  EXCEPTIONS
    program_error    = 1
    OTHERS           = 2.

  IF sy-subrc <> 0.
    MESSAGE 'Erro ao exibir o ALV.' TYPE 'E'.
  ENDIF.
ELSE.
  MESSAGE 'Nenhum cliente encontrado com os critérios fornecidos.' TYPE 'I'.
ENDIF.
*&---------------------------------------------------------------------*
*&  Include           ZOCN_EMAIL_NF_FORM                         final *
*&---------------------------------------------------------------------*

*&---------------------------------------------------------------------*
*&      Form  F_SELECIONA_DADOS
*&---------------------------------------------------------------------*
FORM f_seleciona_dados.

    "Caso o e-mail esteja preenchido
    IF p_email IS NOT INITIAL.

    "Select das notas fiscais que foram faturadas nas datas delecionadas
    SELECT a~docnum, a~nfnum, a~nfenum, c~ernam, c~kunag, c~fkdat, c~vbeln INTO TABLE @DATA(lt_docs)
      FROM j_1bnfdoc AS a
      LEFT JOIN j_1bnflin AS b
        ON a~docnum  EQ b~docnum
     INNER JOIN vbrk      AS c
        ON b~refkey  EQ c~vbeln
     WHERE c~fkdat   IN @so_fkdat
       AND c~kunag   IN @so_kunnr
       AND ( a~nfnum IN @so_nfe OR a~nfenum IN @so_nfe ).

    IF sy-subrc IS NOT INITIAL.
      MESSAGE |Nenhuma Nota Fiscal foi encontrada| TYPE 'E'.
    ENDIF.

    DELETE ADJACENT DUPLICATES FROM lt_docs COMPARING docnum nfnum nfenum ernam kunag fkdat vbeln.

    "Verificando se o e-mail selecionado está cadastrado
    SELECT a~kunnr, a~adrnr, a~name1, b~smtp_addr INTO TABLE @DATA(lt_email)
      FROM kna1 AS a
     INNER JOIN adr6 AS b
        ON a~adrnr EQ b~addrnumber
      LEFT JOIN adrt AS c
        ON b~addrnumber EQ c~addrnumber
       AND b~consnumber EQ c~consnumber
       FOR ALL ENTRIES IN @lt_docs
     WHERE a~kunnr      EQ @lt_docs-kunag
       AND b~smtp_addr  EQ @p_email
       AND b~persnumber EQ @space
       AND c~remark     EQ 'Receber NF'.

    IF sy-subrc IS NOT INITIAL.
      MESSAGE |E-mail não cadastrado para Receber Nota Fiscal| TYPE 'E'.
    ENDIF.

    "Select das ordens de venda relacionadas às faturas
    SELECT vbelv, vbeln INTO TABLE @DATA(lt_vbfa)
      FROM vbfa
       FOR ALL ENTRIES IN @lt_docs
     WHERE vbeln   EQ @lt_docs-vbeln
       AND vbtyp_n EQ 'M'
       AND vbtyp_v EQ 'C'.

    CHECK lt_vbfa IS NOT INITIAL.

    DELETE ADJACENT DUPLICATES FROM lt_vbfa COMPARING vbelv vbeln.

    SELECT a~vbeln, b~aufnr INTO TABLE @DATA(lt_aufnr)
      FROM vbak AS a
 LEFT JOIN vbap AS b
        ON a~vbeln EQ b~vbeln
   FOR ALL ENTRIES IN @lt_vbfa
     WHERE a~vbeln EQ @lt_vbfa-vbelv
       AND b~aufnr IN @so_aufnr.

    IF sy-subrc IS INITIAL.

    DELETE ADJACENT DUPLICATES FROM lt_aufnr COMPARING vbeln aufnr.

      "loop para preenchimento da tabela de saída
      LOOP AT lt_docs ASSIGNING FIELD-SYMBOL(<fs_docs>).

        READ TABLE lt_vbfa ASSIGNING FIELD-SYMBOL(<fs_vbfa>) WITH KEY vbeln = <fs_docs>-vbeln.

        IF sy-subrc IS NOT INITIAL.
          CONTINUE.
        ENDIF.

        READ TABLE lt_aufnr ASSIGNING FIELD-SYMBOL(<fs_aufnr>) WITH KEY vbeln = <fs_vbfa>-vbelv.

        IF sy-subrc IS NOT INITIAL.
          CONTINUE.
        ENDIF.

        READ TABLE lt_email ASSIGNING FIELD-SYMBOL(<fs_email>) WITH KEY kunnr = <fs_docs>-kunag.

        IF sy-subrc IS NOT INITIAL.
          CONTINUE.
        ELSE.

          wa_saida-kunnr      = <fs_docs>-kunag.  "Número do cliente
          PACK wa_saida-kunnr TO wa_saida-kunnr.
          wa_saida-name       = <fs_email>-name1.   "Nome do cliente
          wa_saida-data       = <fs_docs>-fkdat.  "Data de faturamento
          wa_saida-confer     = <fs_docs>-ernam.  "Usuário que fez o faturamento
          wa_saida-smtp_addr  = p_email.          "E-mail cadastrado para receber as NFs

          IF <fs_docs>-nfnum IS NOT INITIAL.
            wa_saida-nf    = <fs_docs>-nfnum.     "Número da Nota Fiscal de serviço
          ELSE.
            wa_saida-nf    = <fs_docs>-nfenum.    "Número da Nota Fiscal
          ENDIF.

          APPEND wa_saida TO gt_saida.

        ENDIF.

      ENDLOOP.

    ELSE.
      MESSAGE |Nenhum registro encontrado.| TYPE 'E'.
    ENDIF.

  "Caso o e-mail não esteja preenchido
  ELSE.

    "Select das notas fiscais que foram faturadas nas datas delecionadas
    SELECT a~docnum, a~nfnum, a~nfenum, c~ernam, c~kunag, c~fkdat, c~vbeln INTO TABLE @lt_docs
      FROM j_1bnfdoc AS a
      LEFT JOIN j_1bnflin AS b
        ON a~docnum  EQ b~docnum
     INNER JOIN vbrk      AS c
        ON b~refkey  EQ c~vbeln
     WHERE c~fkdat   IN @so_fkdat
       AND c~kunag   IN @so_kunnr
       AND ( a~nfnum IN @so_nfe OR a~nfenum IN @so_nfe ).

    IF sy-subrc IS NOT INITIAL.
      MESSAGE |Nenhuma Nota Fiscal foi encontrada| TYPE 'E'.
    ENDIF.

    DELETE ADJACENT DUPLICATES FROM lt_docs COMPARING docnum nfnum nfenum ernam kunag fkdat vbeln.

    "Select das Ordens de Venda relacionadas às faturas
    SELECT vbelv vbeln INTO TABLE lt_vbfa
      FROM vbfa
       FOR ALL ENTRIES IN lt_docs
     WHERE vbeln   EQ lt_docs-vbeln
       AND vbtyp_n EQ 'M'
       AND vbtyp_v EQ 'C'.

    CHECK lt_vbfa IS NOT INITIAL.

    DELETE ADJACENT DUPLICATES FROM lt_vbfa COMPARING vbelv vbeln.

    SELECT a~vbeln, b~aufnr, c~name1, c~adrnr, c~kunnr INTO TABLE @DATA(lt_kunnr)
      FROM vbak AS a
      LEFT JOIN vbap AS b
        ON a~vbeln EQ b~vbeln
     INNER JOIN kna1 AS c
        ON a~kunnr EQ c~kunnr
       FOR ALL ENTRIES IN @lt_vbfa
     WHERE a~vbeln EQ @lt_vbfa-vbelv
       AND b~aufnr IN @so_aufnr.

    IF sy-subrc IS NOT INITIAL.
      MESSAGE |Nenhum registro encontrado.| TYPE 'E'.
    ENDIF.

    DELETE ADJACENT DUPLICATES FROM lt_kunnr COMPARING vbeln aufnr name1 adrnr.

    "Select dos e-mails cadastrados pra receber as NFs
    SELECT a~smtp_addr, a~addrnumber INTO TABLE @DATA(lt_emails)
      FROM adr6 AS a
     INNER JOIN adrt AS b
        ON a~addrnumber EQ b~addrnumber
       AND a~consnumber EQ b~consnumber
       FOR ALL ENTRIES IN @lt_kunnr
     WHERE a~addrnumber EQ @lt_kunnr-adrnr
       AND a~persnumber EQ @space
       AND b~remark     EQ 'Receber NF'.

    IF sy-subrc IS INITIAL.

      "Loop para preencher a tabela de saída
      LOOP AT lt_docs ASSIGNING <fs_docs>.

        READ TABLE lt_vbfa ASSIGNING <fs_vbfa> WITH KEY vbeln = <fs_docs>-vbeln.

        IF sy-subrc IS NOT INITIAL.
          CONTINUE.
        ENDIF.

        READ TABLE lt_kunnr ASSIGNING FIELD-SYMBOL(<fs_kunnr>) WITH KEY vbeln = <fs_vbfa>-vbelv.

        IF sy-subrc IS NOT INITIAL.
          CONTINUE.
        ENDIF.

        LOOP AT lt_emails ASSIGNING FIELD-SYMBOL(<fs_emails>) WHERE addrnumber = <fs_kunnr>-adrnr.

          wa_saida-kunnr      = <fs_docs>-kunag.
          PACK wa_saida-kunnr TO wa_saida-kunnr.
          wa_saida-name       = <fs_kunnr>-name1.
          wa_saida-data       = <fs_docs>-fkdat.
          wa_saida-confer     = <fs_docs>-ernam.
          wa_saida-smtp_addr  = <fs_emails>-smtp_addr.

          IF <fs_docs>-nfnum IS NOT INITIAL.
            wa_saida-nf    = <fs_docs>-nfnum.
          ELSE.
            wa_saida-nf    = <fs_docs>-nfenum.
          ENDIF.

          APPEND wa_saida TO gt_saida.

        ENDLOOP.

      ENDLOOP.

    ENDIF.

  ENDIF.

ENDFORM.