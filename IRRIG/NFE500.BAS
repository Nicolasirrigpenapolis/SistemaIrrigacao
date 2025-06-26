Attribute VB_Name = "modNFE500"

'Rotina de transmisso de NF-e no layout 4.00 (compatvel com verso 5.00)
Public Sub TransmitirNFe500(NFe As String, chaveNFe As String)
   Dim msgDados As String, msgRetWS As String, msgResultado As String, siglaWS As String, certificado As String
   Dim licenca As String, NFeAssinada As String, nroRecibo As String, cStat As Long, versao As String, dhRecibo As String
   Dim tMed As String, i As Integer, objNFeUtil As Object, Arq As Long, proxy As String, usuario As String, senha As String
   Dim tpAmbiente As Long, numeroLote As Long, ide As String
   'Parmetros do identificador202006
   Dim ide_cUF As Long, ide_cNF As String, ide_natOp As String, ide_indPag As Long, ide_mode As Long
   Dim ide_serie As Long, ide_nNF As Long, ide_dhEmi As String, ide_dhSaiEnt As String
   Dim ide_tpNF As Long, ide_idDest As Long, ide_cMunFG As String, ide_NFRefs As String
   Dim ide_tpImp As Long, ide_tpEmis As Long, ide_cDV As Long, ide_tpAmb As Long
   Dim ide_finNFe As Long, ide_indFinal As Long, ide_indPres As Long
   Dim ide_procEmi As Long, ide_verProc As String, ide_dhCont As String, ide_xJust As String
   Dim indIntermed As Long, indSinc As Long

   On Error GoTo DeuErro

   Screen.MousePointer = vbHourglass
   lblProgresso.Caption = "Transmitindo NFe..."
   Set objNFeUtil = CreateObject("NFe_util_2G.util")

   certificado = Parametros_da_NFe![Certificado Digital]
   siglaWS = "SP"
   versao = "4.00"
   licenca = "23433296a6c7b3e6e602aa8c71ded18804ec302656cfa8590ff7b35454b14834f4f5aed7ed5b4e4986a6a2d2a4b1abad6758927b36d70c5a1d4588597eaf0de2"

   indSinc = 1
   ide_indPag = IIf(Nota_Fiscal![Forma de Pagamento] = "Vista", 0, 1)
   indIntermed = 0
   numeroLote = 1
   If Parametros_da_NFe!ambiente = 0 Then
      tpAmbiente = 1
   Else
      tpAmbiente = 2
   End If

   ide = objNFeUtil.identificador202006(ide_cUF, ide_cNF, ide_natOp, ide_indPag, ide_mode, ide_serie, ide_nNF, ide_dhEmi, ide_dhSaiEnt, ide_tpNF, ide_idDest, ide_cMunFG, ide_NFRefs, ide_tpImp, ide_tpEmis, ide_cDV, ide_tpAmb, ide_finNFe, ide_indFinal, ide_indPres, indIntermed, ide_procEmi, ide_verProc, ide_dhCont, ide_xJust, indSinc)

   For i = 1 To 5
      NFeAssinada = objNFeUtil.nfeAutorizacao400(siglaWS, numeroLote, tpAmbiente, versao, indSinc, NFe, certificado, msgDados, msgRetWS, cStat, msgResultado, nroRecibo, dhRecibo, tMed, proxy, usuario, senha, licenca)
      With vgTb
         .Edit
         If Vazio(!Observao) Then
            !Observao = "Transmitindo NFe... Tentativa " & i
         Else
            !Observao = !Observao & vbCrLf & "Transmitindo NFe... Tentativa " & i
         End If
         .Update
         .BookMark = .LastModified
      End With
      If cStat <> 105 Then
         GoTo Continua500
      End If
   Next
   MsgBox "NFe No Enviada! Tente Mais Tarde.", vbExclamation + vbOKOnly, vaTitulo

   With vgTb
      .Edit
      If Vazio(!Observao) Then
         !Observao = "Erro ao Transmitir NFe"
      Else
         !Observao = !Observao & vbCrLf & "Erro ao Transmitir NFe"
      End If
      .Update
      .BookMark = .LastModified
   End With
   GoTo SaidaSub500

Continua500:
   Select Case cStat
      Case 5000 To 7003
         MsgBox "Descrio do erro: " & Chr(13) & Chr(13) + msgResultado, vbCritical + vbOKOnly, "Ateno: O envio da NFe falhou..."
         With vgTb
            .Edit
            If Vazio(!Observao) Then
               !Observao = "Erro ao Transmitir NFe: " & msgResultado
            Else
               !Observao = !Observao & vbCrLf & "Erro ao Transmitir NFe: " & msgResultado
            End If
            .Update
            .BookMark = .LastModified
         End With
         GoTo SaidaSub500
      Case 103
         With Nota_Fiscal
            .Edit
            !Transmitido = True
            ![Nmero do Recibo da NFe] = nroRecibo
            ![Chave de Acesso da NFe] = chaveNFe
            .Update
            .BookMark = .LastModified
         End With
         MsgBox "Transmitido com Sucesso!!!", vbInformation + vbOKOnly, "Ateno: NFe enviada!"
         Arq = FreeFile
         If Parametros_da_NFe!ambiente = 0 Then
            If Dir(Parametros_da_NFe![Diretrio 2 NFe Produo] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_GERADO.xml", vbArchive) <> "" Then Kill Parametros_da_NFe![Diretrio 2 NFe Produo] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_GERADO.xml"
            If Dir(Parametros_da_NFe![Diretrio 2 NFe Produo] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_ASSINADO.xml", vbArchive) <> "" Then Kill Parametros_da_NFe![Diretrio 2 NFe Produo] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_ASSINADO.xml"
            Open Parametros_da_NFe![Diretrio 2 NFe Produo] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_ASSINADO.xml" For Append As #Arq
               Print #Arq, NFeAssinada
            Close #Arq
         Else
            If Dir(Parametros_da_NFe![Diretrio 2 NFe Homologao] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_GERADO.xml", vbArchive) <> "" Then Kill Parametros_da_NFe![Diretrio 2 NFe Homologao] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_GERADO.xml"
            If Dir(Parametros_da_NFe![Diretrio 2 NFe Homologao] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_ASSINADO.xml", vbArchive) <> "" Then Kill Parametros_da_NFe![Diretrio 2 NFe Homologao] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_ASSINADO.xml"
            Open Parametros_da_NFe![Diretrio 2 NFe Homologao] & Format(Nota_Fiscal![Nmero da NFe], "000000") & "_ASSINADO.xml" For Append As #Arq
               Print #Arq, NFeAssinada
            Close #Arq
         End If
         With vgTb
            .Edit
            If Vazio(!Observao) Then
               !Observao = "Transmitido com Sucesso"
            Else
               !Observao = !Observao & vbCrLf & "Transmitido com Sucesso"
            End If
            .Update
            .BookMark = .LastModified
         End With
      Case 108, 109
         MsgBox msgResultado & Chr(13) & Chr(13) + "O Web Service com problemas de recepo", vbCritical + vbOKOnly, "Ateno: WS com problemas na recepo!"
         With vgTb
            .Edit
            If Vazio(!Observao) Then
               !Observao = "Erro ao Transmitir NFe: " & msgResultado
            Else
               !Observao = !Observao & vbCrLf & "Erro ao Transmitir NFe: " & msgResultado
            End If
            .Update
            .BookMark = .LastModified
         End With
         GoTo SaidaSub500
      Case Else
         MsgBox "Descrio do erro: " & Chr(13) & Chr(13) & cStat & " - " & msgResultado, vbCritical + vbOKOnly, "Ateno: O WS rejeitou a NFe..."
         With vgTb
            .Edit
            If Vazio(!Observao) Then
               !Observao = "Erro ao Transmitir NFe: " & msgResultado
            Else
               !Observao = !Observao & vbCrLf & "Erro ao Transmitir NFe: " & msgResultado
            End If
            .Update
            .BookMark = .LastModified
         End With
         GoTo SaidaSub500
   End Select

   RetornoNFe310 NFeAssinada, nroRecibo

DeuErro:
   If Err <> 0 Then MsgBox Err.Description, vbCritical + vbOKOnly, vaTitulo

SaidaSub500:
   Screen.MousePointer = vbDefault
   Set objNFeUtil = Nothing
   If Parametros_da_NFe!ambiente = 0 Then
      lblProgresso.Caption = ""
   Else
      lblProgresso.Caption = "AMBIENTE NFe EM MODO TESTE"
   End If
End Sub
