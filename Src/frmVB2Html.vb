
' Fichier frmVB2Html : Créer un rapport Html d'un projet Visual Basic 6, et VB .Net
' ---------------------------------------------------------------------------------

Imports VB = Microsoft.VisualBasic ' Pour VB6.Left$

Public Class frmVB2Html

    Public m_sCheminPrj$

#Region "Déclarations"

    ' Sauver l'ordre des fichiers du projet dans un fichier ini dans le dossier de chaque projet
    '  de façon à pouvoir rapidement mettre à jour la doc html d'un projet
    Private Const bSauverOrdreFichierIni As Boolean = True

    Private m_patienter As New clsPatienter

    Private KWList_VBNET, KWList_VB6, ActKWList As String
    Private m_sContenuModeleHtml$
    Private ProjectVersion, ProjectTitle As String

    Private Const sMenuCtx_CleCmd$ = "ConvertirEnHtml"
    Private Const sMenuCtx_DescriptionCmd$ = "Convertir en Html"
    Private Const sMenuCtx_TypeFichierVB6$ = "VisualBasic.Project"
    ' Si Visual Studio 2003 est installé, alors il faut ajouter la clé 2003
    ' Si Visual Studio 2005 est installé, alors il faut ajouter la clé 2005
    ' Si Visual Studio 2008 est installé, alors il faut ajouter la clé 2008
    ' Si Visual Studio 2010 est installé, alors il faut ajouter la clé 2010
    ' Conclusion : pour que le logiciel fonctionne partout, il faut ajouter toutes les clés
    Private Const sMenuCtx_TypeFichierVB2003$ = "VBExpress.vbproj.7.0" ' 17/07/2011
    Private Const sMenuCtx_TypeFichierVB2005$ = "VBExpress.vbproj.8.0"
    Private Const sMenuCtx_TypeFichierVB2008$ = "VBExpress.vbproj.9.0"
    Private Const sMenuCtx_TypeFichierVB2010$ = "VBExpress.Launcher.vbproj.10.0" ' 17/07/2011

    ' 15/08/2024 Aucun ne fonctionne, il faut donc cocher tous les fichiers :
    'Private Const sMenuCtx_TypeFichierVB2022$ = "VisualStudio.Launcher.vbproj.12.0"
    'Private Const sMenuCtx_TypeFichierVB2022$ = "VisualStudio.Launcher.vbproj.15.0"
    'Private Const sMenuCtx_TypeFichierVB2022$ = "VisualStudio.vbproj.10.0"
    'Private Const sMenuCtx_TypeFichierVB2022$ = "VisualStudio.vbproj.12.0" 

    Private Const sMenuCtx_TypeFichierTous$ = "*" ' 22/06/2014

    Private Const sExtHtml$ = ".html"
    Private Const sFiltreHtml$ = "Fichier HTML (*.html)|*.html"

#End Region

#Region "Initialisation"

    Private Sub frmMain_Load(eventSender As Object, eventArgs As EventArgs) Handles MyBase.Load

        Me.Text = sNomAppli & " v" & sVersionAppli & " (" & sDateVersionAppli & ")"
        If bDebug Then Me.Text &= " - Debug"

        LoadKeyWords()
        ChargerModeleHtml()
        LoadConfig()

        'If bDebug Then Me.m_sCheminPrj = Application.StartupPath & "VB2Html.vbproj"

        If Not String.IsNullOrEmpty(Me.m_sCheminPrj) Then
            Me.txtCheminProjetAConv.Text = Me.m_sCheminPrj
            VerifierFichierProjet(bInitHtmlDest:=True)
        Else
            VerifierFichierProjet()
        End If
        VerifierMenuCtx()

    End Sub

    Private Sub frmMain_FormClosed(eventSender As Object,
        eventArgs As Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        SaveConfig()

    End Sub

    Private Sub SaveConfig()

        ' Le fichier sera sauvé ici :
        '\Documents and Settings\<utilisateur>\Local Settings\Application Data\
        ' VB2Html.exe_Url_xxx...xxx\2.0.2.xxxxx\user.config

        My.Settings.CheminFichierProjet = Me.txtCheminProjetAConv.Text
        My.Settings.CheminFichierModeleHtml = Me.txtCheminModeleHtml.Text
        My.Settings.CheminFichierExportHtml = Me.txtCheminHtmlDest.Text

        My.Settings.Rendu_GenererTDM = Me.chkTDM.CheckState = CheckState.Checked
        My.Settings.Rendu_GenererMultipageHtml =
            Me.chkMulti.CheckState = CheckState.Checked
        My.Settings.Rendu_TronquerLignes = Me.chkTronc.CheckState = CheckState.Checked
        If Me.txtTroncLen.Text.Length > 0 Then
            My.Settings.Rendu_TronquerLignes_NbLignes = iConv(Me.txtTroncLen.Text, 80)
        End If
        My.Settings.Rendu_DeclarationProcComplete =
            Me.chkCompDec.CheckState = CheckState.Checked
        My.Settings.Rendu_FrameHtml = Me.chkFrames.CheckState = CheckState.Checked

        My.Settings.IgnorerFrmDesigner =
            Me.chkIgnorerModDesign.CheckState = CheckState.Checked
        My.Settings.SelectionnerTousLesFichiers =
            Me.chkSelectTousLesFichiers.CheckState = CheckState.Checked

        My.Settings.VB2Txt = Me.chkVB2Txt.CheckState = CheckState.Checked

        SauverOdreFichiers()

    End Sub

    Private Function sCheminIni$()

        If Me.txtCheminProjetAConv.Text.Length = 0 Then sCheminIni = "" : Exit Function
        sCheminIni = sDossierParent(Me.txtCheminProjetAConv.Text) & "\" &
            sNomDossierFinal(Me.txtCheminProjetAConv.Text) & ".ini"

    End Function

    Private Sub SauverOdreFichiers()

        ' Sauver l'ordre des fichiers et s'ils sont sélectionnés ou pas

        If Not bSauverOrdreFichierIni Then Exit Sub

        ' 17/08/2009 Vérifier si les fichiers indiqués dans le fichier .ini
        '  existent toujours
        ' 13/09/2009 Le chemin peut être vide !
        Dim sCheminPrj$ = Me.txtCheminProjetAConv.Text
        If sCheminPrj.Length = 0 Then Exit Sub
        Dim sCheminDossierPrj$ = IO.Path.GetDirectoryName(sCheminPrj)

        Dim sb As New System.Text.StringBuilder
        Dim i%
        For i = 0 To Me.lbFichiers.Items.Count - 1
            Dim bSel As Boolean = Me.lbFichiers.GetSelected(i)

            Dim sFichier$ = Me.lbFichiers.Items(i).ToString
            Dim sChemin$ = sCheminDossierPrj & "\" & sFichier
            If Not bFichierExiste(sChemin) Then Continue For

            sb.Append(Me.lbFichiers.Items(i).ToString & ":" & bSel & vbCrLf)
        Next
        Dim sCheminIni0$ = sCheminIni()
        If sCheminIni0.Length = 0 Then Exit Sub
        If Not bEcrireFichier(sCheminIni0, sb) Then Exit Sub

    End Sub

    Private Function bLireOrdreFichiers() As Boolean

        ' Rétablir l'ordre des fichiers et s'ils sont sélectionnés ou pas

        If Not bSauverOrdreFichierIni Then Return False

        Dim sCheminIni0$ = sCheminIni()
        If sCheminIni0.Length = 0 Then Return False
        If Not bFichierExiste(sCheminIni0) Then Return False
        Dim asIni$() = IO.File.ReadAllLines(sCheminIni)
        Me.lbFichiers.Items.Clear()
        For Each sLigne As String In asIni
            Dim asChps$() = sLigne.Split(CChar(":"))
            Dim sFichier$ = ""
            Dim s_bSel$ = ""
            If asChps.GetUpperBound(0) >= 0 Then sFichier = asChps(0)
            If asChps.GetUpperBound(0) >= 1 Then s_bSel = asChps(1)
            Me.lbFichiers.Items.Add(sFichier)
            Dim i% = Me.lbFichiers.Items.Count - 1
            If s_bSel = "True" Then
                Me.lbFichiers.SetSelected(i, True)
            Else
                Me.lbFichiers.SetSelected(i, False)
            End If
        Next
        bLireOrdreFichiers = True

    End Function

    Private Sub LoadConfig()

        Me.txtCheminProjetAConv.Text = My.Settings.CheminFichierProjet
        Me.txtCheminHtmlDest.Text = My.Settings.CheminFichierExportHtml

        If My.Settings.Rendu_GenererTDM Then Me.chkTDM.CheckState = CheckState.Checked
        If My.Settings.Rendu_GenererMultipageHtml Then _
            Me.chkMulti.CheckState = CheckState.Checked
        If My.Settings.Rendu_TronquerLignes Then _
            Me.chkTronc.CheckState = CheckState.Checked
        Me.txtTroncLen.Text = My.Settings.Rendu_TronquerLignes_NbLignes.ToString
        If My.Settings.Rendu_DeclarationProcComplete Then _
            Me.chkCompDec.CheckState = CheckState.Checked
        If My.Settings.Rendu_FrameHtml Then Me.chkFrames.CheckState = CheckState.Checked

        If My.Settings.IgnorerFrmDesigner Then _
            Me.chkIgnorerModDesign.CheckState = CheckState.Checked
        If My.Settings.SelectionnerTousLesFichiers Then _
            Me.chkSelectTousLesFichiers.CheckState = CheckState.Checked

        If My.Settings.VB2Txt Then Me.chkVB2Txt.CheckState = CheckState.Checked

    End Sub

    Private Sub LoadKeyWords()

        KWList_VB6 = ""
        KWList_VBNET = ""

        Dim sChemin$ = Application.StartupPath & "\MotsCles\VB6.txt"
        If Not bFichierExiste(sChemin) Then
            Dim sLignes$ = My.Resources.VB6
            Dim asLignes$() = sLignes.Split(CChar(vbCrLf))
            Dim sLigne$
            Dim sbMotsCles As New System.Text.StringBuilder
            For Each sLigne In asLignes
                sbMotsCles.Append("|" & sLigne.Trim.ToLower)
            Next
            sbMotsCles.Append("|")
            KWList_VB6 = sbMotsCles.ToString
        Else
            Dim KWCanal% = FreeFile()
            FileOpen(KWCanal, sChemin, OpenMode.Input)
            Do Until EOF(KWCanal)
                Dim TempLine$ = LineInput(KWCanal)
                KWList_VB6 &= "|" & LCase(Trim(TempLine))
            Loop
            FileClose(KWCanal)
            KWList_VB6 &= "|"
        End If

        sChemin = Application.StartupPath & "\MotsCles\VB_NET.txt"
        If Not bFichierExiste(sChemin) Then
            Dim sLignes$ = My.Resources.VB_NET
            Dim asLignes$() = sLignes.Split(CChar(vbCrLf))
            Dim sLigne$
            Dim sbMotsCles As New System.Text.StringBuilder
            For Each sLigne In asLignes
                sbMotsCles.Append("|" & sLigne.Trim.ToLower)
            Next
            sbMotsCles.Append("|")
            KWList_VBNET = sbMotsCles.ToString
        Else
            Dim KWCanal% = FreeFile()
            FileOpen(KWCanal, sChemin, OpenMode.Input)
            Do Until EOF(KWCanal)
                Dim TempLine$ = LineInput(KWCanal)
                KWList_VBNET &= "|" & LCase(Trim(TempLine))
            Loop
            FileClose(KWCanal)
            KWList_VBNET &= "|"
        End If

    End Sub

    Private Sub ChargerModeleHtml()

        Dim bLireModeleHtml As Boolean = True
        Dim sCheminModele$ = My.Settings.CheminFichierModeleHtml
        If Not bFichierExiste(sCheminModele) Then
            ' Si le fichier modèle n'existe pas alors vérifier le chemin par défaut
            sCheminModele = Application.StartupPath & "\Modeles\Default.html"
            If Not bFichierExiste(sCheminModele) Then
                ' Si le fichier par défaut n'existe pas alors 
                '  prendre le modèle enregistré dans les ressources
                bLireModeleHtml = False
                Me.txtCheminModeleHtml.Text = "(modèle par défaut)"
                Me.m_sContenuModeleHtml = My.Resources._Default
            End If
        End If
        If bLireModeleHtml Then
            Me.m_sContenuModeleHtml = sLireFichier(sCheminModele)
            Me.txtCheminModeleHtml.Text = sCheminModele
        End If
        My.Settings.CheminFichierModeleHtml = Me.txtCheminModeleHtml.Text

    End Sub

    Private Sub chkModDesign_CheckStateChanged(sender As Object, e As EventArgs) _
        Handles chkIgnorerModDesign.CheckStateChanged

        For i As Integer = 0 To Me.lbFichiers.Items.Count - 1
            ' Ne pas sélectionner les .Designer si demandé
            If Me.lbFichiers.Items(i).ToString.ToLower.IndexOf(".designer") = -1 Then Continue For
            Me.lbFichiers.SetSelected(i, Not Me.chkIgnorerModDesign.Checked)
        Next

    End Sub

    Private Sub chkMulti_CheckStateChanged(eventSender As Object,
        eventArgs As EventArgs) Handles chkMulti.CheckStateChanged

        chkFrames.Enabled = (chkMulti.Enabled And
            chkMulti.CheckState = Windows.Forms.CheckState.Checked)

    End Sub

    Private Sub chkTDM_CheckStateChanged(eventSender As Object,
        eventArgs As EventArgs) Handles chkTDM.CheckStateChanged

        chkMulti.Enabled = (chkTDM.CheckState = Windows.Forms.CheckState.Checked)
        chkCompDec.Enabled = (chkTDM.CheckState = Windows.Forms.CheckState.Checked)

        chkMulti_CheckStateChanged(chkMulti, New EventArgs())

    End Sub

    Private Sub chkSelectTousLesFichiers_CheckStateChanged(sender As Object,
        e As EventArgs) Handles chkSelectTousLesFichiers.CheckStateChanged
        SelectionnerFichiers()
    End Sub

    Private Sub SelectionnerFichiers()

        Dim i%
        If Me.chkSelectTousLesFichiers.Checked Then

            For i = 0 To Me.lbFichiers.Items.Count - 1
                ' Ne pas sélectionner les .Designer si demandée
                If Me.chkIgnorerModDesign.Checked And
                   Me.lbFichiers.Items(i).ToString.IndexOf(".Designer") > 0 Then
                    Me.lbFichiers.SetSelected(i, False)
                Else
                    Me.lbFichiers.SetSelected(i, True)
                End If
            Next

        Else 'ChkAllFichier.CheckState = False

            For i = 0 To Me.lbFichiers.Items.Count - 1
                Me.lbFichiers.SetSelected(i, False)
            Next

        End If

    End Sub

    Private Sub cmdHaut_Click(eventSender As Object, eventArgs As EventArgs) Handles cmdHaut.Click

        Dim iIndexSel% = Me.lbFichiers.SelectedIndex
        If iIndexSel <= 0 Then Exit Sub
        If Not bUneSeuleSelection(Me.lbFichiers) Then Exit Sub

        Dim bSel As Boolean = Me.lbFichiers.GetSelected(iIndexSel)
        Dim TempList$ = Me.lbFichiers.Items(iIndexSel).ToString
        Me.lbFichiers.Items(iIndexSel) = Me.lbFichiers.Items(iIndexSel - 1)
        Me.lbFichiers.Items(iIndexSel - 1) = TempList

        Me.lbFichiers.SetSelected(iIndexSel - 1, bSel)
        Me.lbFichiers.SetSelected(iIndexSel, False)

    End Sub

    Private Sub cmdBas_Click(eventSender As Object, eventArgs As EventArgs) Handles cmdBas.Click

        Dim iIndexSel% = Me.lbFichiers.SelectedIndex
        If iIndexSel < 0 Then Exit Sub
        If iIndexSel >= Me.lbFichiers.Items.Count - 1 Then Exit Sub
        If Not bUneSeuleSelection(Me.lbFichiers) Then Exit Sub

        Dim bSel As Boolean = Me.lbFichiers.GetSelected(iIndexSel)
        Dim TempList$ = Me.lbFichiers.Items(iIndexSel).ToString()
        Me.lbFichiers.Items(iIndexSel) = Me.lbFichiers.Items(iIndexSel + 1)
        Me.lbFichiers.Items(iIndexSel + 1) = TempList

        Me.lbFichiers.SetSelected(iIndexSel + 1, bSel)
        Me.lbFichiers.SetSelected(iIndexSel, False)

    End Sub

    Private Function bUneSeuleSelection(lb As Windows.Forms.ListBox) As Boolean

        Dim i%, iNbSel%
        bUneSeuleSelection = False
        For i = 0 To lb.Items.Count - 1
            If lb.GetSelected(i) Then iNbSel += 1
            If iNbSel = 1 Then Return True
            If iNbSel > 1 Then Return False
        Next

    End Function

    Private Sub cmdParcourirHTMLDest_Click(eventSender As Object,
        eventArgs As EventArgs) Handles cmdParcourirHTMLDest.Click

        Dim sCheminFichier$ = ""
        If Me.txtCheminProjetAConv.Text.Length > 0 Then
            sCheminFichier = Me.txtCheminProjetAConv.Text & sExtHtml
        End If
        If Not bChoisirFichier(sCheminFichier, sFiltreHtml, sExtHtml,
            "Choisissez un fichier Html de destination", bDoitExister:=False) Then Exit Sub
        Me.txtCheminHtmlDest.Text = sCheminFichier

    End Sub

    Private Sub cmdParcourirModele_Click(eventSender As Object,
        eventArgs As EventArgs) Handles cmdParcourirModeleHtml.Click

        Dim sCheminFichier$ = Me.txtCheminModeleHtml.Text
        If Not bChoisirFichier(sCheminFichier, sFiltreHtml, sExtHtml,
            "Choisissez un fichier modèle Html") Then Exit Sub
        Me.txtCheminModeleHtml.Text = sCheminFichier
        My.Settings.CheminFichierModeleHtml = Me.txtCheminModeleHtml.Text

    End Sub

    Private Sub cmdParcourirProj_Click(eventSender As Object,
        eventArgs As EventArgs) Handles cmdParcourirProj.Click

        Dim sCheminFichier$ = Me.txtCheminProjetAConv.Text
        If Not bChoisirFichier(sCheminFichier,
            "Fichiers Visual Basic (*.vbp, *.vbproj, *.frm, *.bas, *.cls, *.ctl, *.pag, *.vb, *.txt)|" &
            "*.vbp;*.vbproj;*.frm;*.bas;*.cls;*.ctl;*.pag;*.vb;*.txt|Tous les fichiers (*.*)|*.*",
            "*.*", "Choisissez un fichier de VB6 ou VB.Net") Then Exit Sub
        Me.txtCheminProjetAConv.Text = sCheminFichier
        VerifierFichierProjet(bInitHtmlDest:=True, bNouvProjet:=True)

    End Sub

#End Region

#Region "Gestion des menus contextuels"

    Private Sub VerifierMenuCtx()

        Dim sCleDescriptionCmd$ = sMenuCtx_TypeFichierVB6 & "\shell\" & sMenuCtx_CleCmd
        Dim sCleDescriptionCmd1$ = sMenuCtx_TypeFichierVB2003 & "\shell\" & sMenuCtx_CleCmd
        Dim sCleDescriptionCmd2$ = sMenuCtx_TypeFichierVB2005 & "\shell\" & sMenuCtx_CleCmd
        Dim sCleDescriptionCmd3$ = sMenuCtx_TypeFichierVB2008 & "\shell\" & sMenuCtx_CleCmd
        Dim sCleDescriptionCmd4$ = sMenuCtx_TypeFichierVB2010 & "\shell\" & sMenuCtx_CleCmd
        'Dim sCleDescriptionCmd5$ = sMenuCtx_TypeFichierVB2022 & "\shell\" & sMenuCtx_CleCmd
        Dim sCleDescriptionCmdTous$ = sMenuCtx_TypeFichierTous & "\shell\" & sMenuCtx_CleCmd

        Me.cmdAjouterMenuCtx.Enabled = False
        Me.chkToutFichier.Enabled = False
        Me.cmdEnleverMenuCtx.Enabled = True
        Dim bOk As Boolean = True
        If Not bCleRegistreCRExiste(sCleDescriptionCmd) Then bOk = False
        If Not bCleRegistreCRExiste(sCleDescriptionCmd1) Then bOk = False
        If Not bCleRegistreCRExiste(sCleDescriptionCmd2) Then bOk = False
        If Not bCleRegistreCRExiste(sCleDescriptionCmd3) Then bOk = False
        If Not bCleRegistreCRExiste(sCleDescriptionCmd4) Then bOk = False
        'If Not bCleRegistreCRExiste(sCleDescriptionCmd5) Then bOk = False
        If bCleRegistreCRExiste(sCleDescriptionCmdTous) Then
            Me.chkToutFichier.Checked = True
        Else
            Me.chkToutFichier.Checked = False
        End If
        If Not bOk Then
            Me.cmdAjouterMenuCtx.Enabled = True
            Me.chkToutFichier.Enabled = True
            Me.cmdEnleverMenuCtx.Enabled = False
        End If

    End Sub

    Private Sub cmdAjouterMenuCtx_Click(sender As Object, e As EventArgs) _
        Handles cmdAjouterMenuCtx.Click

        Dim sCheminExe$ = Application.ExecutablePath
        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB6, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, sCheminExe:=sCheminExe)

        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2003, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, sCheminExe:=sCheminExe)
        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2005, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, sCheminExe:=sCheminExe)
        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2008, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, sCheminExe:=sCheminExe)
        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2010, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, sCheminExe:=sCheminExe)
        'bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2022, sMenuCtx_CleCmd, bPrompt:=False,
        '    sDescriptionCmd:=sMenuCtx_DescriptionCmd, sCheminExe:=sCheminExe)
        If Me.chkToutFichier.Checked Then
            bAjouterMenuContextuel(sMenuCtx_TypeFichierTous, sMenuCtx_CleCmd, bPrompt:=False,
                sDescriptionCmd:=sMenuCtx_DescriptionCmd, sCheminExe:=sCheminExe)
        End If

        VerifierMenuCtx()

    End Sub

    Private Sub cmdEnleverMenuCtx_Click(sender As Object, e As EventArgs) _
        Handles cmdEnleverMenuCtx.Click

        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB6, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, bEnlever:=True)

        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2003, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, bEnlever:=True)
        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2005, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, bEnlever:=True)
        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2008, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, bEnlever:=True)
        bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2010, sMenuCtx_CleCmd, bPrompt:=False,
            sDescriptionCmd:=sMenuCtx_DescriptionCmd, bEnlever:=True)
        'bAjouterMenuContextuel(sMenuCtx_TypeFichierVB2022, sMenuCtx_CleCmd, bPrompt:=False,
        '    sDescriptionCmd:=sMenuCtx_DescriptionCmd, bEnlever:=True)
        If Me.chkToutFichier.Checked Then
            bAjouterMenuContextuel(sMenuCtx_TypeFichierTous, sMenuCtx_CleCmd, bPrompt:=False,
                sDescriptionCmd:=sMenuCtx_DescriptionCmd, bEnlever:=True)
        End If

        VerifierMenuCtx()

    End Sub

#End Region

#Region "Création"

    Private Sub VerifierFichierProjet(
        Optional bInitHtmlDest As Boolean = False,
        Optional bNouvProjet As Boolean = False)

        Dim sCheminPrj$ = Me.txtCheminProjetAConv.Text
        If sCheminPrj.Length = 0 Then Exit Sub
        If Not bFichierExiste(Me.txtCheminProjetAConv.Text) Then Exit Sub

        If IsProjectFile(Me.txtCheminProjetAConv.Text) Then
            AnalyserFichierProjet(bNouvProjet)
            gbFichiersATraiter.Enabled = True
            If bInitHtmlDest Then Me.txtCheminHtmlDest.Text = sCheminPrj & ".html"
        Else
            Me.lbFichiers.Items.Clear()
            gbFichiersATraiter.Enabled = False
        End If

        chkTDM_CheckStateChanged(Me.chkTDM, New EventArgs())

        If Not bLireOrdreFichiers() Then SelectionnerFichiers()

    End Sub

    Private Sub AnalyserFichierProjet(bNouvProjet As Boolean)

        Dim sCheminIni0$ = sCheminIni()
        If bNouvProjet AndAlso sCheminIni0.Length > 0 Then
            If bFichierExiste(sCheminIni0) Then
                ' S'il y a un fichier ini, le supprimer si on change de projet
                If Not bSupprimerFichier(sCheminIni0, bPromptErr:=True) Then Exit Sub
            End If
        End If

        Dim AssemblyInfo, AfterEgal, TempLine, PrevFile, TempTitle As String
        TempTitle = ""
        PrevFile = ""
        Dim MinFile, FilesNum, RevisionVer, MajorVer, EgalPos, MinorVer,
            CatType, ReorgFiles, ReorgFiles2 As Integer
        Dim ProjFiles() As String = Nothing
        If Me.txtCheminProjetAConv.Text <> "" Then

            ' Initialiser le titre du projet avec le nom du projet, par défaut
            ProjectVersion = ""
            ProjectTitle = IO.Path.GetFileNameWithoutExtension(Me.txtCheminProjetAConv.Text)

            Me.lbFichiers.Items.Clear()

            FileOpen(1, Me.txtCheminProjetAConv.Text, OpenMode.Input)
            Do Until EOF(1)
                TempLine = LineInput(1)
                TempLine = Trim(TempLine)

                If GetFileType(Me.txtCheminProjetAConv.Text) = ".vbp" Then

                    ' Projet VB6

                    EgalPos = InStr(1, TempLine, "=")

                    If EgalPos > 0 Then
                        AfterEgal = Mid(TempLine, EgalPos + 1)

                        Select Case LCase(VB.Left(TempLine, EgalPos - 1))
                            Case "title"
                                Dim sTitrePrj$ = Trim(Replace(AfterEgal, Chr(34), ""))
                                If sTitrePrj.Length > 0 Then ProjectTitle = sTitrePrj
                            Case "majorver" : MajorVer = CInt(Val(AfterEgal))
                            Case "minorver" : MinorVer = CInt(Val(AfterEgal))
                            Case "revisionver" : RevisionVer = CInt(Val(AfterEgal))
                            Case "form", "usercontrol", "propertypage"
                                ReDim Preserve ProjFiles(FilesNum)

                                ProjFiles(FilesNum) = Trim(AfterEgal)
                                FilesNum = FilesNum + 1
                            Case "module", "class"
                                ReDim Preserve ProjFiles(FilesNum)

                                ProjFiles(FilesNum) = Trim(Mid(AfterEgal, InStr(1, AfterEgal, ";") + 1))
                                FilesNum = FilesNum + 1
                        End Select
                    End If

                    ProjectVersion = " v" & MajorVer & "." & MinorVer & "." & RevisionVer

                Else

                    ' Projet VB7 = 2003 et 2005

                    If LCase(TempLine) = "<settings" And CatType = 0 Then
                        CatType = 1
                    ElseIf CatType = 1 And LCase(TempLine) = ">" Then
                        CatType = 0
                    ElseIf LCase(TempLine) = "<include>" And CatType = 0 Then
                        CatType = 2
                    ElseIf CatType = 2 And LCase(TempLine) = "</include>" Then
                        CatType = 0
                    Else
                        EgalPos = InStr(1, TempLine, "=")

                        If EgalPos > 0 Then
                            AfterEgal = Mid(TempLine, EgalPos + 1)

                            Dim sType$ = Trim(LCase(VB.Left(TempLine, EgalPos - 1)))
                            If sType = "<compile include" Then ' VB 2005
                                PrevFile = Trim(Replace(AfterEgal, Chr(34), ""))
                                PrevFile = PrevFile.Replace("/>", "").Trim
                                PrevFile = PrevFile.Replace(">", "").Trim
                                CatType = 2
                            End If

                            Select Case sType
                                Case "assemblyname"
                                    If CatType = 1 Then
                                        TempTitle = Trim(Replace(AfterEgal, Chr(34), ""))
                                    End If

                                Case "relpath"
                                    If CatType = 2 Then
                                        PrevFile = Trim(Replace(AfterEgal, Chr(34), ""))
                                    End If

                                Case "subtype", "<compile include"
                                    If CatType = 2 Then
                                        ReDim Preserve ProjFiles(FilesNum)

                                        ProjFiles(FilesNum) = PrevFile
                                        FilesNum = FilesNum + 1

                                        If GetFileName(PrevFile) = "assemblyinfo.vb" Then

                                            Dim sCheminFichier$ = VB.Left(Me.txtCheminProjetAConv.Text, InStrRev(Me.txtCheminProjetAConv.Text, "\")) & PrevFile
                                            AssemblyInfo = sLireFichier(sCheminFichier)
                                            ' Les infos peuvent être vides, ne pas effacer les valeurs par défaut
                                            Dim sTitrePrj$ = GetBoundedString(AssemblyInfo,
                                            "<Assembly: AssemblyTitle(""", """)>")
                                            If sTitrePrj.Length > 0 Then ProjectTitle = sTitrePrj
                                            Dim sVersion$ = GetBoundedString(AssemblyInfo,
                                            "<Assembly: AssemblyVersion(""", """)>")
                                            If sVersion.Length > 0 Then ProjectVersion = " v" & sVersion
                                        End If

                                        PrevFile = ""
                                    End If
                                Case "<file" : PrevFile = ""
                            End Select
                        End If
                    End If
                End If
            Loop
            FileClose(1)

            If ProjectTitle = "" Then ProjectTitle = TempTitle

            For ReorgFiles = 0 To FilesNum - 1
                MinFile = 0

                For ReorgFiles2 = 1 To FilesNum - 1
                    If (LCase(ProjFiles(ReorgFiles2)) < LCase(ProjFiles(MinFile)) Or
                        ProjFiles(MinFile) = "") And ProjFiles(ReorgFiles2) <> "" Then _
                        MinFile = ReorgFiles2
                Next ReorgFiles2

                Me.lbFichiers.Items.Add(ProjFiles(MinFile))
                Me.lbFichiers.SetSelected(ReorgFiles, True)
                ProjFiles(MinFile) = ""
            Next ReorgFiles
        End If

    End Sub

    Private Sub ShowInterface(ByRef bAfficherInterface As Boolean)

        Me.gbCreation.Visible = bAfficherInterface
        Me.gbFichierProjet.Visible = bAfficherInterface
        Me.gbFichiersATraiter.Visible = bAfficherInterface
        Me.gbFichierHTMLDest.Visible = bAfficherInterface
        Me.gbModeleEtOptions.Visible = bAfficherInterface
        Sablier(bDesactiver:=bAfficherInterface)
        TraiterMsgSysteme_DoEvents()

    End Sub

    Private Sub cmdCreer_Click(eventSender As Object, eventArgs As EventArgs) Handles cmdCreer.Click
        CreerHtml()
    End Sub

    Private Sub CreerHtml()

        ' Relire les infos du projet au cas où on relance l'analyse sans avoir quitter VB2Html
        ' Non car on perd alors l'ordre des fichiers : isoler la lecture du fichier projet
        'VerifierFichierProjet()

        Dim AddFiles, FilesNum As Integer
        Dim FilesList() As String = Nothing
        If Me.txtCheminProjetAConv.Text <> "" Then

            ' 17/08/2009 Vérifier si les fichiers indiqués dans le fichier .ini
            '  existent toujours
            Dim sCheminDossierPrj$ = IO.Path.GetDirectoryName(Me.txtCheminProjetAConv.Text)

            If IsProjectFile(Me.txtCheminProjetAConv.Text) Then
                For AddFiles = 0 To Me.lbFichiers.Items.Count - 1
                    If Me.lbFichiers.GetSelected(AddFiles) Then
                        Dim sFichier$ = Me.lbFichiers.Items(AddFiles).ToString
                        If Me.chkIgnorerModDesign.Checked = True Then
                            If sFichier.IndexOf(".Designer.vb") > 0 Then
                                'Debug.WriteLine(sFichier)
                                GoTo FichierSuivant
                            End If
                        End If
                        Dim sChemin$ = sCheminDossierPrj & "\" & sFichier
                        If Not bFichierExiste(sChemin) Then GoTo FichierSuivant
                        ReDim Preserve FilesList(FilesNum)
                        FilesList(FilesNum) = sFichier
                        FilesNum = FilesNum + 1
                    End If
FichierSuivant:
                Next AddFiles
            End If

            If FilesNum = 0 Then
                MsgBox("Aucun fichier à convertir !", MsgBoxStyle.Information, m_sTitreMsg)
                Exit Sub
            End If

            ShowInterface(False)

            Dim sCheminPrj$ = Me.txtCheminProjetAConv.Text
            Dim bFichierPrj As Boolean = IsProjectFile(sCheminPrj)
            Dim sTitrePrj$ = ""
            Dim sVersionPrj$ = ""
            If IsProjectFile(sCheminPrj) Then
                sTitrePrj = ProjectTitle
                sVersionPrj = ProjectVersion
            End If
            Dim sListeMotCle$ = ""
            If IsFileDotNet(sCheminPrj) Then
                sListeMotCle = KWList_VBNET
            Else
                sListeMotCle = KWList_VB6
            End If
            Dim iTronquer% = 0
            If Me.chkTronc.Checked Then
                iTronquer = iConv(txtTroncLen.Text, 80)
            End If
            Dim bDeclProcComplete As Boolean = False
            bDeclProcComplete = (chkCompDec.Enabled And chkCompDec.Checked)
            Dim bMultiFichiers As Boolean = False
            bMultiFichiers = chkMulti.Enabled And chkMulti.Checked
            Dim bFramesHtml As Boolean = False
            bFramesHtml = chkFrames.Enabled And chkFrames.Checked
            Dim bTDM As Boolean = False
            bTDM = chkTDM.Checked
            Dim bVB2Txt As Boolean = False
            bVB2Txt = Me.chkVB2Txt.Checked

            If bFrmPatienter Then Me.m_patienter.Init(FilesNum, FilesList, Me.txtCheminProjetAConv.Text)

            ChargerModeleHtml()

            TransformToHTML(sTitrePrj, sVersionPrj, sListeMotCle,
                FilesList, FilesNum, Me.txtCheminProjetAConv.Text,
                Me.txtCheminHtmlDest.Text, Me.m_sContenuModeleHtml,
                iTronquer, bDeclProcComplete,
                bMultiFichiers, bFramesHtml, bTDM, bVB2Txt, Me.m_patienter)

            If bFrmPatienter Then Me.m_patienter.Fermer()

            ShowInterface(True)

            If Me.m_patienter.bAnnuler Then Exit Sub

            If bVB2Txt Then
                MsgBox("La conversion est terminée !",
                    MsgBoxStyle.Exclamation, "Opération terminée")
                Exit Sub
            End If

            If MsgBoxResult.Yes = MsgBox(
                "La conversion est terminée. Voulez-vous ouvrir le fichier HTML créé ?",
                MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Opération terminée") Then
                Dim sCheminFichier$ = Me.txtCheminHtmlDest.Text
                OuvrirAppliAssociee(sCheminFichier)
            End If

        End If

    End Sub

#End Region

End Class