<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmVB2Html : Inherits Form
#Region "Windows Form Designer generated code "
    <System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
    End Sub
    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents chkIgnorerModDesign As Windows.Forms.CheckBox
    Public WithEvents chkFrames As System.Windows.Forms.CheckBox
    Public WithEvents txtTroncLen As System.Windows.Forms.TextBox
    Public WithEvents chkTronc As System.Windows.Forms.CheckBox
    Public WithEvents chkCompDec As System.Windows.Forms.CheckBox
    Public WithEvents chkMulti As System.Windows.Forms.CheckBox
    Public WithEvents chkTDM As System.Windows.Forms.CheckBox
    Public WithEvents cmdParcourirModeleHtml As System.Windows.Forms.Button
    Public WithEvents txtCheminModeleHtml As System.Windows.Forms.TextBox
    Public WithEvents lblChrs As System.Windows.Forms.Label
    Public WithEvents gbModeleEtOptions As System.Windows.Forms.GroupBox
    Public WithEvents cmdCreer As System.Windows.Forms.Button
    Public WithEvents txtCheminHtmlDest As System.Windows.Forms.TextBox
    Public WithEvents cmdParcourirHTMLDest As System.Windows.Forms.Button
    Public WithEvents gbFichierHTMLDest As System.Windows.Forms.GroupBox
    Public WithEvents cmdBas As System.Windows.Forms.Button
    Public WithEvents cmdHaut As System.Windows.Forms.Button
    Public WithEvents lbFichiers As System.Windows.Forms.ListBox
    Public WithEvents gbFichiersATraiter As System.Windows.Forms.GroupBox
    Public WithEvents cmdParcourirProj As System.Windows.Forms.Button
    Public WithEvents txtCheminProjetAConv As System.Windows.Forms.TextBox
    Public WithEvents gbFichierProjet As System.Windows.Forms.GroupBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVB2Html))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdBas = New System.Windows.Forms.Button()
        Me.cmdHaut = New System.Windows.Forms.Button()
        Me.chkIgnorerModDesign = New System.Windows.Forms.CheckBox()
        Me.chkSelectTousLesFichiers = New System.Windows.Forms.CheckBox()
        Me.chkVB2Txt = New System.Windows.Forms.CheckBox()
        Me.chkFrames = New System.Windows.Forms.CheckBox()
        Me.txtCheminModeleHtml = New System.Windows.Forms.TextBox()
        Me.cmdParcourirModeleHtml = New System.Windows.Forms.Button()
        Me.chkTDM = New System.Windows.Forms.CheckBox()
        Me.chkMulti = New System.Windows.Forms.CheckBox()
        Me.cmdCreer = New System.Windows.Forms.Button()
        Me.cmdParcourirHTMLDest = New System.Windows.Forms.Button()
        Me.txtCheminHtmlDest = New System.Windows.Forms.TextBox()
        Me.lbFichiers = New System.Windows.Forms.ListBox()
        Me.cmdParcourirProj = New System.Windows.Forms.Button()
        Me.txtCheminProjetAConv = New System.Windows.Forms.TextBox()
        Me.cmdAjouterMenuCtx = New System.Windows.Forms.Button()
        Me.cmdEnleverMenuCtx = New System.Windows.Forms.Button()
        Me.gbModeleEtOptions = New System.Windows.Forms.GroupBox()
        Me.lblChrs = New System.Windows.Forms.Label()
        Me.txtTroncLen = New System.Windows.Forms.TextBox()
        Me.chkTronc = New System.Windows.Forms.CheckBox()
        Me.chkCompDec = New System.Windows.Forms.CheckBox()
        Me.gbFichierHTMLDest = New System.Windows.Forms.GroupBox()
        Me.gbFichiersATraiter = New System.Windows.Forms.GroupBox()
        Me.gbFichierProjet = New System.Windows.Forms.GroupBox()
        Me.gbCreation = New System.Windows.Forms.GroupBox()
        Me.chkToutFichier = New System.Windows.Forms.CheckBox()
        Me.gbModeleEtOptions.SuspendLayout()
        Me.gbFichierHTMLDest.SuspendLayout()
        Me.gbFichiersATraiter.SuspendLayout()
        Me.gbFichierProjet.SuspendLayout()
        Me.gbCreation.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdBas
        '
        Me.cmdBas.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBas.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBas.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBas.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBas.Location = New System.Drawing.Point(211, 330)
        Me.cmdBas.Name = "cmdBas"
        Me.cmdBas.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBas.Size = New System.Drawing.Size(65, 25)
        Me.cmdBas.TabIndex = 23
        Me.cmdBas.Text = "Bas"
        Me.ToolTip1.SetToolTip(Me.cmdBas, "Sélectionnez un seul fichier à descendre")
        Me.cmdBas.UseVisualStyleBackColor = False
        '
        'cmdHaut
        '
        Me.cmdHaut.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdHaut.BackColor = System.Drawing.SystemColors.Control
        Me.cmdHaut.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdHaut.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdHaut.Location = New System.Drawing.Point(211, 301)
        Me.cmdHaut.Name = "cmdHaut"
        Me.cmdHaut.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdHaut.Size = New System.Drawing.Size(65, 25)
        Me.cmdHaut.TabIndex = 22
        Me.cmdHaut.Text = "Haut"
        Me.ToolTip1.SetToolTip(Me.cmdHaut, "Sélectionnez un seul fichier à monter")
        Me.cmdHaut.UseVisualStyleBackColor = False
        '
        'chkIgnorerModDesign
        '
        Me.chkIgnorerModDesign.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkIgnorerModDesign.BackColor = System.Drawing.SystemColors.Control
        Me.chkIgnorerModDesign.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkIgnorerModDesign.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkIgnorerModDesign.Location = New System.Drawing.Point(27, 304)
        Me.chkIgnorerModDesign.Name = "chkIgnorerModDesign"
        Me.chkIgnorerModDesign.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkIgnorerModDesign.Size = New System.Drawing.Size(178, 22)
        Me.chkIgnorerModDesign.TabIndex = 29
        Me.chkIgnorerModDesign.Text = "Ignorer les .Designer.vb"
        Me.ToolTip1.SetToolTip(Me.chkIgnorerModDesign, "Ignorer les modules de désign de VB >= 2005")
        Me.chkIgnorerModDesign.UseVisualStyleBackColor = False
        '
        'chkSelectTousLesFichiers
        '
        Me.chkSelectTousLesFichiers.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkSelectTousLesFichiers.BackColor = System.Drawing.SystemColors.Control
        Me.chkSelectTousLesFichiers.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkSelectTousLesFichiers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSelectTousLesFichiers.Location = New System.Drawing.Point(27, 331)
        Me.chkSelectTousLesFichiers.Name = "chkSelectTousLesFichiers"
        Me.chkSelectTousLesFichiers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkSelectTousLesFichiers.Size = New System.Drawing.Size(178, 18)
        Me.chkSelectTousLesFichiers.TabIndex = 25
        Me.chkSelectTousLesFichiers.Text = "Sélectionner tous les fichiers"
        Me.ToolTip1.SetToolTip(Me.chkSelectTousLesFichiers, "Sélectionner tous les fichiers")
        Me.chkSelectTousLesFichiers.UseVisualStyleBackColor = False
        '
        'chkVB2Txt
        '
        Me.chkVB2Txt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkVB2Txt.BackColor = System.Drawing.SystemColors.Control
        Me.chkVB2Txt.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkVB2Txt.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkVB2Txt.Location = New System.Drawing.Point(24, 75)
        Me.chkVB2Txt.Name = "chkVB2Txt"
        Me.chkVB2Txt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkVB2Txt.Size = New System.Drawing.Size(77, 22)
        Me.chkVB2Txt.TabIndex = 30
        Me.chkVB2Txt.Text = "VB2Txt"
        Me.ToolTip1.SetToolTip(Me.chkVB2Txt, "Exporter en .txt au lieu de .Html")
        Me.chkVB2Txt.UseVisualStyleBackColor = False
        '
        'chkFrames
        '
        Me.chkFrames.BackColor = System.Drawing.SystemColors.Control
        Me.chkFrames.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFrames.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFrames.Location = New System.Drawing.Point(24, 102)
        Me.chkFrames.Name = "chkFrames"
        Me.chkFrames.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFrames.Size = New System.Drawing.Size(225, 17)
        Me.chkFrames.TabIndex = 28
        Me.chkFrames.Text = "Créer les pages en frames HTML"
        Me.ToolTip1.SetToolTip(Me.chkFrames, "Créer un menu permettant d'accéder à chaque fichier source html dans un cadre")
        Me.chkFrames.UseVisualStyleBackColor = False
        '
        'txtCheminModeleHtml
        '
        Me.txtCheminModeleHtml.AcceptsReturn = True
        Me.txtCheminModeleHtml.BackColor = System.Drawing.SystemColors.Window
        Me.txtCheminModeleHtml.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCheminModeleHtml.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCheminModeleHtml.Location = New System.Drawing.Point(8, 19)
        Me.txtCheminModeleHtml.MaxLength = 0
        Me.txtCheminModeleHtml.Name = "txtCheminModeleHtml"
        Me.txtCheminModeleHtml.ReadOnly = True
        Me.txtCheminModeleHtml.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCheminModeleHtml.Size = New System.Drawing.Size(257, 20)
        Me.txtCheminModeleHtml.TabIndex = 15
        Me.ToolTip1.SetToolTip(Me.txtCheminModeleHtml, "Chemin du modèle de présentation Html")
        '
        'cmdParcourirModeleHtml
        '
        Me.cmdParcourirModeleHtml.BackColor = System.Drawing.SystemColors.Control
        Me.cmdParcourirModeleHtml.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdParcourirModeleHtml.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdParcourirModeleHtml.Location = New System.Drawing.Point(271, 19)
        Me.cmdParcourirModeleHtml.Name = "cmdParcourirModeleHtml"
        Me.cmdParcourirModeleHtml.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdParcourirModeleHtml.Size = New System.Drawing.Size(25, 20)
        Me.cmdParcourirModeleHtml.TabIndex = 16
        Me.cmdParcourirModeleHtml.Text = "..."
        Me.cmdParcourirModeleHtml.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.cmdParcourirModeleHtml, "Choisir un autre fichier modèle de présentation html")
        Me.cmdParcourirModeleHtml.UseVisualStyleBackColor = False
        '
        'chkTDM
        '
        Me.chkTDM.BackColor = System.Drawing.SystemColors.Control
        Me.chkTDM.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTDM.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTDM.Location = New System.Drawing.Point(8, 54)
        Me.chkTDM.Name = "chkTDM"
        Me.chkTDM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTDM.Size = New System.Drawing.Size(233, 17)
        Me.chkTDM.TabIndex = 17
        Me.chkTDM.Text = "Générer la table des matières"
        Me.ToolTip1.SetToolTip(Me.chkTDM, "Générer la table des matières basée sur les procédures du projet")
        Me.chkTDM.UseVisualStyleBackColor = False
        '
        'chkMulti
        '
        Me.chkMulti.BackColor = System.Drawing.SystemColors.Control
        Me.chkMulti.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkMulti.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkMulti.Location = New System.Drawing.Point(16, 78)
        Me.chkMulti.Name = "chkMulti"
        Me.chkMulti.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkMulti.Size = New System.Drawing.Size(233, 17)
        Me.chkMulti.TabIndex = 18
        Me.chkMulti.Text = "Créer plusieurs fichiers HTML"
        Me.ToolTip1.SetToolTip(Me.chkMulti, "Créer un fichier html par fichier source")
        Me.chkMulti.UseVisualStyleBackColor = False
        '
        'cmdCreer
        '
        Me.cmdCreer.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdCreer.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCreer.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCreer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCreer.Location = New System.Drawing.Point(107, 72)
        Me.cmdCreer.Name = "cmdCreer"
        Me.cmdCreer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCreer.Size = New System.Drawing.Size(81, 25)
        Me.cmdCreer.TabIndex = 11
        Me.cmdCreer.Text = "Créer"
        Me.ToolTip1.SetToolTip(Me.cmdCreer, "Lancer la création du rapport html")
        Me.cmdCreer.UseVisualStyleBackColor = False
        '
        'cmdParcourirHTMLDest
        '
        Me.cmdParcourirHTMLDest.BackColor = System.Drawing.SystemColors.Control
        Me.cmdParcourirHTMLDest.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdParcourirHTMLDest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdParcourirHTMLDest.Location = New System.Drawing.Point(271, 20)
        Me.cmdParcourirHTMLDest.Name = "cmdParcourirHTMLDest"
        Me.cmdParcourirHTMLDest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdParcourirHTMLDest.Size = New System.Drawing.Size(25, 19)
        Me.cmdParcourirHTMLDest.TabIndex = 9
        Me.cmdParcourirHTMLDest.Text = "..."
        Me.cmdParcourirHTMLDest.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.cmdParcourirHTMLDest, "Choisir un autre chemin de destination pour le rapport html")
        Me.cmdParcourirHTMLDest.UseVisualStyleBackColor = False
        '
        'txtCheminHtmlDest
        '
        Me.txtCheminHtmlDest.AcceptsReturn = True
        Me.txtCheminHtmlDest.BackColor = System.Drawing.SystemColors.Window
        Me.txtCheminHtmlDest.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCheminHtmlDest.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCheminHtmlDest.Location = New System.Drawing.Point(8, 19)
        Me.txtCheminHtmlDest.MaxLength = 0
        Me.txtCheminHtmlDest.Name = "txtCheminHtmlDest"
        Me.txtCheminHtmlDest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCheminHtmlDest.Size = New System.Drawing.Size(257, 20)
        Me.txtCheminHtmlDest.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.txtCheminHtmlDest, "Chemin du fichier html de destination")
        '
        'lbFichiers
        '
        Me.lbFichiers.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbFichiers.BackColor = System.Drawing.SystemColors.Window
        Me.lbFichiers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lbFichiers.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lbFichiers.IntegralHeight = False
        Me.lbFichiers.Location = New System.Drawing.Point(6, 20)
        Me.lbFichiers.Name = "lbFichiers"
        Me.lbFichiers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lbFichiers.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lbFichiers.Size = New System.Drawing.Size(293, 275)
        Me.lbFichiers.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.lbFichiers, "Liste des fichiers du projet")
        '
        'cmdParcourirProj
        '
        Me.cmdParcourirProj.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdParcourirProj.BackColor = System.Drawing.SystemColors.Control
        Me.cmdParcourirProj.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdParcourirProj.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdParcourirProj.Location = New System.Drawing.Point(580, 19)
        Me.cmdParcourirProj.Name = "cmdParcourirProj"
        Me.cmdParcourirProj.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdParcourirProj.Size = New System.Drawing.Size(26, 20)
        Me.cmdParcourirProj.TabIndex = 4
        Me.cmdParcourirProj.Text = "..."
        Me.cmdParcourirProj.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.ToolTip1.SetToolTip(Me.cmdParcourirProj, "Choisir un autre fichier projet à convertir")
        Me.cmdParcourirProj.UseVisualStyleBackColor = False
        '
        'txtCheminProjetAConv
        '
        Me.txtCheminProjetAConv.AcceptsReturn = True
        Me.txtCheminProjetAConv.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtCheminProjetAConv.BackColor = System.Drawing.SystemColors.Window
        Me.txtCheminProjetAConv.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCheminProjetAConv.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCheminProjetAConv.Location = New System.Drawing.Point(8, 19)
        Me.txtCheminProjetAConv.MaxLength = 0
        Me.txtCheminProjetAConv.Name = "txtCheminProjetAConv"
        Me.txtCheminProjetAConv.ReadOnly = True
        Me.txtCheminProjetAConv.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCheminProjetAConv.Size = New System.Drawing.Size(566, 20)
        Me.txtCheminProjetAConv.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtCheminProjetAConv, "Chemin du projet à exporter en html")
        '
        'cmdAjouterMenuCtx
        '
        Me.cmdAjouterMenuCtx.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAjouterMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAjouterMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAjouterMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAjouterMenuCtx.Location = New System.Drawing.Point(24, 28)
        Me.cmdAjouterMenuCtx.Name = "cmdAjouterMenuCtx"
        Me.cmdAjouterMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAjouterMenuCtx.Size = New System.Drawing.Size(103, 25)
        Me.cmdAjouterMenuCtx.TabIndex = 31
        Me.cmdAjouterMenuCtx.Text = "Ajouter menu ctx."
        Me.ToolTip1.SetToolTip(Me.cmdAjouterMenuCtx, "Ajouter un menu contextuel pour convertir directement un projet depuis l'explorat" & _
        "eur de fichiers")
        Me.cmdAjouterMenuCtx.UseVisualStyleBackColor = False
        '
        'cmdEnleverMenuCtx
        '
        Me.cmdEnleverMenuCtx.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdEnleverMenuCtx.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEnleverMenuCtx.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEnleverMenuCtx.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEnleverMenuCtx.Location = New System.Drawing.Point(169, 28)
        Me.cmdEnleverMenuCtx.Name = "cmdEnleverMenuCtx"
        Me.cmdEnleverMenuCtx.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEnleverMenuCtx.Size = New System.Drawing.Size(103, 25)
        Me.cmdEnleverMenuCtx.TabIndex = 32
        Me.cmdEnleverMenuCtx.Text = "Enlever menu ctx."
        Me.ToolTip1.SetToolTip(Me.cmdEnleverMenuCtx, "Enlever le menu contextuel")
        Me.cmdEnleverMenuCtx.UseVisualStyleBackColor = False
        '
        'gbModeleEtOptions
        '
        Me.gbModeleEtOptions.BackColor = System.Drawing.SystemColors.Control
        Me.gbModeleEtOptions.Controls.Add(Me.chkFrames)
        Me.gbModeleEtOptions.Controls.Add(Me.txtCheminModeleHtml)
        Me.gbModeleEtOptions.Controls.Add(Me.lblChrs)
        Me.gbModeleEtOptions.Controls.Add(Me.txtTroncLen)
        Me.gbModeleEtOptions.Controls.Add(Me.cmdParcourirModeleHtml)
        Me.gbModeleEtOptions.Controls.Add(Me.chkTDM)
        Me.gbModeleEtOptions.Controls.Add(Me.chkTronc)
        Me.gbModeleEtOptions.Controls.Add(Me.chkMulti)
        Me.gbModeleEtOptions.Controls.Add(Me.chkCompDec)
        Me.gbModeleEtOptions.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbModeleEtOptions.Location = New System.Drawing.Point(8, 71)
        Me.gbModeleEtOptions.Name = "gbModeleEtOptions"
        Me.gbModeleEtOptions.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gbModeleEtOptions.Size = New System.Drawing.Size(305, 185)
        Me.gbModeleEtOptions.TabIndex = 13
        Me.gbModeleEtOptions.TabStop = False
        Me.gbModeleEtOptions.Text = "Modèle de fichier HTML et options diverses"
        '
        'lblChrs
        '
        Me.lblChrs.BackColor = System.Drawing.SystemColors.Control
        Me.lblChrs.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblChrs.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblChrs.Location = New System.Drawing.Point(215, 153)
        Me.lblChrs.Name = "lblChrs"
        Me.lblChrs.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblChrs.Size = New System.Drawing.Size(57, 17)
        Me.lblChrs.TabIndex = 27
        Me.lblChrs.Text = "caractères"
        '
        'txtTroncLen
        '
        Me.txtTroncLen.AcceptsReturn = True
        Me.txtTroncLen.BackColor = System.Drawing.SystemColors.Window
        Me.txtTroncLen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtTroncLen.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtTroncLen.Location = New System.Drawing.Point(168, 150)
        Me.txtTroncLen.MaxLength = 3
        Me.txtTroncLen.Name = "txtTroncLen"
        Me.txtTroncLen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtTroncLen.Size = New System.Drawing.Size(41, 20)
        Me.txtTroncLen.TabIndex = 26
        '
        'chkTronc
        '
        Me.chkTronc.BackColor = System.Drawing.SystemColors.Control
        Me.chkTronc.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkTronc.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkTronc.Location = New System.Drawing.Point(8, 150)
        Me.chkTronc.Name = "chkTronc"
        Me.chkTronc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkTronc.Size = New System.Drawing.Size(161, 17)
        Me.chkTronc.TabIndex = 25
        Me.chkTronc.Text = "Tronquer les longues lignes :"
        Me.chkTronc.UseVisualStyleBackColor = False
        '
        'chkCompDec
        '
        Me.chkCompDec.BackColor = System.Drawing.SystemColors.Control
        Me.chkCompDec.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCompDec.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkCompDec.Location = New System.Drawing.Point(16, 126)
        Me.chkCompDec.Name = "chkCompDec"
        Me.chkCompDec.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCompDec.Size = New System.Drawing.Size(273, 17)
        Me.chkCompDec.TabIndex = 24
        Me.chkCompDec.Text = "Afficher les déclarations de procédures complètes"
        Me.chkCompDec.UseVisualStyleBackColor = False
        '
        'gbFichierHTMLDest
        '
        Me.gbFichierHTMLDest.BackColor = System.Drawing.SystemColors.Control
        Me.gbFichierHTMLDest.Controls.Add(Me.cmdParcourirHTMLDest)
        Me.gbFichierHTMLDest.Controls.Add(Me.txtCheminHtmlDest)
        Me.gbFichierHTMLDest.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbFichierHTMLDest.Location = New System.Drawing.Point(8, 262)
        Me.gbFichierHTMLDest.Name = "gbFichierHTMLDest"
        Me.gbFichierHTMLDest.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gbFichierHTMLDest.Size = New System.Drawing.Size(305, 57)
        Me.gbFichierHTMLDest.TabIndex = 7
        Me.gbFichierHTMLDest.TabStop = False
        Me.gbFichierHTMLDest.Text = "Fichier HTML de destination"
        '
        'gbFichiersATraiter
        '
        Me.gbFichiersATraiter.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbFichiersATraiter.BackColor = System.Drawing.SystemColors.Control
        Me.gbFichiersATraiter.Controls.Add(Me.cmdBas)
        Me.gbFichiersATraiter.Controls.Add(Me.lbFichiers)
        Me.gbFichiersATraiter.Controls.Add(Me.cmdHaut)
        Me.gbFichiersATraiter.Controls.Add(Me.chkSelectTousLesFichiers)
        Me.gbFichiersATraiter.Controls.Add(Me.chkIgnorerModDesign)
        Me.gbFichiersATraiter.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbFichiersATraiter.Location = New System.Drawing.Point(326, 71)
        Me.gbFichiersATraiter.Name = "gbFichiersATraiter"
        Me.gbFichiersATraiter.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gbFichiersATraiter.Size = New System.Drawing.Size(305, 372)
        Me.gbFichiersATraiter.TabIndex = 1
        Me.gbFichiersATraiter.TabStop = False
        Me.gbFichiersATraiter.Text = "Fichiers du projet (Forms, Modules, Classes)"
        '
        'gbFichierProjet
        '
        Me.gbFichierProjet.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbFichierProjet.BackColor = System.Drawing.SystemColors.Control
        Me.gbFichierProjet.Controls.Add(Me.cmdParcourirProj)
        Me.gbFichierProjet.Controls.Add(Me.txtCheminProjetAConv)
        Me.gbFichierProjet.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gbFichierProjet.Location = New System.Drawing.Point(8, 8)
        Me.gbFichierProjet.Name = "gbFichierProjet"
        Me.gbFichierProjet.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.gbFichierProjet.Size = New System.Drawing.Size(623, 57)
        Me.gbFichierProjet.TabIndex = 0
        Me.gbFichierProjet.TabStop = False
        Me.gbFichierProjet.Text = "Fichier principal du projet (.vbp ou .vbproj)"
        '
        'gbCreation
        '
        Me.gbCreation.Controls.Add(Me.chkToutFichier)
        Me.gbCreation.Controls.Add(Me.cmdAjouterMenuCtx)
        Me.gbCreation.Controls.Add(Me.chkVB2Txt)
        Me.gbCreation.Controls.Add(Me.cmdCreer)
        Me.gbCreation.Controls.Add(Me.cmdEnleverMenuCtx)
        Me.gbCreation.Location = New System.Drawing.Point(8, 330)
        Me.gbCreation.Name = "gbCreation"
        Me.gbCreation.Size = New System.Drawing.Size(305, 113)
        Me.gbCreation.TabIndex = 33
        Me.gbCreation.TabStop = False
        Me.gbCreation.Text = "Création"
        '
        'chkToutFichier
        '
        Me.chkToutFichier.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkToutFichier.BackColor = System.Drawing.SystemColors.Control
        Me.chkToutFichier.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkToutFichier.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkToutFichier.Location = New System.Drawing.Point(132, 28)
        Me.chkToutFichier.Name = "chkToutFichier"
        Me.chkToutFichier.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkToutFichier.Size = New System.Drawing.Size(31, 22)
        Me.chkToutFichier.TabIndex = 33
        Me.chkToutFichier.Text = "*"
        Me.ToolTip1.SetToolTip(Me.chkToutFichier, "Ajouter le menu pour tous les fichiers (si le fichier projet n'est pas reconnu)")
        Me.chkToutFichier.UseVisualStyleBackColor = False
        '
        'frmVB2Html
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(643, 455)
        Me.Controls.Add(Me.gbCreation)
        Me.Controls.Add(Me.gbModeleEtOptions)
        Me.Controls.Add(Me.gbFichierHTMLDest)
        Me.Controls.Add(Me.gbFichiersATraiter)
        Me.Controls.Add(Me.gbFichierProjet)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 43)
        Me.Name = "frmVB2Html"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "VB2Html"
        Me.gbModeleEtOptions.ResumeLayout(False)
        Me.gbModeleEtOptions.PerformLayout()
        Me.gbFichierHTMLDest.ResumeLayout(False)
        Me.gbFichierHTMLDest.PerformLayout()
        Me.gbFichiersATraiter.ResumeLayout(False)
        Me.gbFichierProjet.ResumeLayout(False)
        Me.gbFichierProjet.PerformLayout()
        Me.gbCreation.ResumeLayout(False)
        Me.ResumeLayout(False)

End Sub
    Public WithEvents chkSelectTousLesFichiers As System.Windows.Forms.CheckBox
    Public WithEvents chkVB2Txt As System.Windows.Forms.CheckBox
    Public WithEvents cmdAjouterMenuCtx As System.Windows.Forms.Button
    Public WithEvents cmdEnleverMenuCtx As System.Windows.Forms.Button
    Friend WithEvents gbCreation As System.Windows.Forms.GroupBox
    Public WithEvents chkToutFichier As System.Windows.Forms.CheckBox
#End Region
End Class
