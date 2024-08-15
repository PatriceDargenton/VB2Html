<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWait
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
Me.lblMessage2 = New System.Windows.Forms.Label
Me.lblMessage = New System.Windows.Forms.Label
Me.pbPrincipale = New System.Windows.Forms.ProgressBar
Me.pbInterm = New System.Windows.Forms.ProgressBar
Me.cmdAnnuler = New System.Windows.Forms.Button
Me.SuspendLayout()
'
'lblMessage2
'
Me.lblMessage2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
Me.lblMessage2.FlatStyle = System.Windows.Forms.FlatStyle.System
Me.lblMessage2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblMessage2.ForeColor = System.Drawing.Color.Red
Me.lblMessage2.Location = New System.Drawing.Point(8, 29)
Me.lblMessage2.Name = "lblMessage2"
Me.lblMessage2.Size = New System.Drawing.Size(309, 20)
Me.lblMessage2.TabIndex = 27
Me.lblMessage2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'lblMessage
'
Me.lblMessage.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
Me.lblMessage.FlatStyle = System.Windows.Forms.FlatStyle.System
Me.lblMessage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
Me.lblMessage.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer))
Me.lblMessage.Location = New System.Drawing.Point(8, 3)
Me.lblMessage.Name = "lblMessage"
Me.lblMessage.Size = New System.Drawing.Size(309, 20)
Me.lblMessage.TabIndex = 25
Me.lblMessage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
'
'pbPrincipale
'
Me.pbPrincipale.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
Me.pbPrincipale.Location = New System.Drawing.Point(9, 112)
Me.pbPrincipale.Maximum = 10000
Me.pbPrincipale.Name = "pbPrincipale"
Me.pbPrincipale.Size = New System.Drawing.Size(309, 16)
Me.pbPrincipale.TabIndex = 26
'
'pbInterm
'
Me.pbInterm.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
Me.pbInterm.Location = New System.Drawing.Point(9, 134)
Me.pbInterm.Name = "pbInterm"
Me.pbInterm.Size = New System.Drawing.Size(309, 16)
Me.pbInterm.Step = 1
Me.pbInterm.TabIndex = 28
'
'cmdAnnuler
'
Me.cmdAnnuler.Location = New System.Drawing.Point(130, 61)
Me.cmdAnnuler.Name = "cmdAnnuler"
Me.cmdAnnuler.Size = New System.Drawing.Size(67, 31)
Me.cmdAnnuler.TabIndex = 29
Me.cmdAnnuler.Text = "Annuler"
Me.cmdAnnuler.UseVisualStyleBackColor = True
'
'frmWait
'
Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
Me.ClientSize = New System.Drawing.Size(327, 158)
Me.ControlBox = False
Me.Controls.Add(Me.cmdAnnuler)
Me.Controls.Add(Me.pbInterm)
Me.Controls.Add(Me.lblMessage2)
Me.Controls.Add(Me.lblMessage)
Me.Controls.Add(Me.pbPrincipale)
Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
Me.MinimumSize = New System.Drawing.Size(320, 80)
Me.Name = "frmWait"
Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
Me.Text = "Veuillez Patienter !"
Me.ResumeLayout(False)

End Sub
    Friend WithEvents lblMessage2 As System.Windows.Forms.Label
    Friend WithEvents lblMessage As System.Windows.Forms.Label
    Friend WithEvents pbPrincipale As System.Windows.Forms.ProgressBar
    Friend WithEvents pbInterm As System.Windows.Forms.ProgressBar
    Friend WithEvents cmdAnnuler As System.Windows.Forms.Button
End Class
