
' Fichier frmWait.vb : Formulaire et classe pour patienter
' ------------------

Friend Class frmWait ' Laisser ces lignes ici pour pouvoir �diter le frm

    Public m_bAnnuler As Boolean

    Private Sub cmdAnnuler_Click(sender As Object, e As EventArgs) Handles cmdAnnuler.Click

        If MsgBoxResult.Yes = MsgBox(
            "Etes-vous s�r de vouloir interrompre l'op�ration en cours ?",
            MsgBoxStyle.Question Or MsgBoxStyle.YesNo) Then m_bAnnuler = True

    End Sub

End Class

#Region "Class Patienter"

Public Class clsPatienter

    Private m_iNumFichierEnCours%
    Private m_lLongFichier&()
    Private m_lLongTotFichierEnCours&, m_lLongTotFichierEnCoursEM1&

    Private m_lLongTotFichiers&
    Private m_asFichiers$()
    ' Ne pas oublier WithEvents pour pouvoir traiter l'�v. frmWait.cmdAnnuler.Click
    Private WithEvents m_frmWait As frmWait

    Private m_iPos% = 0
    Private m_iDirection% = 1

    Public ReadOnly Property bAnnuler() As Boolean
        Get
            bAnnuler = Me.m_frmWait.m_bAnnuler
        End Get
    End Property

    Public Sub Init(iNbFichiers%, asFichiers$(), sCheminFichierRef$)

        Me.m_iNumFichierEnCours = 0
        Me.m_lLongTotFichierEnCours = 0
        Me.m_lLongTotFichierEnCoursEM1 = 0
        Me.m_asFichiers = asFichiers
        Dim sCheminDossierRef$ = IO.Path.GetDirectoryName(sCheminFichierRef)
        Me.m_lLongTotFichiers = lTailleTotalFichiers(iNbFichiers,
            sCheminDossierRef, Me.m_asFichiers)

        Me.m_iPos = 0
        Me.m_iDirection = 1

        Me.m_frmWait = New frmWait
        Me.m_frmWait.pbPrincipale.Maximum = CInt(Me.m_lLongTotFichiers)
        Me.m_frmWait.StartPosition = FormStartPosition.CenterScreen
        Me.m_frmWait.Show()

    End Sub

    Public Sub Fermer()
        Me.m_frmWait.Close()
    End Sub

    Public Sub MettreAJour(iNumFichierEnCours%)

        Me.m_iNumFichierEnCours = iNumFichierEnCours

        ' Attendre d'avoir termin� une �tape pour positionner l'avancement
        If iNumFichierEnCours > 0 Then _
            Me.m_lLongTotFichierEnCoursEM1 += Me.m_lLongFichier(iNumFichierEnCours - 1)

        Me.m_frmWait.lblMessage.Text = Me.m_asFichiers(Me.m_iNumFichierEnCours)
        Me.m_frmWait.lblMessage2.Text =
            "Total : " & Me.m_iNumFichierEnCours + 1 & " / " & Me.m_asFichiers.Length
        MettreAJour_Interm(0)

    End Sub

    Public Sub MettreAJour_Interm(iPosFichierEnCours%)

        ' Mise � jour interm�diaire : d�filement ping-pong de gauche � droite et inv.

        ' La longueur pass�e en argument est celle du Html, 
        '  donc un peu > � celle du fichier d'origine
        If iPosFichierEnCours > Me.m_lLongFichier(Me.m_iNumFichierEnCours) Then
            iPosFichierEnCours = CInt(Me.m_lLongFichier(Me.m_iNumFichierEnCours))
        End If
        Me.m_lLongTotFichierEnCours = iPosFichierEnCours
        ' Attendre d'avoir termin� une �tape pour positionner l'avancement
        If Me.m_iNumFichierEnCours > 0 Then
            Me.m_lLongTotFichierEnCours = Me.m_lLongTotFichierEnCoursEM1 + iPosFichierEnCours
        End If

        Me.m_frmWait.pbPrincipale.Value = CInt(Me.m_lLongTotFichierEnCours)
        Me.m_frmWait.pbInterm.Value = Me.m_iPos
        Me.m_iPos += Me.m_iDirection
        If Me.m_iPos = 100 Then Me.m_iDirection = -1
        If Me.m_iPos = 0 Then Me.m_iDirection = 1

    End Sub

    Private Function lTailleTotalFichiers&(NbrFiles%, sCheminDossierRef$, asFichiers$())

        Dim iNumFichier%
        ReDim Me.m_lLongFichier(NbrFiles - 1)
        ' Poids total en octets de tous les fichiers � traiter
        Dim lngTot& = 0
        For iNumFichier = 0 To NbrFiles - 1
            Dim fi As New IO.FileInfo(sCheminDossierRef & "\" & asFichiers(iNumFichier))
            If Not fi.Exists Then Continue For ' 17/08/2009
            Me.m_lLongFichier(iNumFichier) = fi.Length
            lngTot += Me.m_lLongFichier(iNumFichier)
        Next
        lTailleTotalFichiers = lngTot

    End Function

End Class

#End Region