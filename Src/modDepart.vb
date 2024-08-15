
' Fichier modDepart.vb
' --------------------

Module modDepart

#If DEBUG Then
    Public Const bDebug As Boolean = True
    Public Const bRelease As Boolean = False
    Public Const bTrapErr As Boolean = False
#Else
    Public Const bDebug As Boolean = False
    Public Const bRelease As Boolean = True
    Public Const bTrapErr As Boolean = True
#End If

    'Public Const sTitreMsg$ = "VB2Html"
    Public Const sDateVersionAppli$ = "15/08/2024"
    Public ReadOnly sNomAppli$ = My.Application.Info.Title
    Public ReadOnly m_sTitreMsg$ = sNomAppli
    Public ReadOnly sVersionAppli$ =
        My.Application.Info.Version.Major & "." &
        My.Application.Info.Version.Minor &
        My.Application.Info.Version.Build

    Public Sub Main()

        If bDebug Then Depart() : Exit Sub

        Try
            Depart()
        Catch ex As Exception
            AfficherMsgErreur2(ex, "Main " & m_sTitreMsg)
        End Try

    End Sub

    Private Sub Depart()

        If bAppliDejaOuverte(bMemeExe:=True) Then Exit Sub

        ' Extraire les options passées en argument de la ligne de commande
        ' Ne fonctionne pas avec des chemins contenant des espaces, même entre guillemets
        'Dim asArgs$() = Environment.GetCommandLineArgs()
        Dim sArg0$ = Microsoft.VisualBasic.Interaction.Command

        Const sInfoArg$ = "Argument : Chemin du projet" & vbLf &
            "Pensez à mettre entre guillemets les chemins contenant des espaces"

        Dim sCheminPrj$ = ""
        If sArg0 <> "" Then
            Dim asArgs$() = asArgLigneCmd(sArg0)
            Dim iNbArgs% = UBound(asArgs) + 1
            If iNbArgs = 1 Then
                ' S'il n'y a qu'un argument, alors c'est le chemin du projet
                sCheminPrj = asArgs(0)
            Else
                Dim iNumArg1%
                For iNumArg1 = 0 To (iNbArgs - 1) \ 2
                    Dim sCle$ = asArgs(iNumArg1 * 2)
                    If iNumArg1 * 2 + 1 > UBound(asArgs) Then
                        MsgBox("Erreur : Nombre impair d'arguments !" & vbLf &
                            sInfoArg, MsgBoxStyle.Critical, m_sTitreMsg)
                        Exit Sub
                    End If
                    Dim sVal$ = asArgs(iNumArg1 * 2 + 1)
                    Select Case sCle.ToLower
                        'Case "x".ToLower
                        '    sX = sVal
                        'Case "y".ToLower
                        '    sY = sVal
                        Case Else
                            MsgBox("Erreur : Argument non reconnu : [" & sCle & "]" & vbLf &
                            sInfoArg, MsgBoxStyle.Critical, m_sTitreMsg)
                            Exit Sub
                    End Select
                Next iNumArg1
            End If
        End If

        Dim oFrm As New frmVB2Html
        oFrm.m_sCheminPrj = sCheminPrj
        ' ShowDialog ne fonctionne pas si aucune session n'est ouverte
        'oFrm.ShowDialog()
        Application.Run(oFrm)

        ' Si on n'utilise pas l'application framework (case à cocher 
        '  "Activer l'infrastructure de l'application" : exige que le projet démarre 
        '  sur la form principale), il suffit d'appeler Save !
        My.Settings.Save()

        ' Pour que la sauvegarde des paramètres marche depuis l'IDE Visual Studio .Net
        '  mais en DotNet2, le fichier est ici :
        '\Documents and Settings\<utilisateur>\Local Settings\Application Data\
        ' VB2Html.exe_Url_xxx...xxx\2.0.2.xxxxx\user.config
        'If bDebug Then
        '    Dim sSrc$ = Application.StartupPath & "\" & _
        '        My.Application.Info.AssemblyName & ".exe.config"
        '    Dim sDest$ = Application.StartupPath & "\app.config"
        '    bCopierFichier(sSrc, sDest, bPromptErr:=True)
        'End If

    End Sub

End Module