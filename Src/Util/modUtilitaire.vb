
' Fichier modUtilitaire.vb : Module de fonctions utilitaires diverses
' ------------------------

Imports System.Text

Module Utilitaire


#Region "Convertions"

    Public Function iConv%(sVal$, Optional iValDef% = 0)

        If sVal.Length = 0 Then iConv = iValDef : Exit Function
        Try
            iConv = CInt(sVal)
            'If Not Integer.TryParse(sVal, iConv) Then iConv = iValDef
        Catch
            iConv = iValDef
        End Try

    End Function

#End Region

#Region "Divers"

    Public Function bAppliDejaOuverte(bMemeExe As Boolean) As Boolean

        ' Détecter si l'application est déja lancée :
        ' - depuis n'importe quelle copie de l'exécutable, ou bien seulement
        ' - depuis le même emplacement du fichier exécutable sur le disque dur

        Dim sExeProcessAct$ = Diagnostics.Process.GetCurrentProcess.MainModule.ModuleName
        Dim sNomProcessAct$ = IO.Path.GetFileNameWithoutExtension(sExeProcessAct)
        If Not bMemeExe Then
            ' Détecter si l'application est déja lancée depuis n'importe quel exe
            If Process.GetProcessesByName(sNomProcessAct).Length > 1 Then Return True
            Return False
        End If

        ' Détecter si l'application est déja lancée depuis le même exe
        Dim sCheminProcessAct$ = Diagnostics.Process.GetCurrentProcess.MainModule.FileName
        Dim aProcessAct As Diagnostics.Process() = Process.GetProcessesByName(sNomProcessAct)
        Dim processAct As Diagnostics.Process
        Dim iNbApplis% = 0
        For Each processAct In aProcessAct
            Dim sCheminExe$ = processAct.MainModule.FileName
            If sCheminExe = sCheminProcessAct Then iNbApplis += 1
        Next
        If iNbApplis > 1 Then Return True
        Return False

    End Function

    Public Sub Sablier(Optional ByRef bDesactiver As Boolean = False)

        If bDesactiver Then
            'm_curseur = Cursors.Default
            Cursor.Current = Cursors.Default
        Else
            'm_curseur = Cursors.WaitCursor
            Cursor.Current = Cursors.WaitCursor
        End If

        'm_oFrmMousePointer.Cursor = m_curseur ' Curseur de la feuille
        'Cursor.Current = m_curseur ' Curseur de l'application
        'Exit Sub

        ' Curseur de l'application : il est réinitialisé à chaque Application.DoEvents
        '  ou bien lorsque l'application ne fait rien
        '  du coup, il faut insister grave pour conserver le contrôle du curseur tout en 
        '  voulant afficher des messages de progression et vérifier les interruptions...
        'Dim ctrl As Control
        'For Each ctrl In m_oFrmMousePointer.Controls
        '    ctrl.Cursor = m_curseur ' Curseur de chaque contrôle de la feuille
        'Next ctrl

    End Sub

    Public Sub AfficherMsgErreur(ByRef Erreur As Microsoft.VisualBasic.ErrObject,
        Optional sTitreFct$ = "", Optional sInfo$ = "", Optional sDetailMsgErr$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If Erreur.Number > 0 Then
            sMsg &= vbCrLf & "Err n°" & Erreur.Number.ToString & " :"
            sMsg &= vbCrLf & Erreur.Description
        End If
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        MsgBox(sMsg, MsgBoxStyle.Critical, m_sTitreMsg)

    End Sub

    Public Sub AfficherMsgErreur2(ByRef Ex As Exception,
        Optional sTitreFct$ = "", Optional sInfo$ = "",
        Optional sDetailMsgErr$ = "",
        Optional bCopierMsgPressePapier As Boolean = True,
        Optional ByRef sMsgErrFinal$ = "")

        If Not Cursor.Current.Equals(Cursors.Default) Then _
            Cursor.Current = Cursors.Default
        Dim sMsg$ = ""
        If sTitreFct <> "" Then sMsg = "Fonction : " & sTitreFct
        If sInfo <> "" Then sMsg &= vbCrLf & sInfo
        If sDetailMsgErr <> "" Then sMsg &= vbCrLf & sDetailMsgErr
        If Ex.Message <> "" Then
            sMsg &= vbCrLf & Ex.Message.Trim
            If Not IsNothing(Ex.InnerException) Then _
            sMsg &= vbCrLf & Ex.InnerException.Message
        End If
        If bCopierMsgPressePapier Then CopierPressePapier(sMsg)
        sMsgErrFinal = sMsg
        MsgBox(sMsg, MsgBoxStyle.Critical)

    End Sub

    Public Sub CopierPressePapier(sInfo$)

        ' Copier des informations dans le presse-papier de Windows
        ' (elles resteront jusqu'à ce que l'application soit fermée)

        Try
            Dim dataObj As New DataObject
            dataObj.SetData(DataFormats.Text, sInfo)
            Clipboard.SetDataObject(dataObj)
        Catch ex As Exception
            ' Le presse-papier peut être indisponible
            AfficherMsgErreur2(ex, "CopierPressePapier", bCopierMsgPressePapier:=False)
        End Try

    End Sub

    Public Sub TraiterMsgSysteme_DoEvents()
        Try
            Application.DoEvents() ' Peut planter avec OWC : Try Catch nécessaire
        Catch
        End Try
    End Sub

    Public Sub Attendre(iDelaiMilliSec%)
        Application.DoEvents()
        Threading.Thread.Sleep(iDelaiMilliSec)
        Application.DoEvents()
    End Sub

    Public Function bChoisirFichier(ByRef sCheminFichier$, sFiltre$, sExtDef$,
        sTitre$, Optional sInitDir$ = "",
        Optional bDoitExister As Boolean = True,
        Optional bMultiselect As Boolean = False) As Boolean

        ' Afficher une boite de dialogue pour choisir un fichier
        ' Exemple de filtre : "|Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
        ' On peut indiquer le dossier initial via InitDir, ou bien via le chemin du fichier

        Static bInit As Boolean = False
        Dim ofd As New OpenFileDialog
        With ofd
            If Not bInit Then
                bInit = True
                If sInitDir.Length = 0 Then
                    If sCheminFichier.Length = 0 Then
                        .InitialDirectory = Application.StartupPath
                    Else
                        .InitialDirectory = IO.Path.GetDirectoryName(sCheminFichier)
                    End If
                Else
                    .InitialDirectory = sInitDir
                End If
            End If
            If Not String.IsNullOrEmpty(sCheminFichier) Then .FileName = sCheminFichier
            .CheckFileExists = bDoitExister ' 14/10/2007
            .DefaultExt = sExtDef
            .Filter = sFiltre
            .Multiselect = bMultiselect
            .Title = sTitre
            .ShowDialog()
            If .FileName <> "" Then sCheminFichier = .FileName : Return True
            Return False
        End With

    End Function

    Public Function LireEncodageVB6(sChemin$) As Encoding

        ' Déterminer l'encodage du fichier en analysant ses 1ers octets 
        ' (Byte Order Mark, ou BOM). Par défaut l'encodage sera ASCII si on ne trouve pas

        ' Lecture de la BOM
        Dim bom As Byte() = New Byte(3) {}
        Using file As IO.FileStream = New IO.FileStream(sChemin, IO.FileMode.Open,
        IO.FileAccess.Read, IO.FileShare.ReadWrite) ' 05/01/2018 Need only read-only access, not write access
            file.Read(bom, 0, 4)
        End Using

        ' Analyse de la BOM
        If bom(0) = &H2B AndAlso bom(1) = &H2F AndAlso bom(2) = &H76 Then
            Return Encoding.UTF7
        End If
        If bom(0) = &HEF AndAlso bom(1) = &HBB AndAlso bom(2) = &HBF Then
            Return Encoding.UTF8
        End If

        If bom(0) = &H22 AndAlso bom(1) = &H43 AndAlso bom(2) = &H6F AndAlso bom(3) = &H75 Then
            Return Encoding.UTF8
        End If

        If bom(0) = 50 AndAlso bom(1) = 48 AndAlso bom(2) = 49 AndAlso bom(3) = 54 Then
            Return Encoding.UTF8
        End If

        If bom(0) = 34 AndAlso bom(1) = 105 AndAlso bom(2) = 100 AndAlso bom(3) = 34 Then
            Return Encoding.UTF8
        End If

        If bom(0) = &HFF AndAlso bom(1) = &HFE Then
            Return Encoding.Unicode
        End If

        ' UTF-16LE
        If bom(0) = &HFE AndAlso bom(1) = &HFF Then
            Return Encoding.BigEndianUnicode
        End If

        ' UTF-16BE
        If bom(0) = 0 AndAlso bom(1) = 0 AndAlso bom(2) = &HFE AndAlso bom(3) = &HFF Then
            Return Encoding.UTF32
        End If

        Return Encoding.ASCII

    End Function

#End Region

End Module