
' Fichier modConvertor : Gestion de la conversion d'un projet en Html
' --------------------

Imports System.Text ' StringBuilder

Module modConvertor

    Public Const bFrmPatienter As Boolean = True

    Private Const SEPARATORS As String = " ,:;^+-*/&()=<>." & vbLf
    Private Const CKW_SEPARATORS As String = "."

    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" _
        (lpPath As String) As Integer

    Private m_iMaxLen As Integer 'Short
    Private m_bCompDec As Boolean
    Private m_sActKWList As String
    Private m_bMakeTDM As Boolean

    Public Sub TransformToHTML(ProjectTitle$,
        ProjectVersion$, KWList$,
        ByRef FilesList$(), FilesNum%,
        ProjFile$, HTMLFile$, sContenuModeleHtml$,
        TroncLen%, vCompDec As Boolean, MultiFiles As Boolean,
        MakeFrames As Boolean, vMakeTDM As Boolean,
        bVB2Txt As Boolean, patienter As clsPatienter)

        Dim MajorIndex, TransFiles, FreeCanal, FreeCanal2 As Integer
        Dim PageStruct2, TDM, HTMLBaseDir, ProjBaseDir, PageStruct,
            TempTDM, MenuFile As String ' CopyFile, 
        TDM = ""

        Dim lecture As New clsUtilFichierVB6Lecture
        Dim ecriture As New clsUtilFichierVB6Ecriture

        If (FilesNum > 0 Or IsProjectFile(ProjFile) = False) And HTMLFile <> "" Then
            ProjBaseDir = Left(ProjFile, InStrRev(ProjFile, "\"))
            HTMLBaseDir = Left(HTMLFile, InStrRev(HTMLFile, "\"))

            m_iMaxLen = TroncLen
            m_bCompDec = vCompDec
            m_sActKWList = KWList
            m_bMakeTDM = vMakeTDM

            'Dim encodage As Encoding = Encoding.GetEncoding(iCodePageWindowsLatin1252)
            FreeCanal = ecriture.iFreeCanal(Encoding.UTF8)

            PageStruct2 = sContenuModeleHtml

            If IsProjectFile(ProjFile) Then

                PageStruct = Replace(PageStruct2, "%%<TITLE>%%",
                    ProjectTitle & ProjectVersion & " converted with " & sNomAppli)
                PageStruct = Replace(PageStruct, "%%<CODE>%%",
                    "<div class=""a"">" & ProjectTitle & ProjectVersion & "</div>" & vbLf & vbLf &
                    CStr(IIf(m_bMakeTDM,
                        "<span class=""b"">Table des procédures</span>" & vbLf & vbLf,
                        "")) &
                    "%%<CODE>%%")

                If Not MultiFiles And Not bVB2Txt Then
                    Dim sChemin$ = HTMLFile & CStr(IIf(m_bMakeTDM, ".tmp", ""))
                    ecriture.VB6FileOpen(FreeCanal, sChemin, OpenMode.Output)
                    If m_bMakeTDM = False Then ecriture.VB6Print(FreeCanal,
                        Left(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") - 1))
                End If

                For TransFiles = 0 To FilesNum - 1

                    If bFrmPatienter Then
                        patienter.MettreAJour(TransFiles)
                        TraiterMsgSysteme_DoEvents()
                        If patienter.bAnnuler Then Exit For
                    End If

                    If bVB2Txt Then
                        Dim sChemin$ = HTMLBaseDir & FilesList(TransFiles)
                        Dim sDossierTxt$ = IO.Path.GetDirectoryName(HTMLBaseDir) & "\VB2Txt"
                        ' Au premier fichier supprimer le dossier tmp VB2Txt
                        If TransFiles = 0 Then bSupprimerDossier(sDossierTxt, bPromptErr:=True)
                        If bVerifierCreerDossier(sDossierTxt) Then
                            Dim sCheminRelatif0$ = sCheminRelatif(sChemin, HTMLBaseDir)
                            sCheminRelatif0 = sEnleverSlashInitial(sCheminRelatif0)
                            sCheminRelatif0 = sCheminRelatif0.Replace("\", "_")
                            Dim sCheminTxt0$ = sDossierTxt & "\" & sCheminRelatif0
                            Dim sCheminTxt$ = IO.Path.GetDirectoryName(sCheminTxt0) & "\" &
                                IO.Path.GetFileNameWithoutExtension(sCheminTxt0) &
                                IO.Path.GetExtension(sChemin).Replace(".", "_") & ".txt"
                            If IsFileDotNet(sChemin) Then
                                ' Si c'est une fichier src DotNet, alors il suffit de le copier
                                If bCopierFichier(sChemin, sCheminTxt) Then
                                    'Debug.WriteLine("Fichier actuel : " & sChemin & " -> " & sCheminTxt)
                                End If
                            Else
                                ' Si c'est un fichier VB6, alors il faut enlever jusqu'au dernier
                                '  attribut Attribute VB_xxx = yyy 
                                Dim s$ = sLireFichier(sChemin)
                                '  pb : cela apparait dans le code modConvertor.bas !!!
                                'Dim iPos% = s.LastIndexOf("Attribute VB_")
                                Dim iIndex% = 0
                                ' Trouver le premier attribut toujours présent
                                Dim iPos% = s.IndexOf("Attribute VB_Name = ")
                                If iPos >= 0 Then
                                    Do
                                        ' Trouver la fin de la ligne
                                        Dim iPosFin% = s.IndexOf(vbLf, iPos)
                                        If iPosFin <= 0 Then Exit Do
                                        iIndex = iPosFin + 1
                                        ' Trouver la fin de la ligne suivante
                                        Dim iPosFin2% = s.IndexOf(vbLf, iIndex)
                                        ' Rechercher tous les attributs
                                        iPos = s.IndexOf("Attribute VB_", iIndex)
                                        ' Si iPos est > à la fin de la ligne suiv. 
                                        '  alors on a dépassé la fin des attributs, c'est du code !
                                        If iPos > iPosFin2 Then Exit Do
                                    Loop While iPos >= 0
                                    If bEcrireFichier(sCheminTxt, s.Substring(iIndex)) Then
                                        ' Conserver la date du fichier d'origine
                                        Dim fiSrc As New IO.FileInfo(sChemin)
                                        Dim fiDest As New IO.FileInfo(sCheminTxt)
                                        fiDest.LastWriteTime = fiSrc.LastWriteTime
                                        fiSrc = Nothing
                                        fiDest = Nothing
                                    End If
                                End If
                            End If
                        End If
                        GoTo FichierSuivant
                    End If

                    If MultiFiles Then
                        MakeSureDirectoryPathExists(HTMLBaseDir & FilesList(TransFiles))
                        Dim sFichier$ = HTMLBaseDir & FilesList(TransFiles) & ".html"
                        ecriture.VB6FileOpen(FreeCanal, sFichier, OpenMode.Output)
                        Dim sLigne$ = Replace(Left(PageStruct2,
                            InStr(1, PageStruct2, "%%<CODE>%%") - 1),
                            "%%<TITLE>%%", FilesList(TransFiles) & " converted by " & sNomAppli)
                        ecriture.VB6Print(FreeCanal, TAB, sLigne)
                    End If

                    Dim bIncludeTDM As Boolean = m_bMakeTDM And MakeFrames = False And MultiFiles = True
                    TempTDM = TransformFileToHTML(ProjBaseDir & FilesList(TransFiles), FreeCanal,
                        bIncludeTDM, MajorIndex,
                        CStr(IIf(MakeFrames, "code", "")), patienter, ecriture)

                    If MultiFiles Then
                        TDM = TDM & Replace(Replace(TempTDM, "#",
                            Replace(FilesList(TransFiles), "\", "/") & ".html" & "#"),
                            "#" & MajorIndex & Chr(34), Chr(34)) & vbLf
                    Else
                        TDM = TDM & TempTDM & vbLf
                    End If
                    'Debug.WriteLine(TDM)

                    If MultiFiles Then
                        ecriture.VB6Print(FreeCanal, Mid(PageStruct2,
                            InStr(1, PageStruct2, "%%<CODE>%%") + Len("%%<CODE>%%")))
                        ecriture.VB6FileClose(FreeCanal)
                    End If

FichierSuivant:
                Next TransFiles

                If bVB2Txt Then Exit Sub

                If MultiFiles = False Then
                    If m_bMakeTDM = False Then ecriture.VB6Print(FreeCanal, TAB,
                        Mid(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") + Len("%%<CODE>%%")))

                    ecriture.VB6FileClose(FreeCanal)

                    If m_bMakeTDM Then

                        FreeCanal2 = ecriture.iFreeCanal(Encoding.UTF8)
                        ecriture.VB6FileOpen(FreeCanal2, HTMLFile, OpenMode.Output)
                        ecriture.VB6Print(FreeCanal2, Left(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") - 1))
                        ecriture.VB6Print(FreeCanal2, TDM & vbLf & vbLf)

                        Dim sChemin$ = HTMLFile & ".tmp"
                        ''FileOpen(FreeCanal, sChemin, OpenMode.Binary, OpenAccess.Read)
                        'Dim iCanalLect% = lecture.iFreeCanal() ' 08/09/2018
                        'lecture.VB6FileOpen(iCanalLect, sChemin, OpenMode.Binary, OpenAccess.Read)
                        'Do Until lecture.bEOF(iCanalLect)
                        '    Dim CopyFile$ = New String(Chr(0), 512)
                        '    lecture.VB6FileGet(iCanalLect, CopyFile)
                        '    ecriture.VB6Print(FreeCanal2, CopyFile)
                        'Loop
                        'lecture.VB6FileClose(iCanalLect)
                        ' 09/09/2018 Ne pas trancher un fichier unicode par bloc de 512 octets
                        '  ça provoque de petites coupures
                        Dim sContenu$ = sLireFichier(sChemin, bLectureSeule:=True, bUnicodeUTF8:=True)
                        ecriture.VB6Print(FreeCanal2, sContenu)

                        ecriture.VB6Print(FreeCanal2, Mid(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") + Len("%%<CODE>%%")))
                        ecriture.VB6FileClose(FreeCanal2)

                        Kill(HTMLFile & ".tmp")
                    End If

                ElseIf MultiFiles Then

                    ecriture.VB6FileOpen(FreeCanal, HTMLFile, OpenMode.Output)
                    FreeCanal2 = ecriture.iFreeCanal(Encoding.UTF8)

                    If MakeFrames Then
                        MenuFile = Mid(HTMLFile, InStrRev(HTMLFile, "\") + 1)
                        MenuFile = Left(MenuFile, InStrRev(MenuFile, ".") - 1) & "_main.html"

                        ecriture.VB6Print(FreeCanal, Left(PageStruct, InStr(1, PageStruct, "<body>") - 1))
                        ecriture.VB6Print(FreeCanal, "<frameset cols=""300,*"">" & vbLf & "<frame name=""tdm"" src=""" & MenuFile & """>" & vbLf & "<frame name=""code"" src=""" & FilesList(0) & ".html"">" & vbLf & "</frameset>")
                        ecriture.VB6Print(FreeCanal, Mid(PageStruct, InStr(1, PageStruct, "</body>") + Len("</body>")))

                        ecriture.VB6FileOpen(FreeCanal2, HTMLBaseDir & MenuFile, OpenMode.Output)
                        ecriture.VB6Print(FreeCanal2, Replace(Left(PageStruct2, InStr(1, PageStruct2, "%%<CODE>%%") - 1), "<title>%%<TITLE>%%</title>", ""))
                        ecriture.VB6Print(FreeCanal2, TDM)
                        ecriture.VB6Print(FreeCanal2, Mid(PageStruct2, InStr(1, PageStruct2, "%%<CODE>%%") + Len("%%<CODE>%%")))
                        ecriture.VB6FileClose(FreeCanal2)
                    Else
                        ecriture.VB6Print(FreeCanal, Left(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") - 1))
                        ecriture.VB6Print(FreeCanal, TDM)
                        ecriture.VB6Print(FreeCanal, Mid(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") + Len("%%<CODE>%%")))
                    End If
                    ecriture.VB6FileClose(FreeCanal)
                End If

            Else ' If Not IsProjectFile(ProjFile) Then

                PageStruct = Replace(PageStruct2, "%%<TITLE>%%", Mid(ProjFile, InStrRev(ProjFile, "\") + 1) & " converted with " & sNomAppli)

                Dim sFichier$ = CStr(
                    IIf(MultiFiles,
                        HTMLBaseDir & Mid(ProjFile, InStrRev(ProjFile, "\") + 1) & ".html",
                        HTMLFile))
                ecriture.VB6FileOpen(FreeCanal, sFichier, OpenMode.Output)
                ecriture.VB6Print(FreeCanal, Left(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") - 1))

                TempTDM = TransformFileToHTML(ProjFile, FreeCanal, MakeFrames = False, 0,
                    CStr(IIf(MakeFrames, "code", "")), patienter, ecriture)
                If MultiFiles Then TempTDM = Replace(Replace(TempTDM, "#", Mid(ProjFile, InStrRev(ProjFile, "\") + 1) & ".html" & "#"), "#1" & Chr(34), Chr(34)) & vbLf

                ecriture.VB6Print(FreeCanal, Mid(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") + Len("%%<CODE>%%")))
                ecriture.VB6FileClose(FreeCanal)

                If MultiFiles Then
                    PageStruct = Replace(PageStruct, "%%<CODE>%%", "<div class=""a"">" &
                        Mid(ProjFile, InStrRev(ProjFile, "\") + 1) & "</div>" & vbLf & vbLf &
                        CStr(IIf(m_bMakeTDM,
                                 "<span class=""b"">Table des procédures</span>" & vbLf & vbLf,
                                 "")) & "%%<CODE>%%")

                    ecriture.VB6FileOpen(FreeCanal, HTMLFile, OpenMode.Output)
                    If MakeFrames Then
                        MenuFile = Mid(HTMLFile, InStrRev(HTMLFile, "\") + 1)
                        MenuFile = Left(MenuFile, InStrRev(MenuFile, ".") - 1) & "_main.html"

                        ecriture.VB6Print(FreeCanal, Left(PageStruct, InStr(1, PageStruct, "<body>") - 1))
                        ecriture.VB6Print(FreeCanal, "<frameset cols=""300,*"">" & vbLf & "<frame name=""tdm"" src=""" & MenuFile & """>" & vbLf & "<frame name=""code"" src=""" & Mid(ProjFile, InStrRev(ProjFile, "\") + 1) & ".html"">" & vbLf & "</frameset>")
                        ecriture.VB6Print(FreeCanal, Mid(PageStruct, InStr(1, PageStruct, "</body>") + Len("</body>")))

                        FreeCanal2 = ecriture.iFreeCanal(Encoding.UTF8)

                        ecriture.VB6FileOpen(FreeCanal2, HTMLBaseDir & MenuFile, OpenMode.Output)
                        ecriture.VB6Print(FreeCanal2, Replace(Left(PageStruct2, InStr(1, PageStruct2, "%%<CODE>%%") - 1), "<title>%%<TITLE>%%</title>", ""))
                        ecriture.VB6Print(FreeCanal2, TempTDM)
                        ecriture.VB6Print(FreeCanal2, Mid(PageStruct2, InStr(1, PageStruct2, "%%<CODE>%%") + Len("%%<CODE>%%")))
                        ecriture.VB6FileClose(FreeCanal2)
                    Else
                        ecriture.VB6Print(FreeCanal, Left(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") - 1))
                        ecriture.VB6Print(FreeCanal, TempTDM)
                        ecriture.VB6Print(FreeCanal, Mid(PageStruct, InStr(1, PageStruct, "%%<CODE>%%") + Len("%%<CODE>%%")))
                    End If
                    ecriture.VB6FileClose(FreeCanal)
                End If
            End If

        Else
            If FilesNum > 0 Then _
            MsgBox("Il manque certaines informations pour la transcription du code VB en HTML",
                MsgBoxStyle.Information, "Informations manquantes")
        End If

    End Sub

    Private Function TransformFileToHTML$(FileName$,
        WriteCanal%, IncludeTDM As Boolean,
        ByRef MajorIndex%, PageTarget$,
        patienter As clsPatienter,
        ecriture As clsUtilFichierVB6Ecriture)

        MajorIndex = MajorIndex + 1

        Dim LineLen, ReorgTDM2, ActType, MinorIndex, PrevString, PrevType, ReorgTDM, MinLink, TabLen As Integer 'Short
        Dim AddToTDM, FileExt, FileData, FileChr, PrevInst, ObjectTitle, PrevSep As String

        'Dim TempInst$ = ""
        Dim sbTempInst As New StringBuilder

        Dim sbActLine As New StringBuilder

        Dim sMemPrevInst$ = ""
        PrevSep = ""
        PrevInst = ""
        AddToTDM = ""
        ObjectTitle = ""

        Dim NbInsts, AnaFile, PrevInstNum As Integer
        Dim StringCut, IsComment, IsString, PrevComment, NoAddSpaces As Boolean
        Dim ProgInsts() As String = Nothing
        Dim TDMLinks() As String = Nothing

        FileExt = GetFileType(FileName)

        ' L'encodage peut changer : utiliser la fct DotNet, au moins ici :
        FileData = sLireFichier(FileName)

        FileData = Replace(FileData, vbCr, "")
        FileData = Replace(FileData, vbTab, "    ") ' 15/09/2007

        If FileExt <> ".vb" And FileExt <> ".txt" Then
            ObjectTitle = Replace(GetBoundedString(FileData, "Attribute VB_Name = ", vbLf), Chr(34), "")

            If FileExt = ".bas" Then
                Dim iPos% = InStr(1, Trim(FileData), "Attribute VB_Name")
                If iPos = 0 Then
                    ' 15/09/2007
                    ' Il peut s'agir de n'importe quelle documentation ajoutée au projet
                    TransformFileToHTML = ""
                    Exit Function
                End If
                FileData = Mid(Trim(FileData), InStr(iPos, FileData, vbLf) + 1)
            Else
                Dim iPos% = InStr(1, Trim(FileData), "Attribute VB_Exposed")
                If iPos = 0 Then TransformFileToHTML = "" : Exit Function ' 15/09/2007
                FileData = Mid(Trim(FileData), InStr(iPos, FileData, vbLf) + 1)
            End If
        End If

        TransformFileToHTML = ""
        TransformFileToHTML = TransformFileToHTML & MajorIndex & " - <a href=""#" & MajorIndex & Chr(34) &
            CStr(IIf(PageTarget <> "", " target=""" & PageTarget & Chr(34), "")) & ">" &
            CStr(IIf(ObjectTitle <> "", ObjectTitle & " (" &
                Mid(FileName, InStrRev(FileName, "\") + 1) & ")",
                Mid(FileName, InStrRev(FileName, "\") + 1))) & "</a>" & vbLf

        If Right(FileData, 1) <> vbLf Then FileData = FileData & vbLf

        Dim bProcedure As Boolean = False
        Dim iTailleFichier% = Len(FileData)
        For AnaFile = 1 To iTailleFichier

            If bFrmPatienter And ((AnaFile Mod 300 = 0) Or AnaFile = iTailleFichier) Then
                patienter.MettreAJour_Interm(AnaFile - 1)
                TraiterMsgSysteme_DoEvents()
                If patienter.bAnnuler Then Exit For
            End If

            FileChr = Mid(FileData, AnaFile, 1)

            If FileChr = vbLf Then

                LineLen = (-1)

                If Right(Trim(sbActLine.ToString), 2) <> " _" Then
                    sbActLine.Length = 0
                    TabLen = 0
                Else
                    sbActLine.Length -= 1
                    NoAddSpaces = True
                End If

            ElseIf FileChr = ":" Or FileChr = ";" Then

                sbActLine.Length = 0

            Else

                If (NoAddSpaces = False And sbActLine.Length > 0) Or FileChr <> " " Then

                    NoAddSpaces = False
                    sbActLine.Append(ChrToHTML(FileChr))

                ElseIf sbActLine.Length = 0 Then
                    TabLen = TabLen + 1
                End If

            End If

            If m_iMaxLen > 0 Then LineLen = LineLen + 1

            If StringCut Then
                ReDim Preserve ProgInsts(NbInsts)

                ProgInsts(NbInsts) = " &amp; _" & vbLf & Space(TabLen + 4) &
                    "<span class=""d"">" & Chr(34)
                ProgInsts(PrevInstNum) = ProgInsts(PrevInstNum) & "</span>"
                PrevType = 2
                LineLen = TabLen + 6
                NbInsts = NbInsts + 1

                StringCut = False
                IsString = True
            End If

            If LineLen > m_iMaxLen And m_iMaxLen > 0 And IsComment = False Then

                If IsString = False Then

                    ReDim Preserve ProgInsts(NbInsts)
                    ProgInsts(NbInsts) = " _" & vbLf & Space(TabLen + 4)
                    NbInsts = NbInsts + 1

                Else

                    IsString = False
                    PrevString = 2
                    StringCut = True
                    'TempInst = TempInst & Chr(34)
                    sbTempInst.Append(Chr(34))
                    AnaFile = AnaFile - 1

                End If

                LineLen = 0

            End If

            If FileChr = "'" And IsString = False Then
                IsComment = True
            ElseIf FileChr = vbLf And IsComment Then
                IsComment = False
                PrevComment = True
            ElseIf FileChr = Chr(34) And IsComment = False Then
                IsString = Not IsString
                PrevString = 1
            End If

            If (StringCut Or InStr(1, SEPARATORS, FileChr) > 0) And
                IsComment = False And IsString = False Then

                Dim sTempInst$ = sbTempInst.ToString.ToLower
                'Debug.WriteLine(sTempInst)
                'If sTempInst = "end" Then
                '    Debug.WriteLine("!")
                'End If

                If AddToTDM <> "" Then

                    If (LCase(PrevInst) = "property" And
                       (sTempInst = "get" Or sTempInst = "set" Or sTempInst = "let")) _
                       Or (m_bCompDec And sbActLine.Length > 0) Then

                        AddToTDM = sbActLine.ToString
                        'Debug.WriteLine(AddToTDM)

                    Else

                        ReDim Preserve TDMLinks(MinorIndex)
                        TDMLinks(MinorIndex) = AddToTDM & sbTempInst.ToString
                        AddToTDM = ""

                    End If
                End If

                If sbTempInst.Length > 0 Then

                    ReDim Preserve ProgInsts(NbInsts)

                    If PrevComment Then

                        ProgInsts(NbInsts) = CStr(IIf(PrevType <> 3,
                            "<span class=""e"">", "")) & sbTempInst.ToString
                        ActType = 3

                    ElseIf PrevString = 2 Then

                        ProgInsts(NbInsts) = CStr(IIf(PrevType <> 2,
                            "<span class=""d"">", "")) & sbTempInst.ToString
                        ActType = 2

                    ElseIf IsNumeric(sbTempInst.ToString) Then

                        ProgInsts(NbInsts) = CStr(IIf(PrevType <> 4,
                            "<span class=""f"">", "")) & sbTempInst.ToString
                        ActType = 4

                    ElseIf InStr(1, m_sActKWList, "|" & sTempInst & "|") > 0 And
                           InStr(1, CKW_SEPARATORS, PrevSep) = 0 Then

                        ' 11/05/2014 Ignorer aussi toutes les fonctions lambdas du Linq
                        '  mais il peut y avoir plusieurs fct lambdas consécutives
                        '  une meilleure solution consiste a détecter si on n'est pas déjà dans une
                        '  procédure
                        Dim lstExcep As New List(Of String) From {"end", "declare", "exit"} ', _
                        '"where", "sum", "max", "groupby", "average", "selectmany", _
                        '"aggregate", "takewhile", "skipwhile", "select", "aggregate", _
                        '"any", "todictionary", "orderby", "thenbydescending", "orderbydescending"}
                        If m_bMakeTDM AndAlso
                           (sTempInst = "sub" OrElse sTempInst = "function" OrElse sTempInst = "property") Then

                            If Not bProcedure AndAlso
                               Not lstExcep.Contains(PrevInst.ToLower) AndAlso
                               Not lstExcep.Contains(sMemPrevInst.ToLower) Then
                                MinorIndex += 1
                                AddToTDM = sbActLine.ToString
Retester:                   ' 11/05/2014 Certains attributs ne passent pas
                                If AddToTDM.StartsWith("=") OrElse AddToTDM.StartsWith("&lt;") OrElse
                                   AddToTDM.Contains("&gt;") Then
                                    Const sIndic$ = "&gt;"
                                    Dim iPos% = AddToTDM.IndexOf(sIndic)
                                    If iPos > -1 Then
                                        Dim sCoupe$ = AddToTDM.Substring(iPos + sIndic.Length)
                                        AddToTDM = sCoupe.Trim
                                        GoTo Retester
                                    End If
                                End If
                                'Debug.WriteLine(AddToTDM & " : " & PrevInst & ", " & _
                                '    sTempInst & ", " & sMemPrevInst)
                                'Debug.WriteLine(AddToTDM & " : proc.=" & bProcedure)
                                bProcedure = True
                            End If

                            Dim PrevInstLC$ = PrevInst.ToLower
                            'Dim sInst2$ = sMemPrevInst.ToLower
                            If PrevInstLC = "end" AndAlso (sTempInst = "sub" OrElse sTempInst = "function" OrElse
                                                      sTempInst = "property") Then
                                'Debug.WriteLine(PrevInst & ", " & sTempInst & ", " & sMemPrevInst)
                                bProcedure = False
                                PrevInst = "" ' 15/06/2014 Ignorer la procédure précédente
                            End If

                        End If

                        ActType = 1
                        Dim sTxt$ = CStr(IIf(AddToTDM <> "",
                            "<a name=""" & MajorIndex & MinorIndex & """></a>",
                            "")) & CStr(IIf(PrevType <> 1, "<span class=""c"">", "")) & sbTempInst.ToString
                        ProgInsts(NbInsts) = sTxt

                    Else

                        ProgInsts(NbInsts) = sbTempInst.ToString

                        ActType = 0

                    End If

                    If PrevType <> ActType And PrevType > 0 Then
                        ProgInsts(PrevInstNum) = ProgInsts(PrevInstNum) & "</span>"
                    End If

                    PrevType = ActType

                    sMemPrevInst = PrevInst
                    PrevInst = sbTempInst.ToString
                    sbTempInst.Length = 0
                    PrevInstNum = NbInsts
                    NbInsts += 1

                End If

                If StringCut = False Then

                    ReDim Preserve ProgInsts(NbInsts)

                    If PrevType > 0 And FileChr <> " " And FileChr <> vbLf Then
                        ProgInsts(PrevInstNum) = ProgInsts(PrevInstNum) & "</span>"
                        PrevType = 0
                    End If

                    ProgInsts(NbInsts) = ChrToHTML(FileChr)
                    NbInsts += 1

                End If

                PrevSep = FileChr
            Else

                'TempInst = TempInst & ChrToHTML(FileChr)
                sbTempInst.Append(ChrToHTML(FileChr))
                'Debug.WriteLine(sbTempInst.ToString)

            End If

            PrevComment = False

            If PrevString > 0 Then
                PrevString += 1
                If PrevString = 3 Then PrevString = 0
            End If

        Next AnaFile

        For ReorgTDM = 1 To MinorIndex
            MinLink = 1

            For ReorgTDM2 = 1 To MinorIndex
                Dim bTest1 As Boolean = False
                Dim bTest2 As Boolean = False
                Dim bTest3 As Boolean = False
                Dim iMax% = TDMLinks.GetUpperBound(0)
                If ReorgTDM2 <= iMax AndAlso MinLink <= iMax AndAlso
                   LCase(TDMLinks(ReorgTDM2)) < LCase(TDMLinks(MinLink)) Then bTest1 = True
                If MinLink <= iMax AndAlso TDMLinks(MinLink) = "" Then bTest2 = True
                If ReorgTDM2 <= iMax AndAlso TDMLinks(ReorgTDM2) <> "" Then bTest3 = True
                If (bTest1 Or bTest2) And bTest3 Then MinLink = ReorgTDM2
                ' ReorgTDM2 pointe au delà de la taille de TDMLinks 
                '  (à cause de vbTab : corrigé plus haut maintenant)
                'If (LCase(TDMLinks(ReorgTDM2)) < LCase(TDMLinks(MinLink)) Or _
                '    TDMLinks(MinLink) = "") And TDMLinks(ReorgTDM2) <> "" Then MinLink = ReorgTDM2
            Next ReorgTDM2

            TransformFileToHTML = TransformFileToHTML & "    " & MajorIndex & "." &
                ReorgTDM & " - <a href=""#" & MajorIndex & MinLink & Chr(34) &
                CStr(IIf(PageTarget <> "", " target=""" & PageTarget & Chr(34), "")) & ">" &
                TDMLinks(MinLink) & "</a>" & vbLf

            TDMLinks(MinLink) = ""

        Next ReorgTDM

        If ActType > 0 Then ProgInsts(PrevInstNum) = ProgInsts(PrevInstNum) & "</span>"

        If IncludeTDM Then
            ecriture.VB6Print(WriteCanal, "<div class=""a"">" &
                CStr(IIf(ObjectTitle <> "",
                    ObjectTitle & " (" &
                        Mid(FileName, InStrRev(FileName, "\") + 1) & ")",
                        Mid(FileName, InStrRev(FileName, "\") + 1))) &
                "</div>" & vbLf & vbLf & vbLf)
            ecriture.VB6Print(WriteCanal, "<span class=""b"">Table des procédures</span>" & vbLf & vbLf)
            ecriture.VB6Print(WriteCanal, TransformFileToHTML & vbLf & vbLf)
        End If

        ecriture.VB6Print(WriteCanal, CStr(IIf(m_bMakeTDM,
            "<a name=""" & MajorIndex & """></a>",
            "")) & "<span class=""b"">" & CStr(IIf(ObjectTitle <> "",
            ObjectTitle & " (" &
                Mid(FileName, InStrRev(FileName, "\") + 1) & ")",
                Mid(FileName, InStrRev(FileName, "\") + 1))) &
            "</span>" & vbLf & vbLf & vbLf)

        For AnaFile = 0 To NbInsts - 1
            ecriture.VB6Print(WriteCanal, ProgInsts(AnaFile))
        Next AnaFile

        ecriture.VB6Print(WriteCanal, vbLf & vbLf)

    End Function

    Public Function GetFileType(FileName As String) As String

        'GetFileType = LCase(Mid(FileName, InStrRev(FileName, ".")))
        GetFileType = IO.Path.GetExtension(FileName).ToLower

    End Function

    Private Function ChrToHTML(StrData As String) As String

        If StrData = "&" Then
            ChrToHTML = "&amp;"
        ElseIf StrData = "<" Then
            ChrToHTML = "&lt;"
        ElseIf StrData = ">" Then
            ChrToHTML = "&gt;"
        Else
            ChrToHTML = StrData
        End If

    End Function

    Public Function IsFileDotNet(FileName As String,
        Optional bPrompt As Boolean = True) As Boolean

        IsFileDotNet = False
        Select Case GetFileType(FileName)
            Case ".vb", ".vbproj" : IsFileDotNet = True
            Case ".txt"
                If bPrompt Then
                    If MsgBoxResult.Yes = MsgBox(
                    "Voulez-vous que ce fichier texte soit considéré comme du Visual Basic.Net ?",
                    MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "VB.Net ?") Then IsFileDotNet = True
                End If
                'Case Else : IsFileDotNet = False
        End Select

    End Function

    Public Function IsProjectFile(FileName As String) As Boolean

        Select Case GetFileType(FileName)
            Case ".vbp", ".vbproj" : IsProjectFile = True
            Case Else : IsProjectFile = False
        End Select

    End Function

    Public Function GetFileName(FileName As String) As String

        Dim AntiSlashPos%

        AntiSlashPos = InStrRev(FileName, "\")
        If AntiSlashPos > 0 Then GetFileName = LCase(Mid(FileName, AntiSlashPos + 1)) Else GetFileName = LCase(FileName)

    End Function

    Public Function GetBoundedString(ByRef PrString As String, ByRef BeginStr As String, ByRef EndStr As String) As String

        Dim BeginPos%

        BeginPos = InStr(1, PrString, BeginStr)
        If BeginPos = 0 Then GetBoundedString = "" : Exit Function
        GetBoundedString = Mid(PrString, BeginPos + Len(BeginStr),
            InStr(BeginPos + Len(BeginStr), PrString, EndStr) - (BeginPos + Len(BeginStr)))

    End Function

End Module