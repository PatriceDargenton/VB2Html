
' Gestion des fichiers via du code VB6 porté en .Net

Option Infer On

Imports System.Text ' Encoding

#Const bVB6 = False

Public MustInherit Class clsUtilFichierVB6

    'Public MustOverride Function iFreeCanal%()                     Lecture
    'Public MustOverride Function iFreeCanal%(encodage As Encoding) Ecriture

    Public MustOverride Sub VB6FileClose(iFreeCanal%)
    Public MustOverride Sub VB6FileOpen(iFreeCanal%, sChemin$,
        OpenMode As Microsoft.VisualBasic.OpenMode,
        Optional OpenAccess As Microsoft.VisualBasic.OpenAccess = OpenAccess.Default)

    Protected m_lstIndex As New List(Of Integer)

    Public Class clsFichier
        Public sChemin$
        Public encod As Encoding
        Public fs As IO.FileStream
        Public sw As IO.StreamWriter
        Public sb As StringBuilder
    End Class
    Protected m_lstF As New List(Of clsFichier)

End Class

Public Class clsUtilFichierVB6Lecture : Inherits clsUtilFichierVB6

    Public Function iFreeCanal%()

#If bVB6 Then
        Return FreeFile()
#Else
        Dim iNextIndex% = m_lstIndex.Count
        m_lstIndex.Add(iNextIndex)
        Return iNextIndex
#End If

    End Function

    Private Function OpenAccessToFileAccess(
        OpenAccess As Microsoft.VisualBasic.OpenAccess) As IO.FileAccess

        Dim fa As New IO.FileAccess
        Select Case OpenAccess
            Case Microsoft.VisualBasic.OpenAccess.Default : fa = IO.FileAccess.Read
            Case Microsoft.VisualBasic.OpenAccess.Read : fa = IO.FileAccess.Read
            Case Microsoft.VisualBasic.OpenAccess.ReadWrite : fa = IO.FileAccess.ReadWrite
            Case Microsoft.VisualBasic.OpenAccess.Write : fa = IO.FileAccess.Write
        End Select
        Return fa

    End Function

    Public Overrides Sub VB6FileOpen(iFreeCanal%, sChemin$,
        OpenMode As Microsoft.VisualBasic.OpenMode,
        Optional OpenAccess As Microsoft.VisualBasic.OpenAccess = OpenAccess.Default)

#If bVB6 Then
        FileOpen(iFreeCanal, sChemin, OpenMode, OpenAccess)
#Else
        Dim fa = Me.OpenAccessToFileAccess(OpenAccess)
        Dim fs As New IO.FileStream(sChemin, IO.FileMode.Open, fa)
        Dim encod As Encoding = LireEncodageVB6(sChemin)
        Dim fichier As New clsFichier
        fichier.encod = encod
        fichier.sChemin = sChemin
        fichier.fs = fs
        m_lstF.Add(fichier)
#End If

    End Sub

    Public Overrides Sub VB6FileClose(iFreeCanal%)

#If bVB6 Then
        FileClose(iFreeCanal)
#Else
        Dim fichier = m_lstF(iFreeCanal)
        fichier.fs.Close()
#End If

    End Sub

    Public Sub VB6FileGet(iFreeCanal%, ByRef sVal$)

#If bVB6 Then
        FileGet(iFreeCanal, sVal)
#Else
        Dim fichier = m_lstF(iFreeCanal)
        Dim iLen% = sVal.Length
        Dim lFilePos& = fichier.fs.Position
        Dim iFilePos% = 0
        If lFilePos < Integer.MaxValue Then iFilePos = CInt(lFilePos)
        Dim iLenFin% = iLen
        If iFilePos + iLen > fichier.fs.Length Then _
            iLenFin = CInt(fichier.fs.Length - iFilePos)
        Dim aByte(0 To iLenFin - 1) As Byte
        fichier.fs.Read(aByte, 0, iLenFin)
        sVal = fichier.encod.GetString(aByte)
#End If

    End Sub

    Public Function bEOF(iFreeCanal%) As Boolean

#If bVB6 Then
         Return EOF(iFreeCanal)
#Else
        Dim fichier = m_lstF(iFreeCanal)
        Dim bEOF0 As Boolean = (fichier.fs.Length = fichier.fs.Position)
        Return bEOF0
#End If

    End Function

End Class

Public Class clsUtilFichierVB6Ecriture : Inherits clsUtilFichierVB6

    Public Function iFreeCanal%(encodage As Encoding)

#If bVB6 Then
        Return FreeFile()
#Else
        Dim iNextIndex% = m_lstIndex.Count
        m_lstIndex.Add(iNextIndex)
        Dim fichier As New clsFichier
        fichier.encod = encodage
        fichier.sb = New StringBuilder
        m_lstF.Add(fichier)
        Return iNextIndex
#End If

    End Function

    Public Overrides Sub VB6FileOpen(iFreeCanal%, sChemin$,
        OpenMode As Microsoft.VisualBasic.OpenMode,
        Optional OpenAccess As Microsoft.VisualBasic.OpenAccess = OpenAccess.Default)

#If bVB6 Then
        FileOpen(iFreeCanal, sChemin, OpenMode, OpenAccess)
#Else
        Dim fichier = m_lstF(iFreeCanal)
        fichier.sw = New IO.StreamWriter(sChemin, append:=False, encoding:=fichier.encod)
#End If

    End Sub

    Public Overrides Sub VB6FileClose(iFreeCanal%)

#If bVB6 Then
        FileClose(iFreeCanal)
#Else
        Dim fichier = m_lstF(iFreeCanal)
        fichier.sw.Write(fichier.sb)
        fichier.sw.Close()
#End If

    End Sub

    'Public Sub VB6Print(iFreeCanal%, ParamArray Output() As Object)
    '    Print(iFreeCanal, Output)
    'End Sub

    Public Sub VB6Print(iFreeCanal%, sTexte$)

#If bVB6 Then
        Print(iFreeCanal, sTexte)
#Else
        Dim fichier = m_lstF(iFreeCanal)
        'fichier.sw.Write(sTexte)
        fichier.sb.Append(sTexte)
#End If

    End Sub

    Public Sub VB6Print(iFreeCanal%, TAB As Microsoft.VisualBasic.TabInfo, sTexte$)

#If bVB6 Then
        Print(iFreeCanal, TAB, sTexte)
#Else
        Dim fichier = m_lstF(iFreeCanal)
        'fichier.sw.Write("       " & sTexte) ' vbTab & sTexte
        fichier.sb.Append("       " & sTexte) ' vbTab & sTexte
#End If

    End Sub

End Class