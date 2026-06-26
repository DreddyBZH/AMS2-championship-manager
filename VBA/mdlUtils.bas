Attribute VB_Name = "mdlUtils"
Option Explicit

Public Function GetColumnByHeader( _
    ws As Worksheet, _
    HeaderText As String) As Long

    Dim c As Long

    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

        If Trim(CStr(ws.Cells(1, c).Value)) = HeaderText Then
            GetColumnByHeader = c
            Exit Function
        End If

    Next c

End Function


Public Function SheetExists( _
    SheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = Sheets(SheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0

End Function
Public Function GetPointsBarreme(ByVal Position As Variant) As Double

    Dim ws As Worksheet
    Dim r As Long
    Dim cPos As Long
    Dim cPts As Long

    Set ws = Sheets("Tableau des points")

    cPos = GetColumnByHeader(ws, "Position")
    cPts = GetColumnByHeader(ws, "Points")

    If cPos = 0 Or cPts = 0 Then Exit Function

    r = 2

    Do While Trim(ws.Cells(r, cPos).Text) <> ""

        If CStr(ws.Cells(r, cPos).Text) = _
           CStr(Position) Then

            GetPointsBarreme = _
                CDbl(ws.Cells(r, cPts).Value)

            Exit Function

        End If

        r = r + 1

    Loop

End Function
Public Function MancheCompte( _
    ws As Worksheet) As Boolean

    Dim wsM As Worksheet
    Dim num As Long
    Dim cCompte As Long
    Dim cTerminee As Long

    If Not ws.Name Like "Manche *-*" Then Exit Function

    Set wsM = Sheets("Manches")

    num = NumeroManche(ws.Name)

    cCompte = _
        GetColumnByHeader( _
            wsM, _
            "Compte (OUI/NON)")

    cTerminee = _
        GetColumnByHeader( _
            wsM, _
            "Terminée (OUI/NON)")

    If cCompte = 0 _
       Or cTerminee = 0 Then Exit Function

    MancheCompte = _
        (UCase(Trim( _
            wsM.Cells(num + 2, cCompte).Text)) = "OUI") _
        And _
        (UCase(Trim( _
            wsM.Cells(num + 2, cTerminee).Text)) = "OUI")

End Function
Public Function ToDouble(v As Variant) As Double

    If Trim(CStr(v)) = "" Then
        ToDouble = 0
    Else
        ToDouble = _
            CDbl(Replace(CStr(v), ".", ","))
    End If

End Function


Public Function LireFichier( _
    ByVal chemin As String) As String

    Dim f As Integer
    Dim txt As String
    Dim p1 As Long
    Dim p2 As Long

    f = FreeFile

    Open chemin For Binary As #f
        txt = Space$(LOF(f))
        Get #f, , txt
    Close #f

    p1 = InStr(txt, "{")
    p2 = InStr(txt, "[")

    If p1 = 0 Then p1 = Len(txt) + 1
    If p2 = 0 Then p2 = Len(txt) + 1

    If p1 < p2 Then
        txt = Mid$(txt, p1)
    ElseIf p2 < Len(txt) + 1 Then
        txt = Mid$(txt, p2)
    End If

    LireFichier = txt

End Function
Public Function TrouverSteamID( _
    ByVal participants As Object, _
    ByVal participantID As Long) As String

    Dim p As Object

    If participants.Exists(CStr(participantID)) Then

        Set p = participants(CStr(participantID))

        If p.Exists("SteamID") Then
            TrouverSteamID = CStr(p("SteamID"))
        End If

    End If

End Function
Public Function MsVersExcel( _
    ByVal ms As Double) _
    As Double

    MsVersExcel = _
        ms / 86400000#

End Function

Public Function estIA( _
    ByVal participants As Object, _
    ByVal participantID As Long) _
    As Boolean

    Dim p As Object

    If participants.Exists(CStr(participantID)) Then

        Set p = _
            participants(CStr(participantID))

        If p.Exists("IsPlayer") Then

            estIA = _
                (CLng(p("IsPlayer")) = 0)

        End If

    End If

End Function

Public Function ClePilote( _
    ByVal participants As Object, _
    ByVal participantID As Long) _
    As String

    Dim p As Object
    Dim nom As String

    If Not participants.Exists( _
        CStr(participantID)) Then Exit Function

    Set p = _
        participants(CStr(participantID))

    nom = ""

    If p.Exists("Name") Then
        nom = CStr(p("Name"))
    End If

    If estIA(participants, participantID) Then

        ClePilote = _
            "IA_" & nom

    ElseIf p.Exists("SteamID") Then

        ClePilote = _
            CStr(p("SteamID"))

    End If

End Function

Public Sub AjouterTexte( _
    ByVal ws As Worksheet, _
    ByVal txt As String, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal largeur As Double, _
    ByVal hauteur As Double, _
    ByVal taille As Double, _
    ByVal couleur As Long, _
    ByVal gras As Boolean)

    Dim shp As Shape

    Set shp = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        x, _
        y, _
        largeur, _
        hauteur)

    With shp

        .Line.visible = msoFalse
        .Fill.visible = msoFalse

        With .TextFrame2

            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0

            .TextRange.Text = txt

            With .TextRange.Font

                .Size = taille
                .Fill.ForeColor.RGB = couleur
                .Bold = gras

            End With

        End With

    End With

End Sub
Public Function DerniereLigneDonnees( _
    ws As Worksheet, _
    col As Long) As Long

    Dim r As Long

    r = 2

    Do While _
        Trim(ws.Cells(r, col).Text) <> ""

        r = r + 1

    Loop

    DerniereLigneDonnees = r - 1

End Function
Public Function NumeroManche( _
    ByVal nomFeuille As String) _
    As Long

    Dim p As Long

    p = InStr(nomFeuille, "-")

    NumeroManche = _
        CLng(Trim( _
            Mid( _
                nomFeuille, _
                8, _
                p - 8)))

End Function

Public Function CreerDepuisModele( _
    ByVal nomModele As String, _
    ByVal nomNouvelleFeuille As String, _
    Optional ByVal visible As Boolean = True, _
    Optional ByVal remplacer As Boolean = False) _
    As Worksheet

    Dim wsModele As Worksheet
    Dim wsNouvelle As Worksheet

    On Error GoTo Fin

    ' Vérification du modèle
    If Not SheetExists(nomModele) Then Exit Function

    ' Gestion d'une feuille déjà existante
    If SheetExists(nomNouvelleFeuille) Then

        If remplacer Then

            Application.DisplayAlerts = False
            Sheets(nomNouvelleFeuille).Delete
            Application.DisplayAlerts = True

        Else

            Set CreerDepuisModele = _
                Sheets(nomNouvelleFeuille)

            Exit Function

        End If

    End If

    ' Duplication du modèle
    Set wsModele = _
        Sheets(nomModele)

    wsModele.Copy _
        After:=wsModele

    Set wsNouvelle = _
        ActiveSheet

    ' Renommage
    wsNouvelle.Name = _
        nomNouvelleFeuille

    ' Visibilité
    If visible Then
        wsNouvelle.visible = _
            xlSheetVisible
    Else
        wsNouvelle.visible = _
            xlSheetHidden
    End If

    Set CreerDepuisModele = _
        wsNouvelle

Fin:

    Application.DisplayAlerts = True

    If Err.Number <> 0 Then

        MsgBox _
            "Erreur lors de la création de la feuille :" & vbCrLf & _
            nomNouvelleFeuille & vbCrLf & vbCrLf & _
            Err.Description, _
            vbExclamation, _
            "Création de feuille"

        Set CreerDepuisModele = Nothing

    End If

End Function

Public Sub ColorerCelluleManche(ByVal c As Range)

    Dim v As String
    Dim pts As Double

    v = Trim(CStr(c.Value))

    ' Remise à zéro
    c.Interior.Pattern = xlSolid
    c.Font.Color = RGB(0, 0, 0)

    If v = "" Then
        c.Interior.ColorIndex = xlNone
        Exit Sub
    End If

    Select Case UCase(v)

        Case "DNS"

            c.Interior.Color = RGB(180, 180, 180)
            c.Font.Color = RGB(255, 255, 255)

        Case "DNF"

            c.Interior.Color = RGB(255, 165, 0)
            c.Font.Color = RGB(255, 255, 255)

        Case "DSQ"

            c.Interior.Color = RGB(220, 50, 50)
            c.Font.Color = RGB(255, 255, 255)

        Case Else

            If IsNumeric(v) Then

                pts = GetPointsBarreme(CLng(v))

                If pts > 0 Then

                    c.Interior.Color = RGB(0, 176, 80)
                    c.Font.Color = RGB(255, 255, 255)

                Else

                    c.Interior.Color = RGB(255, 255, 255)
                    c.Font.Color = RGB(0, 0, 0)

                End If

            Else

                c.Interior.Color = RGB(255, 255, 255)
                c.Font.Color = RGB(0, 0, 0)

            End If

    End Select

End Sub

Public Sub DebugParametres()

    Dim i As Long

    For i = 1 To 20

        Debug.Print _
            i & _
            "=[" & _
            Sheets("Parametres").Cells(i, 1).Text & _
            "] -> [" & _
            Sheets("Parametres").Cells(i, 2).Text & _
            "]"

    Next i

End Sub
