Attribute VB_Name = "mdlAccueil"
Option Explicit

Public Sub CreerAccueilFightClub()

    Dim ws As Worksheet

    Set ws = Sheets("Accueil")

    Application.ScreenUpdating = False

    ws.Cells.ClearFormats

    On Error Resume Next
    ws.Shapes.SelectAll
    Selection.Delete
    On Error GoTo 0

    With ws

        .Cells.Interior.Color = RGB(235, 235, 235)

        Dim c As Long
        For c = 1 To 35     ' A:AI
            .Columns(c).ColumnWidth = 14
        Next c

        .Rows.RowHeight = 28

    End With

    CreerBanniere ws
    CreerBlocActions ws
    CreerBlocParametres ws
    InitialiserSwitchs
    CreerDashboard ws
    MajDashboard

    Application.ScreenUpdating = True

End Sub
Private Sub CreerBanniere(ws As Worksheet)

    Dim shp As Shape

    With ws.Range("A1:AG10")

        Set shp = ws.Shapes.AddShape( _
            msoShapeRectangle, _
            .Left, _
            .Top, _
            .Width, _
            .Height)

    End With

    With shp
        .Name = "Banniere"
        .Fill.ForeColor.RGB = RGB(5, 5, 15)
        .Line.visible = msoFalse
    End With

End Sub
Private Sub CreerBlocActions(ws As Worksheet)

    AjouterBouton ws, _
        "Générer le championnat", _
        60, _
        310, _
        RGB(0, 80, 200), _
        "GenererChampionnat", _
        "trophy.png"

    AjouterBouton ws, _
        "Importer une manche AMS2", _
        60, _
        365, _
        RGB(0, 120, 40), _
        "ImporterSMSStats", _
        "import.png"

    AjouterBouton ws, _
        "Mettre à jour les classements", _
        60, _
        420, _
        RGB(210, 100, 0), _
        "BtnMajClassements", _
        "ranking.png"

    AjouterBouton ws, _
        "Exporter la configuration AMS2", _
        60, _
        475, _
        RGB(120, 0, 180), _
        "ExporterAMS2", _
        "export.png"

    AjouterBouton ws, _
        "Réinitialiser une manche", _
        60, _
        530, _
        RGB(180, 0, 40), _
        "ReinitialiserManche", _
        "reset.png"


End Sub

Private Sub CreerBlocParametres( _
    ByVal ws As Worksheet)

    Dim shp As Shape
    Dim xBloc As Double
    Dim yBloc As Double
    Dim largeur As Double
    Dim hauteur As Double

    xBloc = 1370
    yBloc = 305
    largeur = 480
    hauteur = 240

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        xBloc, _
        yBloc, _
        largeur, _
        hauteur)

    With shp

        .Name = "BlocParametres"

        .Fill.ForeColor.RGB = RGB(8, 18, 35)

        .Line.ForeColor.RGB = RGB(0, 180, 255)
        .Line.Weight = 1.5

        ' Même arrondi que les boutons
        .Adjustments.Item(1) = 0.06

        .Shadow.visible = msoTrue

    End With

    '==========================
    ' TITRE
    '==========================

    AjouterTexte ws, _
        "PARAMÈTRES DU CHAMPIONNAT", _
        xBloc + 35, _
        yBloc + 20, _
        450, _
        35, _
        18, _
        RGB(255, 255, 255), _
        True
    '==========================
    ' Nom championnat
    '==========================
AjouterTexte ws, _
    "NOM DU CHAMPIONNAT", _
    xBloc + 35, _
    yBloc + 55, _
    200, _
    22, _
    12, _
    RGB(255, 255, 255), _
    True

AjouterChampTexte _
    ws, _
    "txtNomChampionnat", _
    CStr(GetConfig("cfg_NomChampionnat")), _
    xBloc + 220, _
    yBloc + 50, _
    220

    '==========================
    ' N° Saison
    '==========================

AjouterTexte ws, _
    "SAISON", _
    xBloc + 35, _
    yBloc + 90, _
    120, _
    22, _
    12, _
    RGB(255, 255, 255), _
    True

AjouterChampTexte _
    ws, _
    "txtSaison", _
    CStr(GetConfig("cfg_Saison")), _
    xBloc + 220, _
    yBloc + 85, _
    80
    
    '==========================
    ' IMPORTER IA
    '==========================

    AjouterTexte ws, _
        "IMPORTER LES IA", _
        xBloc + 35, _
        yBloc + 135, _
        250, _
        25, _
        16, _
        RGB(255, 255, 255), _
        True

    AjouterTexte ws, _
        "Inclure les pilotes IA lors de l'import", _
        xBloc + 35, _
        yBloc + 162, _
        350, _
        18, _
        10, _
        RGB(180, 180, 180), _
        False
        
    AjouterSwitch _
        ws, _
        "swImporterIA", _
        xBloc + 340, _
        yBloc + 140, _
        True

    '==========================
    ' CLASSER IA
    '==========================

    AjouterTexte ws, _
        "CLASSER LES IA", _
        xBloc + 35, _
        yBloc + 190, _
        250, _
        25, _
        16, _
        RGB(255, 255, 255), _
        True

    AjouterTexte ws, _
        "Inclure les IA dans les classements", _
        xBloc + 35, _
        yBloc + 217, _
        350, _
        18, _
        10, _
        RGB(180, 180, 180), _
        False
    
    AjouterSwitch _
        ws, _
        "swClasserIA", _
        xBloc + 340, _
        yBloc + 195, _
        True

    '==========================
    ' COMPTER POINTS IA
    '==========================

AjouterTexte ws, _
    "ATTRIBUER DES POINTS AUX IA", _
    1180, 740, _
    450, 25, _
    18, _
    RGB(255, 255, 255), _
    True

AjouterSwitch _
    ws, _
    "swPointsIA", _
    1450, _
    765, _
    True


End Sub

Private Sub AjouterBouton( _
    ws As Worksheet, _
    texte As String, _
    x As Double, _
    y As Double, _
    couleur As Long, _
    macro As String, _
    fichierIcone As String)

    Dim shp As Shape
    Dim ico As Shape
    Dim chemin As String

    chemin = _
        ThisWorkbook.Path & _
        "\Images\" & _
        fichierIcone

    Set shp = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            x, _
            y, _
            450, _
            45)

    With shp

        .Name = _
            "BTN_" & _
            Replace(texte, " ", "_")

        .Fill.ForeColor.RGB = couleur
        .Line.visible = msoFalse

        .Shadow.visible = msoTrue
        .Shadow.Blur = 6
        .Shadow.OffsetX = 2
        .Shadow.OffsetY = 2

        .TextFrame2.TextRange.Text = _
            Space(12) & texte

        .TextFrame2.TextRange.Font.Size = 16
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

        .OnAction = macro

    End With

    If Dir(chemin) <> "" Then

        Set ico = _
            ws.Shapes.AddPicture( _
                chemin, _
                msoFalse, _
                msoTrue, _
                x + 12, _
                y + 6, _
                32, _
                32)

        ico.Name = _
            "ICO_" & _
            Replace(texte, " ", "_")

    End If

End Sub
Public Function GetConfig( _
    ByVal nom As String) _
    As Variant

    On Error Resume Next
    GetConfig = _
        ThisWorkbook.Names(nom) _
        .RefersToRange.Value
    On Error GoTo 0

End Function

Public Sub SetConfig( _
    ByVal nom As String, _
    ByVal valeur As String)

    On Error Resume Next

    ThisWorkbook.Names(nom) _
        .RefersToRange.Value = valeur

    On Error GoTo 0

End Sub

Private Sub AjouterSwitch( _
    ws As Worksheet, _
    nom As String, _
    x As Double, _
    y As Double, _
    valeur As Boolean)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        x, _
        y, _
        80, _
        28)

    shp.Name = nom

    If valeur Then

        shp.Fill.ForeColor.RGB = RGB(60, 200, 60)
        shp.TextFrame2.TextRange.Text = "OUI"

    Else

        shp.Fill.ForeColor.RGB = RGB(100, 100, 100)
        shp.TextFrame2.TextRange.Text = "NON"

    End If

    shp.Line.visible = msoFalse
    shp.OnAction = "BasculerSwitch"
    With shp.TextFrame2.TextRange

        .Font.Size = 13
        .Font.Bold = True
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

    End With

End Sub
Public Sub SwitchAccueil( _
    ByVal nom As String, _
    ByVal actif As Boolean)

    Dim shp As Shape

    Set shp = _
        Sheets("Accueil") _
        .Shapes(nom)

    If actif Then

        shp.Fill.ForeColor.RGB = _
            RGB(60, 200, 60)

        shp.TextFrame2.TextRange.Text = _
            "OUI"

    Else

        shp.Fill.ForeColor.RGB = _
            RGB(120, 120, 120)

        shp.TextFrame2.TextRange.Text = _
            "NON"

    End If

End Sub

Public Sub InitialiserSwitchs()

    SwitchAccueil _
        "swImporterIA", _
        (GetConfig("cfg_ImporterIA") = "OUI")

    SwitchAccueil _
        "swClasserIA", _
        (GetConfig("cfg_ClasserIA") = "OUI")
        
    SwitchAccueil _
        "swPointsIA", _
        (GetConfig("cfg_PointsIA") = "OUI")

End Sub
Public Sub BasculerSwitch()

    Dim shp As Shape
    Dim nom As String
    Dim v As Variant
    Dim etat As Boolean

    ' Sécurisation du Caller
    v = Application.Caller

    If VarType(v) <> vbString Then Exit Sub

    nom = CStr(v)

    On Error Resume Next
    Set shp = Sheets("Accueil").Shapes(nom)
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    ' État actuel
    etat = _
        (shp.TextFrame2.TextRange.Text = "OUI")

    ' Inversion
    etat = Not etat

    ' Mise à jour visuelle
    If etat Then

        shp.Fill.ForeColor.RGB = _
            RGB(60, 200, 60)

        shp.TextFrame2.TextRange.Text = _
            "OUI"

    Else

        shp.Fill.ForeColor.RGB = _
            RGB(120, 120, 120)

        shp.TextFrame2.TextRange.Text = _
            "NON"

    End If

    ' Sauvegarde dans les paramètres
    Select Case nom

        Case "swImporterIA"

            SetConfig _
                "cfg_ImporterIA", _
                IIf(etat, "OUI", "NON")

        Case "swClasserIA"

            SetConfig _
                "cfg_ClasserIA", _
                IIf(etat, "OUI", "NON")
                
        Case "swPointsIA"
            SetConfig _
                "cfg_PointsIA", _
                IIf(etat, "OUI", "NON")

    End Select

    ' Rafraîchissement de l'accueil
    MajDashboard

End Sub
Private Sub CreerDashboard( _
    ByVal ws As Worksheet)

    CreerTuile ws, _
        "tuilePilotes", _
        "PILOTES", _
        "INSCRITS", _
        560, _
        305, _
        RGB(0, 180, 255)

    CreerTuile ws, _
        "tuileIA", _
        "PILOTES IA", _
        "ENREGISTRÉS", _
        840, _
        305, _
        RGB(255, 50, 180)

    CreerTuile ws, _
        "tuileEquipes", _
        "ÉQUIPES", _
        "CRÉÉES", _
        1120, _
        305, _
        RGB(255, 150, 0)

    CreerTuile ws, _
        "tuileManches", _
        "MANCHES", _
        "AU PROGRAMME", _
        700, _
        445, _
        RGB(180, 80, 255)

    CreerTuile ws, _
        "tuileTerminees", _
        "MANCHES JOUÉES", _
        "TERMINÉES", _
        980, _
        445, _
        RGB(120, 255, 80)

End Sub
Private Sub CreerTuile( _
    ByVal ws As Worksheet, _
    ByVal nom As String, _
    ByVal titre As String, _
    ByVal sousTitre As String, _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal couleur As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        x, _
        y, _
        230, _
        110)

    With shp

        .Name = nom

        .Fill.ForeColor.RGB = RGB(10, 10, 15)

        .Line.ForeColor.RGB = couleur
        .Line.Weight = 1.5

        .Adjustments.Item(1) = 0.06

        .Shadow.visible = msoTrue

    End With

    AjouterTexte ws, _
        titre, _
        x + 65, _
        y + 10, _
        200, _
        25, _
        14, _
        couleur, _
        True

    AjouterTexte ws, _
        "0", _
        x + 65, _
        y + 35, _
        200, _
        45, _
        32, _
        RGB(255, 255, 255), _
        True

    ws.Shapes(ws.Shapes.Count).Name = _
        nom & "_Valeur"

    AjouterTexte ws, _
        sousTitre, _
        x + 65, _
        y + 80, _
        200, _
        20, _
        10, _
        couleur, _
        False

End Sub
Private Sub MajTuile( _
    ByVal nom As String, _
    ByVal valeur As Long)

    Dim shp As Shape

    Set shp = _
        Sheets("Accueil") _
        .Shapes(nom & "_Valeur")

    shp.TextFrame2.TextRange.Text = _
        CStr(valeur)

End Sub

Public Sub MajDashboard()

    Dim wsPilote As Worksheet
    Dim wsEquipe As Worksheet
    Dim wsManches As Worksheet

    Dim nbPilotes As Long
    Dim nbIA As Long
    Dim nbEquipes As Long
    Dim nbManches As Long
    Dim nbTerminees As Long

    Dim lastRow As Long
    Dim i As Long
    Dim steamID As String

    Set wsPilote = _
        Sheets("Pilote")

    Set wsEquipe = _
        Sheets("Equipe")

    Set wsManches = _
        Sheets("Manches")

'
' Nombre Pilotes
'
lastRow = _
    DerniereLigneDonnees( _
        wsPilote, 5)

For i = 2 To lastRow

    steamID = _
        Trim(wsPilote.Cells(i, 5).Text)

    If Left(steamID, 3) = "IA_" Then
        nbIA = nbIA + 1
    Else
        nbPilotes = nbPilotes + 1
    End If

Next i

'
' Nombre Equipes
'
    nbEquipes = _
        WorksheetFunction.Max( _
        0, _
        WorksheetFunction.CountA( _
        wsEquipe.Columns(1)) - 1)

    nbManches = _
        WorksheetFunction.Max( _
        0, _
        WorksheetFunction.CountA( _
        wsManches.Columns(1)) - 1)


'
' Nombre Manches
'
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

If ws.Name Like "Manche *-*" Then

        If MancheCompte(ws) Then
            nbTerminees = nbTerminees + 1
        End If

    End If

Next ws

    MajTuile "tuilePilotes", nbPilotes
    MajTuile "tuileIA", nbIA
    MajTuile "tuileEquipes", nbEquipes
    MajTuile "tuileManches", nbManches
    MajTuile "tuileTerminees", nbTerminees

End Sub
Private Sub AjouterChampTexte( _
    ws As Worksheet, _
    nom As String, _
    valeur As String, _
    x As Double, _
    y As Double, _
    largeur As Double)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        x, _
        y, _
        largeur, _
        28)

    With shp

        .Name = nom

        .Fill.ForeColor.RGB = _
            RGB(245, 245, 245)

        .Line.ForeColor.RGB = _
            RGB(190, 190, 190)

        .Line.Weight = 1

        .Adjustments.Item(1) = 0.08

        .TextFrame2.TextRange.Text = valeur

        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(0, 70, 180)

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

    End With

End Sub

Public Function ImporterIA() _
    As Boolean

    ImporterIA = _
        (UCase( _
        GetConfig("cfg_ImporterIA")) = _
        "OUI")

End Function
