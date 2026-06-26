Attribute VB_Name = "mdlChampionship"
Option Explicit

Public Sub GenererChampionnat()

    Dim wsManches As Worksheet
    Dim wsPilotes As Worksheet
    Dim ws As Worksheet

    Dim i As Long
    Dim j As Long
    Dim lig As Long

    Dim nomFeuille As String
    Dim steamID As String

    Dim cPseudo As Long
    Dim cSteam As Long
    Dim cBallast As Long
    Dim cPuissance As Long
    Dim cTrainee As Long

    Dim cM_Pseudo As Long
    Dim cM_Steam As Long
    Dim cM_Ballast As Long
    Dim cM_Puissance As Long
    Dim cM_Trainee As Long

    Dim nbManches As Long
    Dim nbPilotesAjoutes As Long

    Set wsManches = Sheets("Manches")
    Set wsPilotes = Sheets("Pilote")

    cPseudo = GetColumnByHeader(wsPilotes, "Pseudo pilote")
    cSteam = GetColumnByHeader(wsPilotes, "SteamID")
    cBallast = GetColumnByHeader(wsPilotes, "Ballast défaut")
    cPuissance = GetColumnByHeader(wsPilotes, "Puissance défaut")
    cTrainee = GetColumnByHeader(wsPilotes, "Trainée défaut")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    i = 2

    Do While Trim(wsManches.Cells(i, 1).Value) <> ""

        nomFeuille = _
            "Manche " & _
            wsManches.Cells(i, 1).Value & _
            " - " & _
            wsManches.Cells(i, 2).Value

        nomFeuille = Left(nomFeuille, 31)

        Dim feuilleExiste As Boolean

        feuilleExiste = _
            SheetExists(nomFeuille)

        Set ws = _
            CreerDepuisModele( _
                "MODELE_MANCHE", _
                nomFeuille)

        If ws Is Nothing Then
            GoTo MancheSuivante
        End If

        If Not feuilleExiste Then
            nbManches = _
                nbManches + 1
        End If

        cM_Pseudo = _
            GetColumnByHeader(ws, _
            "Pseudo pilote")

        cM_Steam = _
            GetColumnByHeader(ws, _
            "SteamID")

        cM_Ballast = _
            GetColumnByHeader(ws, _
            "Ballast")

        cM_Puissance = _
            GetColumnByHeader(ws, _
            "Puissance")

        cM_Trainee = _
            GetColumnByHeader(ws, _
            "Trainée")

        Dim lastPilotRow As Long

        lastPilotRow = _
            wsPilotes.Cells( _
                wsPilotes.Rows.Count, _
                cPseudo).End(xlUp).Row

        For j = 2 To lastPilotRow

            steamID = _
                Trim(CStr( _
                    wsPilotes.Cells( _
                        j, _
                        cSteam).Text))

            If steamID <> "" Then

                If Not SteamIDExists( _
                    ws, _
                    cM_Steam, _
                    steamID) Then

                    '
                    ' Première ligne libre
                    '
                    lig = 2

                    Do While _
                        Trim(CStr( _
                            ws.Cells( _
                                lig, _
                                cM_Pseudo).Value)) <> ""

                        lig = lig + 1

                    Loop

                    '
                    ' Pseudo
                    '
                    ws.Cells( _
                        lig, _
                        cM_Pseudo).Value = _
                        wsPilotes.Cells( _
                            j, _
                            cPseudo).Value

                    '
                    ' SteamID
                    '
                    With ws.Cells( _
                        lig, _
                        cM_Steam)

                        .NumberFormat = "@"
                        .Value = steamID

                    End With

                    '
                    ' Valeurs par défaut
                    '
                    ws.Cells( _
                        lig, _
                        cM_Ballast).Value = _
                        wsPilotes.Cells( _
                            j, _
                            cBallast).Value

                    ws.Cells( _
                        lig, _
                        cM_Puissance).Value = _
                        wsPilotes.Cells( _
                            j, _
                            cPuissance).Value

                    ws.Cells( _
                        lig, _
                        cM_Trainee).Value = _
                        wsPilotes.Cells( _
                            j, _
                            cTrainee).Value

                    nbPilotesAjoutes = _
                        nbPilotesAjoutes + 1

                End If

            End If

        Next j

MancheSuivante:

        i = i + 1

    Loop

    '
    ' Classement
    '
    Dim wsClassement As Worksheet

    Set wsClassement = _
        CreerDepuisModele( _
            "MODELE_CLASSEMENT", _
            "Classement", _
            True, _
            True)

    If wsClassement Is Nothing Then
        Exit Sub
    End If

    InitialiserColonnesClassement

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox _
        nbManches & _
        " nouvelle(s) manche(s) créée(s)." & vbCrLf & _
        nbPilotesAjoutes & _
        " pilote(s) ajouté(s).", _
        vbInformation, _
        "Génération du championnat"

End Sub


Public Sub InitialiserColonnesClassement()

    Dim ws As Worksheet
    Dim wsM As Worksheet

    Dim nbManches As Long
    Dim cInsert As Long
    Dim i As Long

    Set ws = Sheets("Classement")
    Set wsM = Sheets("Manches")

    If GetColumnByHeader(ws, "M0") > 0 Then Exit Sub

    nbManches = _
        DerniereLigneDonnees(wsM, 1) - 1

    cInsert = _
        GetColumnByHeader( _
            ws, _
            "Points") + 1

    ws.Columns(cInsert).Resize( _
        , nbManches).Insert _
        Shift:=xlToRight

    For i = 0 To nbManches - 1

        ws.Cells(1, cInsert + i) = _
            "M" & i

    Next i

End Sub
Private Sub ViderClassement( _
    ByVal ws As Worksheet)

    Dim lastRow As Long

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, 1) _
            .End(xlUp).Row

    If lastRow >= 2 Then

        ws.Rows( _
            "2:" & lastRow).Delete

    End If

End Sub

Private Function SteamIDExists( _
    ws As Worksheet, _
    colSteam As Long, _
    steamID As String) As Boolean

    Dim lastRow As Long
    Dim i As Long

    If colSteam = 0 Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, colSteam).End(xlUp).Row

    For i = 2 To lastRow

        If Trim(ws.Cells(i, colSteam).Text) = steamID Then
            SteamIDExists = True
            Exit Function
        End If

    Next i

End Function


Public Sub CalculerClassementGestionIA( _
    ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim i As Long

    Dim cSteam As Long
    Dim cReel As Long
    Dim cCalc As Long

    Dim steamID As String
    Dim classementReel As String

    Dim rangHumain As Long
    Dim classerIA As Boolean

    cSteam = _
        GetColumnByHeader( _
            ws, _
            "SteamID")

    cReel = _
        GetColumnByHeader( _
            ws, _
            "ClassementReel")

    cCalc = _
        GetColumnByHeader( _
            ws, _
            "ClassementCalc")

    classerIA = _
        (UCase(Trim( _
            CStr(GetConfig( _
                "cfg_ClasserIA")))) _
            = "OUI")

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, 1) _
            .End(xlUp).Row

    rangHumain = 0

    For i = 2 To lastRow

        steamID = _
            Trim(ws.Cells( _
                i, _
                cSteam).Text)

        classementReel = _
            Trim(CStr( _
                ws.Cells( _
                    i, _
                    cReel).Value))

        '
        ' IA transparentes
        '
        If Left(steamID, 3) = "IA_" _
           And Not classerIA Then

            ws.Cells( _
                i, _
                cCalc).ClearContents

        Else

            If classerIA Then

                ws.Cells( _
                    i, _
                    cCalc).Value = _
                    classementReel

            Else

                If IsNumeric( _
                    classementReel) Then

                    rangHumain = _
                        rangHumain + 1

                    ws.Cells( _
                        i, _
                        cCalc).Value = _
                        rangHumain

                Else

                    ws.Cells( _
                        i, _
                        cCalc).Value = _
                        classementReel

                End If

            End If

        End If

    Next i

End Sub

Public Sub CalculerPoints( _
    ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim i As Long

    Dim cSteam As Long
    Dim cCalc As Long
    Dim cPoints As Long

    Dim steamID As String
    Dim classementCalc As String

    Dim estIA As Boolean
    Dim pointsIA As Boolean

    Dim pts As Double

    cSteam = _
        GetColumnByHeader( _
            ws, _
            "SteamID")

    cCalc = _
        GetColumnByHeader( _
            ws, _
            "ClassementCalc")

    cPoints = _
        GetColumnByHeader( _
            ws, _
            "Points")

    pointsIA = _
        (UCase(Trim( _
            CStr(GetConfig( _
                "cfg_PointsIA")))) _
            = "OUI")

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, 1) _
            .End(xlUp).Row

    For i = 2 To lastRow

        steamID = _
            Trim(ws.Cells( _
                i, _
                cSteam).Text)

        classementCalc = _
            Trim(CStr( _
                ws.Cells( _
                    i, _
                    cCalc).Value))

        estIA = _
            (Left( _
                steamID, 3) = "IA_")

        pts = 0

        If classementCalc <> "" Then

            pts = _
                GetPointsBarreme( _
                    classementCalc)

        End If

        If estIA _
           And Not pointsIA Then

            pts = 0

        End If

        ws.Cells( _
            i, _
            cPoints).Value = _
            pts

    Next i

End Sub

Public Sub AfficherClassement( _
    ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim i As Long

    Dim cClassement As Long
    Dim cCalc As Long

    cClassement = _
        GetColumnByHeader( _
            ws, _
            "Classement")

    cCalc = _
        GetColumnByHeader( _
            ws, _
            "ClassementCalc")

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, 1) _
            .End(xlUp).Row

    For i = 2 To lastRow

        ws.Cells( _
            i, _
            cClassement).Value = _
            ws.Cells( _
                i, _
                cCalc).Value

    Next i

End Sub


Public Sub BtnMajClassements()

    RecalculClassements

End Sub

Public Sub CalculerClassementPilotes( _
    ByVal wsClassement As Worksheet)

    Dim d As Object

    Set d = _
        ConstruireListeConcurrents()

    ViderClassement wsClassement
    InitialiserColonnesClassement

    EcrireClassementPilotes _
        wsClassement, _
        d

End Sub
Private Function ConstruireListeConcurrents() _
    As Object

    Dim d As Object
    Dim info As Object

    Dim ws As Worksheet

    Dim i As Long
    Dim lastRow As Long

    Dim steamID As String
    Dim pseudo As String

    Dim estIA As Boolean
    Dim classerIA As Boolean

    Set d = _
        CreateObject( _
            "Scripting.Dictionary")

    classerIA = _
        (UCase( _
            Trim( _
                CStr( _
                    GetConfig( _
                        "cfg_ClasserIA")))) _
            = "OUI")

    For Each ws _
        In ThisWorkbook.Worksheets

        If MancheCompte(ws) Then

            lastRow = _
                ws.Cells( _
                    ws.Rows.Count, 1) _
                    .End(xlUp).Row

            For i = 2 To lastRow

                steamID = _
                    Trim( _
                        ws.Cells( _
                            i, _
                            GetColumnByHeader( _
                                ws, _
                                "SteamID")).Text)

                pseudo = _
                    Trim( _
                        ws.Cells( _
                            i, _
                            GetColumnByHeader( _
                                ws, _
                                "Pseudo pilote")).Text)

                If steamID <> "" Then

                    estIA = _
                        (Left( _
                            steamID, 3) = "IA_")

                    If _
                        (Not estIA) _
                        Or classerIA Then

                        If Not d.Exists( _
                            steamID) Then

                            Set info = _
                                CreateObject( _
                                    "Scripting.Dictionary")

                            info.Add _
                                "Pseudo", _
                                pseudo

                            info.Add _
                                "EstIA", _
                                estIA

                            d.Add _
                                steamID, _
                                info

                        End If

                    End If

                End If

            Next i

        End If

    Next ws

    Set ConstruireListeConcurrents = d

End Function
Private Sub CalculerStatsConcurrent( _
    ByVal steamID As String, _
    ByRef totalPts As Double, _
    ByRef nbP1 As Long, _
    ByRef nbP2 As Long, _
    ByRef nbP3 As Long, _
    ByRef ballastTotal As Double, _
    ByRef puissanceMoy As Double, _
    ByRef traineeMoy As Double, _
    ByRef dPositions As Object)

    Dim ws As Worksheet
    Dim r As Long
    Dim pos As Variant

    Dim nbManches As Long

    totalPts = 0

    nbP1 = 0
    nbP2 = 0
    nbP3 = 0

    ballastTotal = 0
    puissanceMoy = 0
    traineeMoy = 0

    nbManches = 0

    Set dPositions = _
        CreateObject( _
            "Scripting.Dictionary")

    For Each ws _
        In ThisWorkbook.Worksheets

        If MancheCompte(ws) Then

            r = _
                TrouverLignePilote( _
                    ws, _
                    steamID)

            If r > 0 Then

                pos = _
                    ws.Cells( _
                        r, _
                        GetColumnByHeader( _
                            ws, _
                            "Classement")).Value

                dPositions( _
                    NumeroManche( _
                        ws.Name)) = pos

                totalPts = _
                    totalPts + _
                    ToDouble( _
                        ws.Cells( _
                            r, _
                            GetColumnByHeader( _
                                ws, _
                                "Points")).Value)

                If IsNumeric(pos) Then

                    Select Case CLng(pos)

                        Case 1
                            nbP1 = _
                                nbP1 + 1

                        Case 2
                            nbP2 = _
                                nbP2 + 1

                        Case 3
                            nbP3 = _
                                nbP3 + 1

                    End Select

                End If

                ballastTotal = _
                    ballastTotal + _
                    ToDouble( _
                        ws.Cells( _
                            r, _
                            GetColumnByHeader( _
                                ws, _
                                "Ballast")).Value)

                puissanceMoy = _
                    puissanceMoy + _
                    ToDouble( _
                        ws.Cells( _
                            r, _
                            GetColumnByHeader( _
                                ws, _
                                "Puissance")).Value)

                traineeMoy = _
                    traineeMoy + _
                    ToDouble( _
                        ws.Cells( _
                            r, _
                            GetColumnByHeader( _
                                ws, _
                                "Trainée")).Value)

                nbManches = _
                    nbManches + 1

            End If

        End If

    Next ws

    If nbManches > 0 Then

        puissanceMoy = _
            Round( _
                puissanceMoy / _
                nbManches, _
                1)

        traineeMoy = _
            Round( _
                traineeMoy / _
                nbManches, _
                1)

    End If

End Sub
Private Sub EcrireClassementPilotes( _
    ByVal wsClassement As Worksheet, _
    ByVal d As Object)

    Dim steamID As Variant
    Dim info As Object

    Dim r As Long

    Dim totalPts As Double
    Dim nbP1 As Long
    Dim nbP2 As Long
    Dim nbP3 As Long

    Dim ballastTotal As Double
    Dim puissanceMoy As Double
    Dim traineeMoy As Double

    Dim dPositions As Object
    Dim numManche As Variant

    Dim cPseudo As Long
    Dim cPoints As Long
    Dim cP1 As Long
    Dim cP2 As Long
    Dim cP3 As Long
    Dim cBallast As Long
    Dim cPuissance As Long
    Dim cTrainee As Long
    Dim cM As Long

    cPseudo = _
        GetColumnByHeader( _
            wsClassement, _
            "Pseudo pilote")

    cPoints = _
        GetColumnByHeader( _
            wsClassement, _
            "Points")

    cP1 = _
        GetColumnByHeader( _
            wsClassement, _
            "P1")

    cP2 = _
        GetColumnByHeader( _
            wsClassement, _
            "P2")

    cP3 = _
        GetColumnByHeader( _
            wsClassement, _
            "P3")

    cBallast = _
        GetColumnByHeader( _
            wsClassement, _
            "Ballast total")

    cPuissance = _
        GetColumnByHeader( _
            wsClassement, _
            "Puissance moyenne")

    cTrainee = _
        GetColumnByHeader( _
            wsClassement, _
            "Trainée moyenne")

    cM = _
        GetColumnByHeader( _
            wsClassement, _
            "M0")

    r = 2

    For Each steamID _
        In d.Keys

        Set info = d(steamID)

        CalculerStatsConcurrent _
            CStr(steamID), _
            totalPts, _
            nbP1, _
            nbP2, _
            nbP3, _
            ballastTotal, _
            puissanceMoy, _
            traineeMoy, _
            dPositions

        wsClassement.Cells( _
            r, _
            cPseudo).Value = _
            info("Pseudo")

        wsClassement.Cells( _
            r, _
            cPoints).Value = _
            totalPts

        wsClassement.Cells( _
            r, _
            cP1).Value = _
            nbP1

        wsClassement.Cells( _
            r, _
            cP2).Value = _
            nbP2

        wsClassement.Cells( _
            r, _
            cP3).Value = _
            nbP3

        wsClassement.Cells( _
            r, _
            cBallast).Value = _
            ballastTotal

        wsClassement.Cells( _
            r, _
            cPuissance).Value = _
            puissanceMoy

        wsClassement.Cells( _
            r, _
            cTrainee).Value = _
            traineeMoy

Debug.Print
Debug.Print "==== " & info("Pseudo") & " ===="

For Each numManche In dPositions.Keys

    Debug.Print _
        "M" & numManche & _
        "=[" & dPositions(numManche) & "]"

Next numManche

        For Each numManche _
            In dPositions.Keys

            wsClassement.Cells( _
                r, _
                cM + CLng(numManche)).Value = _
                dPositions(numManche)

            ColorerCelluleManche _
                wsClassement.Cells( _
                    r, _
                    cM + CLng(numManche))

        Next numManche

        r = r + 1

    Next steamID

End Sub

Private Sub TrierClassementPilotes()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long

    Dim cPoints As Long
    Dim cClassement As Long

    Set ws = Sheets("Classement")

    lastRow = _
        DerniereLigneDonnees(ws, 1)

    If lastRow < 2 Then Exit Sub

    cPoints = _
        GetColumnByHeader( _
            ws, _
            "Points")

    cClassement = _
        GetColumnByHeader( _
            ws, _
            "Classement")

    lastCol = _
        ws.Cells( _
            1, _
            ws.Columns.Count).End(xlToLeft).Column

    With ws.Sort

        .SortFields.Clear

        .SortFields.Add _
            key:=ws.Range( _
                ws.Cells(2, cPoints), _
                ws.Cells(lastRow, cPoints)), _
            Order:=xlDescending

        .SetRange _
            ws.Range( _
                ws.Cells(1, 1), _
                ws.Cells(lastRow, lastCol))

        .Header = xlYes
        .Apply

    End With

    '
    ' Réécriture des positions du championnat
    '
    For i = 2 To lastRow

        ws.Cells( _
            i, _
            cClassement).Value = i - 1

    Next i

End Sub

Public Sub RecalculClassements()

    Dim ws As Worksheet

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets

        If MancheCompte(ws) Then

            MettreAJourClassementManche ws

        End If

    Next ws

    CalculerClassementPilotes _
        Sheets("Classement")

    TrierClassementPilotes

    RecalculClassementEquipes

    Application.ScreenUpdating = True

End Sub
