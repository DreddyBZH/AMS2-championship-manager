Attribute VB_Name = "mdlImportSMSStats"
Option Explicit

Public Sub ImporterSMSStats()

    Dim f As Variant
    Dim txt As String
    Dim json As Object
    Dim hist As Object
    Dim s As Object
    Dim i As Long

    f = Application.GetOpenFilename( _
            "Fichier JSON (*.json),*.json")

    If f = False Then Exit Sub

    txt = LireFichier(CStr(f))

    Set json = _
        JsonConverter.ParseJson(txt)

    If Not json.Exists("stats") Then

        MsgBox _
            "Section 'stats' introuvable.", _
            vbExclamation

        Exit Sub

    End If

    If Not _
        json("stats").Exists("history") Then

        MsgBox _
            "Historique introuvable.", _
            vbExclamation

        Exit Sub

    End If

    Set hist = _
        json("stats")("history")

    Debug.Print _
        "Nb sessions = " & _
        hist.Count

    For i = 1 To hist.Count

        Set s = hist(i)

        If s("finished") = True Then

            ImporterHistorique s

        End If

    Next i

    MsgBox _
        "Import terminé.", _
        vbInformation

End Sub
Private Sub ImporterHistorique( _
    ByVal h As Object)

    Dim trackID As String
    Dim ws As Worksheet

    '
    ' Recherche de la manche via le TrackID AMS2
    '

    trackID = _
        CStr(h("setup")("TrackId"))

    Set ws = _
        TrouverOngletManche(trackID)

    If ws Is Nothing Then

        Debug.Print _
            "TrackID introuvable : " & _
            trackID

        Exit Sub

    End If

    '
    ' Import Qualifs
    '

    If h("stages").Exists( _
        "qualifying1") Then

        ImporterSessionQualif _
            ws, _
            h("stages")("qualifying1"), _
            h("participants")

    End If

    '
    ' Import Course
    '

    If h("stages").Exists( _
        "race1") Then

        ImporterSessionCourse _
            ws, _
            h("stages")("race1"), _
            h("participants")
    
    PilotesAbsentsImport ws
    MettreAJourClassementManche ws

End If
    
End Sub

Public Sub MettreAJourClassementManche( _
    ByVal ws As Worksheet)

    CalculerClassementGestionIA ws

    CalculerPoints ws

    AfficherClassement ws

End Sub

Private Sub ImporterSessionQualif( _
    ByVal ws As Worksheet, _
    ByVal stage As Object, _
    ByVal participants As Object)

    Dim res As Object
    Dim r As Long

    Dim steamID As String
    Dim participantID As Long
    Dim bestLap As Double
    Dim pos As Variant
    Dim voiture As Variant

    Dim s1 As Double
    Dim s2 As Double
    Dim s3 As Double

    If Not stage.Exists("results") Then Exit Sub

    For Each res In stage("results")

        participantID = _
            CLng(res("participantid"))

steamID = _
    ClePilote( _
        participants, _
        participantID)
        
    
        If steamID <> "" Then
        
If Left(steamID, 3) = "IA_" Then

    If Not ImporterIA Then
        GoTo SuitePilote
    End If

End If
            r = _
    TrouverLignePilote( _
        ws, _
        steamID)

If r = 0 Then

    r = _
        AjouterPiloteManche( _
            ws, _
            participants, _
            participantID)

End If

If r > 0 Then

                bestLap = 0
                pos = ""
                voiture = ""

                If res("attributes").Exists( _
                    "FastestLapTime") Then

                    bestLap = _
                        ToDouble( _
                            res("attributes") _
                            ("FastestLapTime"))

                End If

                If res("attributes").Exists( _
                    "RacePosition") Then

                    pos = _
                        res("attributes") _
                        ("RacePosition")

                End If

                If res("attributes").Exists( _
                    "VehicleId") Then

                    voiture = _
                        res("attributes") _
                        ("VehicleId")

                End If

                Call ChercherSecteurs( _
                    stage, _
                    participantID, _
                    bestLap, _
                    s1, _
                    s2, _
                    s3)

  
    
                SetValeur ws, _
                    r, _
                    "ParticipantId", _
                    participantID

                SetValeur ws, _
                    r, _
                    "Voiture", _
                    voiture

                SetValeur ws, r, _
                    "Temps qualif", _
                    MsVersExcel(bestLap)
    
                SetValeur ws, r, _
                    "Position qualif", _
                    pos

                SetValeur ws, r, _
                    "Q1", _
                    MsVersExcel(s1)

                SetValeur ws, r, _
                    "Q2", _
                    MsVersExcel(s2)

                SetValeur ws, r, _
                    "Q3", _
                    MsVersExcel(s3)
    
            End If

        End If
SuitePilote:
    Next res

End Sub
Private Sub ImporterSessionCourse( _
    ByVal ws As Worksheet, _
    ByVal stage As Object, _
    ByVal participants As Object)

    Dim res As Object
    Dim r As Long

    Dim steamID As String
    Dim participantID As Long
    Dim bestLap As Double
    Dim tempsCourse As Double
    Dim pos As Variant
    Dim etat As String
    Dim voiture As Variant

    Dim s1 As Double
    Dim s2 As Double
    Dim s3 As Double

    If Not stage.Exists("results") Then Exit Sub

    For Each res In stage("results")

        participantID = _
            CLng(res("participantid"))

steamID = _
    ClePilote( _
        participants, _
        participantID)

        If steamID <> "" Then
        
If Left(steamID, 3) = "IA_" Then

    If Not ImporterIA Then
        GoTo SuitePilote
    End If

End If
            r = _
    TrouverLignePilote( _
        ws, _
        steamID)

If r = 0 Then

    r = _
        AjouterPiloteManche( _
            ws, _
            participants, _
            participantID)

End If

If r > 0 Then

                bestLap = 0
                tempsCourse = 0
                pos = ""
                etat = ""
                voiture = ""

                If res("attributes").Exists( _
                    "FastestLapTime") Then

                    bestLap = _
                        ToDouble( _
                            res("attributes") _
                            ("FastestLapTime"))

                End If

                If res("attributes").Exists( _
                    "TotalTime") Then

                    tempsCourse = _
                        ToDouble( _
                            res("attributes") _
                            ("TotalTime"))

                End If

                If res("attributes").Exists( _
                    "RacePosition") Then

                    pos = _
                        res("attributes") _
                        ("RacePosition")

                End If

                If res("attributes").Exists( _
                    "State") Then

                    etat = _
                        CStr( _
                            res("attributes") _
                            ("State"))

                End If

                If res("attributes").Exists( _
                    "VehicleId") Then

                    voiture = _
                        res("attributes") _
                        ("VehicleId")

                End If

                Call ChercherSecteurs( _
                    stage, _
                    participantID, _
                    bestLap, _
                    s1, _
                    s2, _
                    s3)

                '
                ' Écriture
                '

    
                SetValeur ws, _
                    r, _
                    "ParticipantId", _
                    participantID

                SetValeur ws, _
                    r, _
                    "Voiture", _
                    voiture

                SetValeur ws, r, _
                    "Temps course", _
                    MsVersExcel(tempsCourse)

                SetValeur ws, r, _
                    "Meilleur tour", _
                    MsVersExcel(bestLap)

                SetValeur ws, r, _
                    "R1", _
                    MsVersExcel(s1)

                SetValeur ws, r, _
                    "R2", _
                    MsVersExcel(s2)

                SetValeur ws, r, _
                    "R3", _
                    MsVersExcel(s3)


                SetValeur ws, _
                    r, _
                    "ClassementReel", _
                    DeterminerClassementReel( _
                    pos, _
                    etat)
                    
                SetValeur ws, _
                    r, _
                    "État", _
                    etat

            End If

        End If
SuitePilote:
    Next res

End Sub
Public Sub PilotesAbsentsImport( _
    ByVal ws As Worksheet)

    Dim lastRow As Long
    Dim i As Long

    Dim cReel As Long
    Dim cEtat As Long

    cReel = _
        GetColumnByHeader( _
            ws, _
            "ClassementReel")

    cEtat = _
        GetColumnByHeader( _
            ws, _
            "État")

    lastRow = _
        ws.Cells( _
            ws.Rows.Count, 1) _
            .End(xlUp).Row

    For i = 2 To lastRow

        If Trim(CStr( _
            ws.Cells(i, cReel).Value)) = "" Then

            ws.Cells(i, cReel).Value = _
                DeterminerClassementReel( _
                    "", _
                    Trim(CStr( _
                        ws.Cells(i, cEtat).Value)))

        End If

    Next i

End Sub

Public Function DeterminerClassementReel( _
    ByVal pos As Variant, _
    ByVal etat As String) _
    As String

    If Trim(CStr(pos)) <> "" Then

        DeterminerClassementReel = _
            CStr(pos)

        Exit Function

    End If

    Select Case UCase(Trim(etat))

        Case ""
            DeterminerClassementReel = _
                "DNS"

        Case "DNF", _
             "RETIRED"

            DeterminerClassementReel = _
                "DNF"

        Case "DISQUALIFIED"

            DeterminerClassementReel = _
                "DSQ"

        Case Else

            DeterminerClassementReel = _
                "DNS"

    End Select

End Function

Public Function TrouverLignePilote( _
    ByVal ws As Worksheet, _
    ByVal steam As String) _
    As Long

    Dim cSteam As Long
    Dim r As Long
    Dim lastRow As Long

    cSteam = _
        GetColumnByHeader(ws, _
            "SteamID")

    lastRow = _
        ws.Cells(ws.Rows.Count, cSteam) _
          .End(xlUp).Row

    For r = 2 To lastRow

        If Trim( _
            ws.Cells(r, cSteam).Text) _
            = steam Then

            TrouverLignePilote = r
            Exit Function

        End If

    Next r

End Function

Public Sub SetValeur( _
    ByVal ws As Worksheet, _
    ByVal r As Long, _
    ByVal entete As String, _
    ByVal v As Variant)

    Dim c As Long

    c = _
        GetColumnByHeader(ws, _
            entete)

    If c > 0 Then
        ws.Cells(r, c).Value = v
    End If

End Sub

Public Function TrouverOngletManche( _
    ByVal trackID As String) As Worksheet

    Dim ws As Worksheet
    Dim wsManches As Worksheet
    Dim r As Long
    Dim cTrackID As Long
    Dim cNom As Long
    Dim nomOnglet As String

    Set wsManches = Sheets("Manches")

    cTrackID = GetColumnByHeader( _
                    wsManches, _
                    "TrackID AMS2")

    cNom = GetColumnByHeader( _
                    wsManches, _
                    "Nom circuit")

    If cTrackID = 0 _
    Or cNom = 0 Then Exit Function

    r = 2

    Do While _
        Trim(wsManches.Cells(r, cNom).Text) <> ""

        If CStr(wsManches.Cells(r, cTrackID).Value) _
            = trackID Then

            nomOnglet = _
                "Manche " & _
                wsManches.Cells(r, 1).Value & _
                " - " & _
                wsManches.Cells(r, cNom).Text

            If SheetExists(nomOnglet) Then
                Set TrouverOngletManche = _
                    Sheets(nomOnglet)
            End If

            Exit Function

        End If

        r = r + 1

    Loop

End Function
Private Sub ChercherSecteurs( _
    ByVal stage As Object, _
    ByVal participantID As Long, _
    ByVal bestLap As Double, _
    ByRef s1 As Double, _
    ByRef s2 As Double, _
    ByRef s3 As Double)

    Dim evt As Object
    Dim i As Long

    s1 = 0
    s2 = 0
    s3 = 0

    If bestLap <= 0 Then Exit Sub
    If Not stage.Exists("events") Then Exit Sub

    For i = 1 To stage("events").Count

        Set evt = stage("events")(i)

        If evt("event_name") = "Lap" Then

            If CLng(evt("participantid")) = participantID Then

                If ToDouble( _
                    evt("attributes")("LapTime")) _
                    = bestLap Then

                    s1 = ToDouble( _
                        evt("attributes") _
                        ("Sector1Time"))

                    s2 = ToDouble( _
                        evt("attributes") _
                        ("Sector2Time"))

                    s3 = ToDouble( _
                        evt("attributes") _
                        ("Sector3Time"))

                    Exit Sub

                End If

            End If

        End If

    Next i

End Sub
Private Function AjouterPiloteManche( _
    ByVal ws As Worksheet, _
    ByVal participants As Object, _
    ByVal participantID As Long) _
    As Long

    Dim p As Object

    Dim r As Long
    Dim steamID As String
    Dim nom As String

    Set p = _
        participants(CStr(participantID))

    steamID = _
        ClePilote( _
            participants, _
            participantID)

    nom = _
        CStr(p("Name"))

    '
    ' Recherche de la première ligne libre
    '
    r = 2

    Do While _
        Trim(CStr( _
            ws.Cells(r, 1).Value)) <> ""

        r = r + 1

    Loop

    ws.Cells(r, 1).Value = nom
    ws.Cells(r, 2).Value = steamID

    AjouterPiloteManche = r

End Function
