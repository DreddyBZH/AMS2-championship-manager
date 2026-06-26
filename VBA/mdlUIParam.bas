Attribute VB_Name = "mdlUIParam"
Option Explicit

Public Type TDimension
    Largeur As Double
    Hauteur As Double
End Type

Public Type TPoint
    X As Double
    Y As Double
End Type

Public Type TPolice
    Taille As Double
    Gras As Boolean
End Type

Public Type TUI
    Resolution As TDimension
    Bandeau As TDimension
    Bouton As TDimension
    Tuile As TDimension
    Boite As TDimension
    Switch As TDimension
    Champ As TDimension
    PoliceTexte As TPolice
    PoliceTitre As TPolice
    PoliceValeur As TPolice
    RayonArrondi As Double
    NbBoutons As Long
    NbTuilesL As Long
    NbTuilesH As Long
    EcartHorizontal As Double
    EcartVertical As Double
    MargeHorizontale As Double
    MargeVerticale As Double
End Type

Public gUI As TUI

Public Sub InitialiserUI()

    With gUI

        .Resolution.Largeur = 3440
        .Resolution.Hauteur = 1440

        .Bandeau.Hauteur = 280

        .Bouton.Largeur = 450
        .Bouton.Hauteur = 45

        .Tuile.Largeur = 230
        .Tuile.Hauteur = 110

        .Boite.Largeur = 480
        .Boite.Hauteur = 240

        .Switch.Largeur = 80
        .Switch.Hauteur = 28

        .Champ.Largeur = 220
        .Champ.Hauteur = 28

        .PoliceTexte.Taille = 12
        .PoliceTexte.Gras = False

        .PoliceTitre.Taille = 18
        .PoliceTitre.Gras = True

        .PoliceValeur.Taille = 32
        .PoliceValeur.Gras = True

        .RayonArrondi = 0.06

        .NbBoutons = 5
        .NbTuilesL = 3
        .NbTuilesH = 2

    End With

    CalculerGeometrieUI

End Sub

Public Sub CalculerGeometrieUI()

    With gUI

        .EcartHorizontal = (.Resolution.Largeur - .Bouton.Largeur - (.NbTuilesL * .Tuile.Largeur) - .Boite.Largeur) / (.NbTuilesL + 4)
        .MargeHorizontale = .EcartHorizontal

        .EcartVertical = (.Resolution.Hauteur - .Bandeau.Hauteur - (.NbBoutons * .Bouton.Hauteur)) / (.NbBoutons + 1)
        .MargeVerticale = .EcartVertical

    End With

End Sub
