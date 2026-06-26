Attribute VB_Name = "mdlUILayout"
Option Explicit

'=========================================================
' Calcul des positions de l'interface
' Dépend de mdlUIParam
'=========================================================

Public Function PositionBouton(ByVal Index As Long) As TPoint

    With PositionBouton

        .X = gUI.MargeHorizontale
        .Y = gUI.Bandeau.Hauteur _
           + gUI.EcartVertical _
           + (Index - 1) * (gUI.Bouton.Hauteur + gUI.EcartVertical)

    End With

End Function

Public Function PositionTuile(ByVal Colonne As Long, _
                              ByVal Ligne As Long) As TPoint

    With PositionTuile

        .X = gUI.MargeHorizontale _
           + gUI.Bouton.Largeur _
           + 2 * gUI.EcartHorizontal _
           + (Colonne - 1) * (gUI.Tuile.Largeur + gUI.EcartHorizontal)

        .Y = gUI.Bandeau.Hauteur _
           + gUI.EcartVertical _
           + (Ligne - 1) * (gUI.Tuile.Hauteur + gUI.EcartVertical)

    End With

End Function

Public Function PositionBoiteParametres() As TPoint

    With PositionBoiteParametres

        .X = gUI.MargeHorizontale _
           + gUI.Bouton.Largeur _
           + 2 * gUI.EcartHorizontal _
           + gUI.NbTuilesL * (gUI.Tuile.Largeur + gUI.EcartHorizontal)

        .Y = gUI.Bandeau.Hauteur + gUI.EcartVertical

    End With

End Function

Public Function PositionSwitch(ByVal Index As Long) As TPoint

    Dim P As TPoint

    P = PositionBoiteParametres

    With PositionSwitch

        .X = P.X + gUI.Boite.Largeur - gUI.Switch.Largeur - 20
        .Y = P.Y + 20 + (Index - 1) * 40

    End With

End Function

Public Function PositionChampTexte(ByVal Index As Long) As TPoint

    Dim P As TPoint

    P = PositionBoiteParametres

    With PositionChampTexte

        .X = P.X + gUI.Boite.Largeur - gUI.Champ.Largeur - 20
        .Y = P.Y + 20 + (Index - 1) * 40

    End With

End Function

Public Function RectangleTuile(ByVal Colonne As Long, _
                               ByVal Ligne As Long) As TPoint
    RectangleTuile = PositionTuile(Colonne, Ligne)
End Function

Public Function RectangleBouton(ByVal Index As Long) As TPoint
    RectangleBouton = PositionBouton(Index)
End Function
