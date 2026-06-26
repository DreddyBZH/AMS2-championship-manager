Attribute VB_Name = "mdlReglesChampionnat"
Public Function EstClassementNumerique( _
    ByVal classement As Variant) _
    As Boolean

    EstClassementNumerique = _
        IsNumeric( _
            Trim(CStr(classement)))

End Function



Public Function AttribuerPointsIA( _
    ByVal estIA As Boolean) _
    As Boolean

    If Not estIA Then

        AttribuerPointsIA = True
        Exit Function

    End If

    AttribuerPointsIA = _
        (UCase( _
            Trim( _
                GetConfig( _
                    "cfg_ClasserIA"))) = "OUI") _
        And _
        (UCase( _
            Trim( _
                GetConfig( _
                    "cfg_PointsIA"))) = "OUI")

End Function



