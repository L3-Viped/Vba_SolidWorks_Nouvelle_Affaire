Attribute VB_Name = "Nouvelle_Affaire"
Public Fin
Sub Nouvelle_Affaire()
    Dim New_N_Affaire
    Dim Copy_Folder


Afficher_UserForm_N蚊ffaire:
    Load N起Affaire
    N起Affaire.Show
    If Fin = 1 Then GoTo Fin:
    New_N_Affaire = N起Affaire.N起Affaire_Case.Value

Cr嶪r_Nom:
    Nom = N起Affaire.N起Affaire_Case.Value
    If N起Affaire.N起Real_Case.Value <> "" Then Nom = Nom & "-" & N起Affaire.N起Real_Case.Value
    If N起Affaire.Nbr_dIdrapal_Case.Value <> "" Then Nom = Nom & "-" & N起Affaire.Nbr_dIdrapal_Case.Value
    Nom = Nom & " IDRAPAL"
    If N起Affaire.Client_Case.Value <> "" Then Nom = Nom & "-" & N起Affaire.Client_Case.Value
    If N起Affaire.Client鉤Case.Value <> "" Then Nom = Nom & "-" & N起Affaire.Client鉤Case.Value
    If N起Affaire.Ville_Case.Value <> "" Then Nom = Nom & "-" & N起Affaire.Ville_Case.Value
    If N起Affaire.Pays_Case.Value <> "" Then Nom = Nom & "-" & N起Affaire.Pays_Case.Value

V廨ifier_Existance_Dossier:
    If Dir("\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\" & Nom, vbDirectory) <> "" Then
    Nom_Existant = MsgBox("Le dossier " & Nom & " existe d嶴a." & Chr(10) & "Voulez-vous recommencer ?", vbRetryCancel + vbCritical + vbDefaultButton1, "Nom Existant")
    If Nom_Existant = vbCancel Then GoTo Fin:
    If Nom_Existant = vbRetry Then GoTo Afficher_UserForm_N蚊ffaire:
    End If

Cr嶧tion_du_Dossier:
    Set Copy_Folder = CreateObject("Scripting.FileSystemObject")
    Copy_Folder.CopyFolder "\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\AF-IDRAPAL", "\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\" & Nom, False
    Shell Environ("WINDIR") & "\explorer.exe " & "\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\" & Nom & "\", vbNormalFocus

Fin:
    Unload N起Affaire
    Set Fin = Nothing
    End Sub


