Attribute VB_Name = "Nouvelle_Affaire"
Public Fin
Sub Nouvelle_Affaire()
    Dim New_N_Affaire
    Dim Copy_Folder


Afficher_UserForm_N°Affaire:
    Load N°_Affaire
    N°_Affaire.Show
    If Fin = 1 Then GoTo Fin:
    New_N_Affaire = N°_Affaire.N°_Affaire_Case.Value

Créer_Nom:
    Nom = N°_Affaire.N°_Affaire_Case.Value
    If N°_Affaire.N°_Real_Case.Value <> "" Then Nom = Nom & "-" & N°_Affaire.N°_Real_Case.Value
    If N°_Affaire.Nbr_dIdrapal_Case.Value <> "" Then Nom = Nom & "-" & N°_Affaire.Nbr_dIdrapal_Case.Value
    Nom = Nom & " IDRAPAL"
    If N°_Affaire.Client_Case.Value <> "" Then Nom = Nom & "-" & N°_Affaire.Client_Case.Value
    If N°_Affaire.Client¹_Case.Value <> "" Then Nom = Nom & "-" & N°_Affaire.Client¹_Case.Value
    If N°_Affaire.Ville_Case.Value <> "" Then Nom = Nom & "-" & N°_Affaire.Ville_Case.Value
    If N°_Affaire.Pays_Case.Value <> "" Then Nom = Nom & "-" & N°_Affaire.Pays_Case.Value

Vérifier_Existance_Dossier:
    If Dir("\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\" & Nom, vbDirectory) <> "" Then
    Nom_Existant = MsgBox("Le dossier " & Nom & " existe déja." & Chr(10) & "Voulez-vous recommencer ?", vbRetryCancel + vbCritical + vbDefaultButton1, "Nom Existant")
    If Nom_Existant = vbCancel Then GoTo Fin:
    If Nom_Existant = vbRetry Then GoTo Afficher_UserForm_N°Affaire:
    End If

Création_du_Dossier:
    Set Copy_Folder = CreateObject("Scripting.FileSystemObject")
    Copy_Folder.CopyFolder "\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\AF-IDRAPAL", "\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\" & Nom, False 'Répertoire à changer
    Shell Environ("WINDIR") & "\explorer.exe " & "\\192.168.0.15\idra-service\Tiers\##IDRAPAL##\IDRAPAL\" & Nom & "\", vbNormalFocus

Fin:
    Unload N°_Affaire
    Set Fin = Nothing
    End Sub


