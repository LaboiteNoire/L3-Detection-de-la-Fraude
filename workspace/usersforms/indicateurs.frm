VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormIndicateurs 
   Caption         =   "UserFormIndicateurs"
   ClientHeight    =   4620
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7320
   OleObjectBlob   =   "indicateurs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormIndicateurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CommandValider_Click()
    'Traîter les informations : quels indicateurs ont été cochés

    Dim nomsIndicateurs(1 To 12) As String
    nomsIndicateurs(1) = "Âge incohérent"
    nomsIndicateurs(2) = "Dates incohérentes"
    nomsIndicateurs(3) = "Incohérence météo"
    nomsIndicateurs(4) = "Délai de déclaration"
    nomsIndicateurs(5) = "Sinistre avant souscription"
    nomsIndicateurs(6) = "Sinistre trop proche"
    nomsIndicateurs(7) = "Sinistre un dimanche"
    nomsIndicateurs(8) = "Coût anormal"
    nomsIndicateurs(9) = "CRM élevé"
    nomsIndicateurs(10) = "Consultation contrat récente"
    nomsIndicateurs(11) = "Incoherence Description"
    nomsIndicateurs(12) = "Sinistralité élevé"
    
    'rempli le tableau des choix des indicateurs avec true ou false
    ChoixIndicateurs(1) = UserFormIndicateurs.CheckAgeIncoherent.Value
    ChoixIndicateurs(2) = UserFormIndicateurs.CheckDatesIncoherentes.Value
    ChoixIndicateurs(3) = UserFormIndicateurs.CheckIncoherenceMeteo.Value
    ChoixIndicateurs(4) = UserFormIndicateurs.CheckDelaiDeclaration.Value
    ChoixIndicateurs(5) = UserFormIndicateurs.CheckSinistreAvantSouscription.Value
    ChoixIndicateurs(6) = UserFormIndicateurs.CheckSinistreTropProche.Value
    ChoixIndicateurs(7) = UserFormIndicateurs.CheckSinistreDimanche.Value
    ChoixIndicateurs(8) = UserFormIndicateurs.CheckCoutAnormal.Value
    ChoixIndicateurs(9) = UserFormIndicateurs.CheckCrmEleve.Value
    ChoixIndicateurs(10) = UserFormIndicateurs.CheckConsultationContrat.Value
    ChoixIndicateurs(11) = UserFormIndicateurs.CheckIncoherenceDescription
    ChoixIndicateurs(12) = UserFormIndicateurs.CheckSinistreParContrat
    
    nbChoix = 0
    For i = 1 To 12
        If ChoixIndicateurs(i) Then nbChoix = nbChoix + 1
    Next i
    
    ReDim Indicateur(1 To nbChoix)
    Dim j As Integer
    j = 1
    For i = 1 To 12
        If ChoixIndicateurs(i) Then
            Indicateur(j) = nomsIndicateurs(i) 'stockage du nom des indicateurs sélectionnés
            j = j + 1
        End If
    Next i

    
    ReDim ChoixPonderations(1 To nbChoix)
    Unload UserFormIndicateurs
    
             
End Sub

Private Sub UserForm_Click()

End Sub
