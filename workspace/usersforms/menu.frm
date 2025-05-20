VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} menu 
   Caption         =   "detection de la fraude"
   ClientHeight    =   10296
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   19452
   OleObjectBlob   =   "menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Load deckarations
    deckarations.Show
End Sub
Private Sub CommandButton2_Click()
    ' supprimer l'élément selectionné
    Dim i As Integer
    For i = 1 To Worksheets("options").Cells(2, 1)
        Worksheets("register").Cells(ListView2.SelectedItem.Index + 1, i) = ""
    Next i
    Worksheets("options").Cells(3, 13) = ListView2.SelectedItem.Index + 1
    ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
    CleanRegister
End Sub
Private Sub CommandButton4_Click()
    ' ajouter a la base de données
    Worksheets("options").Cells(2, 2) = Worksheets("options").Cells(2, 2) + 1
    Dim ligne As Integer
    ligne = Worksheets("options").Cells(2, 2)
    ' ListView2.SelectedItem
    Worksheets("Sheet1").Cells(ligne, 1) = Worksheets("register").Cells(ListView2.SelectedItem.Index + 1, 1)
    Worksheets("Sheet1").Cells(ligne, 2) = ListView2.SelectedItem.SubItems(2)
    Worksheets("Sheet1").Cells(ligne, 3) = ListView2.SelectedItem.SubItems(3)
    Worksheets("Sheet1").Cells(ligne, 13) = ListView2.SelectedItem.SubItems(14)
    Worksheets("Sheet1").Cells(ligne, 14) = ListView2.SelectedItem.SubItems(12)
    Worksheets("Sheet1").Cells(ligne, 12) = ListView2.SelectedItem.SubItems(13)
    
    Worksheets("Sheet1").Cells(ligne, 7) = CDate(ListView2.SelectedItem.SubItems(7))
    Worksheets("Sheet1").Cells(ligne, 8) = CDate(ListView2.SelectedItem.SubItems(8))
    Worksheets("Sheet1").Cells(ligne, 4) = CDate(ListView2.SelectedItem.SubItems(4))
    Worksheets("Sheet1").Cells(ligne, 14) = CDate(ListView2.SelectedItem.SubItems(14))
    
    Worksheets("Sheet1").Cells(ligne, 11) = ListView2.SelectedItem.SubItems(11)
    Worksheets("Sheet1").Cells(ligne, 9) = ListView2.SelectedItem.SubItems(9)
    Worksheets("Sheet1").Cells(ligne, 6) = ListView2.SelectedItem.SubItems(6)

    Worksheets("Sheet1").Cells(ligne, 5) = ListView2.SelectedItem.SubItems(5)
    Worksheets("Sheet1").Cells(ligne, 10) = ListView2.SelectedItem.SubItems(10)
    Worksheets("Sheet1").Cells(ligne, 15) = Worksheets("register").Cells(ListView2.SelectedItem.Index + 1, 15)

    ListView1.ListItems.Add , , Worksheets("register").Cells(ListView2.SelectedItem.Index + 1, 1)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , "Non"
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(2)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(3)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(4)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(5)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(6)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(7)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(8)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(9)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(10)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(11)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(12)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(13)
    ListView1.ListItems(ligne - 1).ListSubItems.Add , , ListView2.SelectedItem.SubItems(14)

    Dim i As Integer
    For i = 1 To Worksheets("options").Cells(2, 1)
        Worksheets("register").Cells(ListView2.SelectedItem.Index + 1, i) = ""
    Next i
    Worksheets("options").Cells(3, 13) = ListView2.SelectedItem.Index + 1
    ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
    CleanRegister
End Sub
Private Sub Label22_Click()

End Sub
Private Sub Label38_Click()

End Sub
Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)

End Sub
Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If ListView1.SelectedItem.SubItems(1) = "Oui" Then
        Label1.Caption = "Ce sinistre est une fraude."
    Else
        Label1.Caption = "Ce sinistre n'est pas une fraude."
    End If
    
    ' infos perso
    Label4.Caption = ListView1.SelectedItem.SubItems(2)
    Label6.Caption = ListView1.SelectedItem.SubItems(3)
    Label8.Caption = ListView1.SelectedItem.SubItems(14)
    Label10.Caption = ListView1.SelectedItem.SubItems(12)
    Label28.Caption = ListView1.SelectedItem.SubItems(13)
    
    ' dates
    Label13.Caption = ListView1.SelectedItem.SubItems(7)
    Label15.Caption = ListView1.SelectedItem.SubItems(8)
    Label17.Caption = ListView1.SelectedItem.SubItems(4)
    Label19.Caption = ListView1.SelectedItem.SubItems(14)
    
    ' sinistre
    Label22.Caption = Format(ListView1.SelectedItem.SubItems(11), "#,##0.00€")
    Label24.Caption = ListView1.SelectedItem.SubItems(9)
    Label26.Caption = ListView1.SelectedItem.SubItems(6)
    Label30.Caption = ListView1.SelectedItem.SubItems(7)
    Label32.Caption = ListView1.SelectedItem.SubItems(5)
    Label34.Caption = ListView1.SelectedItem.SubItems(10)
End Sub
Private Sub ListView2_BeforeLabelEdit(Cancel As Integer)

End Sub
Private Sub ListView2_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ' Affiche sur l'écran le contract selectionner
    Label36.Caption = Worksheets("register").Cells(ListView2.SelectedItem.Index + 1, 1)
    Label38.Caption = ListView2.SelectedItem.SubItems(2)
End Sub
Private Sub MultiPage1_Change()

End Sub
Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
    ' initialisation du menu
    Dim num_col As Integer
    Dim num_lin As Integer
    num_col = 14
    num_lin = 1
    
    ' initialisation de la premiere listview
    With Me.ListView1
        .ListItems.Clear
        .Gridlines = True
        With .ColumnHeaders
            .Clear
            .Add 1, , "id", 30
            .Add 2, , "caractere frauduleux", 60
            .Add 3, , "nom", 40
            .Add 4, , "age", 25
            .Add 5, , "date souscription", 50
            .Add 6, , "type sinistre", 50
            .Add 7, , "lieu", 50
            
            .Add 8, , "date sinistre", 50
            .Add 9, , "date déclaration", 60
            .Add 10, , "météo", 40
            .Add 11, , "déscription", 100
            .Add 12, , "cout", 25
            .Add 13, , "usage", 30
            .Add 14, , "crm", 25
            .Add 15, , "derniere consultation", 70
        End With
        While Worksheets("Sheet1").Cells(num_lin + 1, 1) <> ""
            ' commencer apres l'entete
            .ListItems.Add , , Worksheets("Sheet1").Cells(num_lin + 1, 1)
            .ListItems(num_lin).ListSubItems.Add , , Worksheets("Sheet1").Cells(num_lin + 1, 17)
            Dim j As Integer
            For j = 2 To num_col
                .ListItems(num_lin).ListSubItems.Add , , Worksheets("Sheet1").Cells(num_lin + 1, j)
            Next j
            num_lin = num_lin + 1
        Wend
    End With
    Worksheets("options").Cells(2, 1) = num_col
    Worksheets("options").Cells(2, 2) = num_lin
    
    
    
    ' initialisation de la deuixeme listview
    Dim num_lin_reg As Integer
    num_lin_reg = 1
    
    With Me.ListView2
        .ListItems.Clear
        .Gridlines = True
        With .ColumnHeaders
            .Clear
            .Add 1, , "id", 30
            .Add 2, , "score", 40
            .Add 3, , "nom", 40
            .Add 4, , "age", 25
            .Add 5, , "date souscription", 50
            .Add 6, , "type sinistre", 50
            .Add 7, , "lieu", 50
            
            .Add 8, , "date sinistre", 50
            .Add 9, , "date déclaration", 60
            .Add 10, , "météo", 40
            .Add 11, , "déscription", 100
            .Add 12, , "cout", 25
            .Add 13, , "usage", 30
            .Add 14, , "crm", 25
            .Add 15, , "derniere consultation", 70
        End With
        While Worksheets("register").Cells(num_lin_reg + 1, 1) <> ""
            .ListItems.Add , , Worksheets("register").Cells(num_lin_reg + 1, 1)
            .ListItems(num_lin_reg).ListSubItems.Add , , Worksheets("register").Cells(num_lin_reg + 1, 16)
            Dim v As Integer
            For v = 2 To num_col
                .ListItems(num_lin_reg).ListSubItems.Add , , Worksheets("register").Cells(num_lin_reg + 1, v)
            Next v
            num_lin_reg = num_lin_reg + 1
        Wend
    End With
    Worksheets("options").Cells(2, 3) = num_lin_reg
    
End Sub
