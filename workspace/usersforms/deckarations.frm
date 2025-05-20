VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} deckarations 
   Caption         =   "fenetre de déclaration"
   ClientHeight    =   10548
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9060.001
   OleObjectBlob   =   "deckarations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "deckarations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox2_Change()

End Sub
Private Sub ComboBox3_Change()

End Sub
Private Sub CommandButton1_Click()
    Load UserFormIndicateurs
    UserFormIndicateurs.Show
    ' enregistrer le score dans une variable
End Sub
Private Sub CommandButton2_Click()
    ' enregistrement de notre nouveau sinistre dans la base register
    Dim ligne As Integer
    Worksheets("options").Cells(2, 3) = Worksheets("options").Cells(2, 3) + 1
    ligne = Worksheets("options").Cells(2, 3)
    Worksheets("register").Cells(ligne, 1) = TextBox1.Value
    Worksheets("register").Cells(ligne, 2) = TextBox2.Value
    Worksheets("register").Cells(ligne, 3) = TextBox6.Value
    Worksheets("register").Cells(ligne, 4) = CDate(TextBox10.Value)
    Worksheets("register").Cells(ligne, 5) = ComboBox1.Value
    Worksheets("register").Cells(ligne, 6) = TextBox3.Value
    Worksheets("register").Cells(ligne, 7) = CDate(TextBox8.Value)
    Worksheets("register").Cells(ligne, 8) = CDate(TextBox9.Value)
    Worksheets("register").Cells(ligne, 9) = ComboBox2.Value
    
    ' format de la description
    Worksheets("register").Cells(ligne, 10) = ("Sinistre déclaré " + ComboBox1.Value + "survenu à" + TextBox3.Value + ".")
    Worksheets("register").Cells(ligne, 11) = TextBox12.Value
    Worksheets("register").Cells(ligne, 12) = ComboBox3.Value
    Worksheets("register").Cells(ligne, 13) = Format(TextBox4.Value, "#####.00")
    Worksheets("register").Cells(ligne, 14) = CDate(TextBox11.Value)
    Worksheets("register").Cells(ligne, 15) = TextBox7.Value
    Worksheets("register").Cells(ligne, 16) = 0
    menu.ListView2.ListItems.Add , , TextBox1.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , 0
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox2.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox6.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox10.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , ComboBox1.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox3.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox8.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox9.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , ComboBox2.Value
    
    ' format de la description
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , ("Sinistre déclaré " + ComboBox1.Value + "survenu à" + TextBox3.Value + ".")
    
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox12.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , ComboBox3.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox4.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox11.Value
    menu.ListView2.ListItems(ligne - 1).ListSubItems.Add , , TextBox9.Value
    Unload deckarations
End Sub
Private Sub Label11_Click()

End Sub

Private Sub SpinButton1_Change()

End Sub
Private Sub SpinButton1_SpinDown()
    Dim i As Single
    On Error Resume Next
    i = TextBox4.Value - 1
    TextBox4.Value = i
End Sub
Private Sub SpinButton1_SpinUp()
    Dim i As Single
    On Error Resume Next
    i = TextBox4.Value + 1
    TextBox4.Value = i
End Sub
Private Sub SpinButton2_Change()

End Sub
Private Sub SpinButton2_SpinDown()
    Dim i As Integer
    On Error Resume Next
    i = CInt(TextBox6.Value) - 1
    TextBox6.Value = i
End Sub
Private Sub SpinButton2_SpinUp()
    Dim i As Integer
    On Error Resume Next
    i = CInt(TextBox6.Value) + 1
    TextBox6.Value = i
End Sub
Private Sub SpinButton3_Change()

End Sub
Private Sub SpinButton3_SpinDown()
    Dim i As Single
    On Error Resume Next
    i = TextBox12.Value - 1
    TextBox12.Value = i
End Sub
Private Sub SpinButton3_SpinUp()
    Dim i As Single
    On Error Resume Next
    i = TextBox12.Value + 1
    TextBox12.Value = i
End Sub

Private Sub TextBox10_AfterUpdate()
    On Error Resume Next
    TextBox10 = CDate(TextBox10)
End Sub
Private Sub TextBox10_Change()

End Sub
Private Sub TextBox11_AfterUpdate()
    On Error Resume Next
    TextBox11 = CDate(TextBox11)
End Sub
Private Sub TextBox12_AfterUpdate()
    On Error Resume Next
    TextBox12.Value = Format(TextBox12.Value, "#,##0.00€")
End Sub
Private Sub TextBox12_Change()

End Sub

Private Sub TextBox4_AfterUpdate()
    On Error Resume Next
    TextBox4.Value = Format(TextBox4.Value, "#####.00")
End Sub
Private Sub TextBox4_Change()

End Sub
Private Sub TextBox6_Change()

End Sub
Private Sub TextBox8_AfterUpdate()
    On Error Resume Next
    TextBox8 = CDate(TextBox8)
End Sub

Private Sub TextBox8_Change()

End Sub
Private Sub TextBox9_AfterUpdate()
    On Error Resume Next
    TextBox9 = CDate(TextBox9)
End Sub

Private Sub TextBox9_Change()

End Sub

Private Sub UserForm_Click()

End Sub
