Private Sub CommandButton1_Click()

    Dim variable As Range
    Dim variable1 As Range
    Dim variable2 As Range
    Dim variable3 As Range
    Dim variable4 As Range
    Dim variable5 As Range
    Dim variable6 As Range
    Dim variable7 As Range
    Dim variable8 As Range
    Dim variable9 As Range
    Dim variable10 As Range
    Dim variable11 As Range
    Dim variable12 As Range
    Dim variable13 As Range
    Dim variable14 As Range
    Dim variable15 As Range
    Dim variable16 As Range
    
    Set variable = ActiveDocument.Bookmarks("fecha_inicio_contrato").Range
    variable = TextBox8.Value
    
    Set variable1 = ActiveDocument.Bookmarks("cc_arrendatario").Range
    variable1 = TextBox6.Value
    
    Set variable2 = ActiveDocument.Bookmarks("arrendatario").Range
    variable2 = TextBox5.Value
    
    Set variable3 = ActiveDocument.Bookmarks("celular_arrendatario").Range
    variable3 = TextBox22.Value
    
    Set variable4 = ActiveDocument.Bookmarks("cc_arrendador").Range
    variable4 = TextBox14.Value
    
    Set variable5 = ActiveDocument.Bookmarks("arrendador").Range
    variable5 = TextBox13.Value
    
    Set variable6 = ActiveDocument.Bookmarks("celular_arrendador").Range
    variable6 = TextBox21.Value
    
    Set variable7 = ActiveDocument.Bookmarks("dia_cada_mes").Range
    variable7 = TextBox23.Value
    
    Set variable8 = ActiveDocument.Bookmarks("personas").Range
    variable8 = TextBox7.Value
    
    Set variable9 = ActiveDocument.Bookmarks("fecha_firma_contrato").Range
    variable9 = TextBox9.Value
    
    Set variable10 = ActiveDocument.Bookmarks("CC_1").Range
    variable10 = TextBox18.Value
    
    Set variable11 = ActiveDocument.Bookmarks("NAME_1").Range
    variable11 = TextBox19.Value
    
    Set variable12 = ActiveDocument.Bookmarks("CEL_1").Range
    variable12 = TextBox20.Value
    
    Set variable13 = ActiveDocument.Bookmarks("CC_2").Range
    variable13 = TextBox10.Value
    
    Set variable14 = ActiveDocument.Bookmarks("NAME_2").Range
    variable14 = TextBox11.Value
    
    Set variable15 = ActiveDocument.Bookmarks("CEL_2").Range
    variable15 = TextBox12.Value
    
    Set variable16 = ActiveDocument.Bookmarks("numero_personas").Range
    variable16 = TextBox24.Value
    
    Unload Formulario
    
End Sub


