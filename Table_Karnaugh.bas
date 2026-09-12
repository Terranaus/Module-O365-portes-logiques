Attribute VB_Name = "Table_Karnaugh"
' ========================================
' TABLEAU DE KARNAUGH
' ========================================
Sub TableauKarnaugh(Optional control As IRibbonControl)
    Dim reponse As Variant
    Dim nbVar As Integer
    Dim nbLignes As Integer, nbColonnes As Integer
    Dim bitsLignes As Integer, bitsColonnes As Integer
    Dim varsLignes As String, varsColonnes As String
    Dim nomsVars() As String
    Dim saisieNoms As String
    Dim tbl As Table
    Dim i As Integer, j As Integer
    Dim grayVal As Integer
    
    ' --- Nombre de variables ---
    reponse = InputBox("Nombre de variables (2 à 10) :", _
                       "Tableau de Karnaugh", "")
    If reponse = "" Then Exit Sub
    If Not IsNumeric(reponse) Then
        MsgBox "Veuillez entrer un nombre.", vbExclamation
        Exit Sub
    End If
    
    nbVar = CInt(reponse)
    If nbVar < 2 Or nbVar > 10 Then
        MsgBox "Le nombre de variables doit être compris entre 2 et 10.", vbExclamation
        Exit Sub
    End If
    
    ' --- Noms des variables ---
    saisieNoms = InputBox("Entrez les noms des variables séparés par des virgules :", _
                          "Noms des variables", "")
    If saisieNoms = "" Then Exit Sub
    
    nomsVars = Split(saisieNoms, ",")
    For i = 0 To UBound(nomsVars)
        nomsVars(i) = Trim(nomsVars(i))
    Next i
    
    If UBound(nomsVars) + 1 < nbVar Then
        MsgBox "Vous devez fournir au moins " & nbVar & " noms de variables.", vbExclamation
        Exit Sub
    End If
    
    ' --- Dimensions (répartition : moitié en lignes, moitié en colonnes) ---
    ' Les bits de colonnes sont prioritaires (colonnes plus larges pour la lisibilité)
    bitsLignes = nbVar \ 2
    bitsColonnes = nbVar - bitsLignes
    
    nbLignes = 2 ^ bitsLignes
    nbColonnes = 2 ^ bitsColonnes
    
    ' --- Construire les chaînes de noms pour les lignes et colonnes ---
    varsLignes = ""
    For i = 0 To bitsLignes - 1
        varsLignes = varsLignes & nomsVars(i)
    Next i
    
    varsColonnes = ""
    For i = bitsLignes To bitsLignes + bitsColonnes - 1
        varsColonnes = varsColonnes & nomsVars(i)
    Next i
    
    ' --- Création de la table ---
    Set tbl = ActiveDocument.Tables.Add(Range:=Selection.Range, _
        NumRows:=nbLignes + 1, NumColumns:=nbColonnes + 1)
    tbl.Borders.Enable = True
    
    ' --- Coin supérieur gauche ---
    tbl.cell(1, 1).Range.Text = varsLignes & "\" & varsColonnes
    
    ' --- En-têtes de colonnes (expressions OMath avec barres) ---
    Dim varsCol() As String
    Dim bitsCol() As Integer
    ReDim varsCol(bitsColonnes - 1)
    ReDim bitsCol(bitsColonnes - 1)
    
    For j = 0 To nbColonnes - 1
        grayVal = j Xor (j \ 2)
        For i = 0 To bitsColonnes - 1
            varsCol(i) = nomsVars(bitsLignes + i)
            bitsCol(i) = (grayVal \ (2 ^ (bitsColonnes - 1 - i))) Mod 2
        Next i
        SetCellEquation tbl.cell(1, j + 2), varsCol, bitsCol
    Next j
    
    ' --- En-têtes de lignes (expressions OMath avec barres) ---
    Dim varsLig() As String
    Dim bitsLig() As Integer
    ReDim varsLig(bitsLignes - 1)
    ReDim bitsLig(bitsLignes - 1)
    
    For i = 0 To nbLignes - 1
        grayVal = i Xor (i \ 2)
        For j = 0 To bitsLignes - 1
            varsLig(j) = nomsVars(j)
            bitsLig(j) = (grayVal \ (2 ^ (bitsLignes - 1 - j))) Mod 2
        Next j
        SetCellEquation tbl.cell(i + 2, 1), varsLig, bitsLig
    Next i
    
    ' --- Mise en forme (centrage uniquement, pas de gras) ---
    Dim c As cell
    For Each c In tbl.Range.Cells
        c.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        c.VerticalAlignment = wdCellAlignVerticalCenter
    Next c
    
    ' --- Positionner le curseur après le tableau ---
    Selection.EndKey Unit:=wdStory
End Sub

' ========================================
' REMPLIR UNE CELLULE AVEC UNE ÉQUATION OMath
' ========================================
Sub SetCellEquation(cell As cell, varNames() As String, bits() As Integer)
    Dim rng As Range
    Dim omath As omath
    Dim i As Integer
    Dim funcBarre As OMathFunction
    Dim endRng As Range
    
    ' Vider la cellule
    cell.Range.Delete
    
    ' Créer une équation vide dans la cellule
    Set rng = cell.Range
    Set rng = rng.OMaths.Add(rng)
    Set omath = rng.OMaths(1)
    
    ' Construire l'expression : variables séparées par ·
    For i = 0 To UBound(varNames)
        If i > 0 Then AjouterTexte omath, ChrW(&HB7)
        
        If bits(i) = 1 Then
            AjouterTexte omath, varNames(i)
        Else
            Set endRng = omath.Range
            endRng.Collapse Direction:=wdCollapseEnd
            
            On Error Resume Next
            Set funcBarre = omath.Functions.Add(endRng, wdOMathFunctionBar)
            If Not funcBarre Is Nothing Then
                funcBarre.Bar.BarTop = True
                funcBarre.Bar.E.Range.Text = varNames(i)
            End If
            On Error GoTo 0
        End If
    Next i
    
    omath.BuildUp
End Sub
