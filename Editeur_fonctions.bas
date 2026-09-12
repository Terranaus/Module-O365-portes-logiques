Attribute VB_Name = "Editeur_fonctions"
' ============================================================
' ÉDITEUR DE FONCTION LOGIQUE — version OMath corrigée
' ============================================================

Sub RedigerFonctionLogique(Optional control As IRibbonControl)
    Dim saisie As String
    Dim pos As Long
    
    saisie = InputBox( _
        "Entrez votre fonction logique." & vbCrLf & vbCrLf & _
        "Syntaxe :                     " & vbCrLf & _
        "           -X  =>  NON        " & vbCrLf & _
        "           +   =>  OU         " & vbCrLf & _
        "           *   =>  XOR        " & vbCrLf & _
        "           .   =>  ET         " & vbCrLf & _
        "           ( ) =>  Groupement ", "Éditeur de fonction logique")
    
    If saisie = "" Then Exit Sub
    
    saisie = Replace(saisie, " ", "")
    
    On Error GoTo Erreur
    
    ' Sortir de toute équation existante si le curseur y est piégé
    Do While Selection.OMaths.Count > 0
        Selection.OMaths(1).Range.Select
        Selection.Collapse Direction:=wdCollapseEnd
        Selection.MoveRight Unit:=wdCharacter, Count:=1
        If Selection.OMaths.Count > 0 Then Exit Do  ' sécurité anti-boucle
    Loop
    
    ' Vérifier qu'on n'est plus dans une OMath
    If Selection.OMaths.Count > 0 Then
        MsgBox "Impossible de sortir de l'équation existante. " & _
               "Placez le curseur ailleurs dans le document.", vbExclamation
        Exit Sub
    End If
    
    Dim rng As Range
    Set rng = Selection.Range
    Set rng = rng.OMaths.Add(rng)
    
    If rng.OMaths.Count = 0 Then
        MsgBox "Impossible de créer l'équation. " & _
               "Vérifiez que le curseur n'est pas dans une équation existante.", vbExclamation
        Exit Sub
    End If
    
    Dim omath As omath
    Set omath = rng.OMaths(1)
    
    pos = 1
    InsererDansEquation omath, saisie, pos
    
    If pos <= Len(saisie) Then
        MsgBox "Caractère inattendu en position " & pos & " : '" & _
               Mid(saisie, pos, 1) & "'", vbExclamation
        Exit Sub
    End If
    
    omath.BuildUp
    
    ' Sortir clairement de l'équation en insérant un espace après
    Dim afterRng As Range
    Set afterRng = omath.Range
    afterRng.Collapse Direction:=wdCollapseEnd
    afterRng.Select
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.TypeText " "
    
    Exit Sub
Erreur:
    MsgBox "Erreur : " & Err.Description & vbCrLf & _
           "Position : " & pos, vbExclamation
End Sub
' ============================================================
' PARSER — insère directement dans l'équation
' ============================================================
Sub InsererDansEquation(omath As omath, s As String, ByRef pos As Long)
    InsererEt omath, s, pos
    Do While pos <= Len(s) And Mid(s, pos, 1) = "+"
        pos = pos + 1
        AjouterTexte omath, "+"
        InsererEt omath, s, pos
    Loop
End Sub

Sub InsererEt(omath As omath, s As String, ByRef pos As Long)
    InsererXor omath, s, pos
    Do While pos <= Len(s) And Mid(s, pos, 1) = "."
        pos = pos + 1
        AjouterTexte omath, ChrW(&HB7)
        InsererXor omath, s, pos
    Loop
End Sub


Sub InsererXor(omath As omath, s As String, ByRef pos As Long)
    InsererNon omath, s, pos
    Do While pos <= Len(s) And Mid(s, pos, 1) = "*"
        pos = pos + 1
        AjouterTexte omath, ChrW(&H2295)
        InsererNon omath, s, pos
    Loop
End Sub

Sub InsererNon(omath As omath, s As String, ByRef pos As Long)
    If pos <= Len(s) And Mid(s, pos, 1) = "-" Then
        pos = pos + 1
        AjouterBarre omath, s, pos
    Else
        InsererAtome omath, s, pos
    End If
End Sub

Sub InsererAtome(omath As omath, s As String, ByRef pos As Long)
    Dim c As String, startPos As Long
    If pos > Len(s) Then Err.Raise 5, , "Expression incomplète"
    c = Mid(s, pos, 1)
    
    If c = "(" Then
        pos = pos + 1
        AjouterTexte omath, "("
        InsererDansEquation omath, s, pos
        If pos > Len(s) Or Mid(s, pos, 1) <> ")" Then
            Err.Raise 5, , "Parenthèse fermante manquante"
        End If
        pos = pos + 1
        AjouterTexte omath, ")"
    ElseIf c = ")" Then
        Err.Raise 5, , "Parenthèse fermante inattendue"
    Else
        startPos = pos
        Do While pos <= Len(s)
            c = Mid(s, pos, 1)
            If (c >= "a" And c <= "z") Or (c >= "A" And c <= "Z") Or _
               (c >= "0" And c <= "9") Or c = "_" Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        If pos = startPos Then
            Err.Raise 5, , "Caractère invalide : '" & Mid(s, pos, 1) & "'"
        End If
        AjouterTexte omath, Mid(s, startPos, pos - startPos)
    End If
End Sub

' ============================================================
' OUTILS D'INSERTION
' ============================================================
Sub AjouterTexte(omath As omath, texte As String)
    Dim rng As Range
    Set rng = omath.Range
    
    ' Détecter le placeholder d'équation vide (localisé FR ou EN)
    Dim txt As String
    txt = rng.Text
    
    If InStr(txt, "Tapez") > 0 Or InStr(txt, "Type equation") > 0 Then
        ' Premier ajout : remplacer le placeholder
        rng.Text = texte
    Else
        ' Ajouts suivants : insertion normale à la fin
        rng.Collapse Direction:=wdCollapseEnd
        rng.InsertBefore texte
    End If
End Sub

Sub AjouterBarre(omath As omath, s As String, ByRef pos As Long)
    Dim startPos As Long, endPos As Long
    Dim rngContenu As Range
    Dim funcBarre As OMathFunction
    Dim txt As String
    
    ' Détecter si le placeholder est encore présent
    txt = omath.Range.Text
    If InStr(txt, "Tapez") > 0 Or InStr(txt, "Type equation") > 0 Then
        ' Placeholder présent : on démarre depuis le VRAI début de l'équation
        ' (le placeholder sera remplacé par le contenu via AjouterTexte)
        startPos = omath.Range.Start
    Else
        ' Placeholder absent : on démarre à la fin du contenu existant
        startPos = omath.Range.End
    End If
    
    ' Insérer le contenu qui sera sous la barre
    InsererNon omath, s, pos
    
    endPos = omath.Range.End
    If endPos <= startPos Then Exit Sub
    
    Set rngContenu = ActiveDocument.Range(startPos, endPos)
    
    On Error Resume Next
    Set funcBarre = omath.Functions.Add(rngContenu, wdOMathFunctionBar)
    If Not funcBarre Is Nothing Then
        funcBarre.Bar.BarTop = True
    End If
    On Error GoTo 0
End Sub
