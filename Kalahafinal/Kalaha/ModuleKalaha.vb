Module ModuleKalaha
    Public tabjeton(14) As Single
    Public m As Single
    Public aupanier As Boolean
    Public vide As Boolean
    Public gagner As String

    Public Sub commencer()
        Dim i As Single
        i = 1
        For i = 1 To 14
            If (i = 7 Or i = 14) Then
                tabjeton(i) = 0
            Else
                tabjeton(i) = 3
            End If
        Next
    End Sub



    Public Sub ajouter(n As Single)
        Dim nbj As Single
        Dim i As Single
        nbj = tabjeton(n)
        tabjeton(n) = 0
        For i = 1 To nbj
            If n + i <= 14 Then
                tabjeton(n + i) = tabjeton(n + i) + 1
            Else
                tabjeton(n + i - 14) = tabjeton(n + i - 14) + 1
            End If
        Next

    End Sub

    Public Sub aupanierJ(ByVal n As Single, ByVal nbj As Single)
        aupanier = False
        If m = 1 Then
            If n + nbj = 7 Then
                aupanier = True
            End If
        ElseIf m = 2 Then
            If n + nbj = 14 Then
                aupanier = True
            End If
        End If
    End Sub

    Public Sub lignevide()
        Dim i As Single
        vide = True
        If m = 1 Then
            For i = 1 To 6
                If tabjeton(i) <> 0 Then
                    vide = False
                End If
            Next
        ElseIf m = 2 Then
            For i = 8 To 13
                If tabjeton(i) <> 0 Then
                    vide = False
                End If
            Next
        End If
    End Sub

    Public Sub endgame()
        gagner = ""
        If tabjeton(7) > 18 Then
            gagner = "Joueur1"
        ElseIf tabjeton(14) > 18 Then
            gagner = "Joueur2"
        ElseIf (tabjeton(7) = 18 And tabjeton(14) = 18) Then
            gagner = "Personne gagne"
        End If

    End Sub

End Module
