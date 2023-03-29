Public Class Kalaha

    Private Sub btnjouer_Click(sender As Object, e As EventArgs) Handles btnjouer.Click
        Dim m As Single
        m = 1
        Call commencer()
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        btn1.Enabled = True
        btn2.Enabled = True
        btn3.Enabled = True
        btn4.Enabled = True
        btn5.Enabled = True
        btn6.Enabled = True
        btn7.Enabled = False
        btn8.Enabled = False
        btn9.Enabled = False
        btn10.Enabled = False
        btn11.Enabled = False
        btn12.Enabled = False
        btn13.Enabled = False
        btn14.Enabled = False

    End Sub

    Private Sub Kalaha_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btn1_Click(sender As Object, e As EventArgs) Handles btn1.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        m = 1
        n = 1
        nb = CSng(btn1.Text)

        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 2
            Call lignevide()
            If vide = False Then
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub btn2_Click(sender As Object, e As EventArgs) Handles btn2.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        m = 1

        n = 2
        nb = CSng(btn2.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 2
            Call lignevide()
            If vide = False Then
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub btn3_Click(sender As Object, e As EventArgs) Handles btn3.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        m = 1

        n = 3
        nb = CSng(btn3.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 2
            Call lignevide()
            If vide = False Then
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub btn4_Click(sender As Object, e As EventArgs) Handles btn4.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        n = 4
        m = 1

        nb = CSng(btn4.Text)

        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 2
            Call lignevide()
            If vide = False Then
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub btn5_Click(sender As Object, e As EventArgs) Handles btn5.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        n = 5
        m = 1

        nb = CSng(btn5.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 2
            Call lignevide()
            If vide = False Then
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub btn6_Click(sender As Object, e As EventArgs) Handles btn6.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        n = 6
        m = 1

        nb = CSng(btn6.Text)

        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next
        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 2
            Call lignevide()
            If vide = False Then
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            Else
                btn1.Enabled = False
                btn2.Enabled = False
                btn3.Enabled = False
                btn4.Enabled = False
                btn5.Enabled = False
                btn6.Enabled = False
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        End If

    End Sub

    Private Sub btn8_Click(sender As Object, e As EventArgs) Handles btn8.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        n = 8
        m = 2

        nb = CSng(btn8.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 1
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            Else
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            ElseIf vide = True Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            End If
        End If
    End Sub

    Private Sub btn9_Click(sender As Object, e As EventArgs) Handles btn9.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        n = 9
        m = 2

        nb = CSng(btn9.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 1
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            Else
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            ElseIf vide = True Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            End If
        End If
    End Sub

    Private Sub btn10_Click(sender As Object, e As EventArgs) Handles btn10.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        n = 10
        m = 2


        nb = CSng(btn10.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 1
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            Else
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            ElseIf vide = True Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            End If
        End If
    End Sub

    Private Sub btn11_Click(sender As Object, e As EventArgs) Handles btn11.Click
        Dim n As Single
        Dim nb As Single
        Dim i As Single
        n = 11
        m = 2


        nb = CSng(btn11.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 1
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            Else
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            ElseIf vide = True Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            End If
        End If
    End Sub

    Private Sub btn12_Click(sender As Object, e As EventArgs) Handles btn12.Click
        Dim n As Single
        Dim nb As Single
        n = 12
        m = 2


        nb = CSng(btn12.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 1
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            Else
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            ElseIf vide = True Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            End If
        End If
    End Sub

    Private Sub btn13_Click(sender As Object, e As EventArgs) Handles btn13.Click
        Dim n As Single
        Dim nb As Single
        n = 13
        m = 2


        nb = CSng(btn13.Text)
        ajouter(n)
        For i = 1 To 14
            Me.Controls("btn" & i).Text = tabjeton(i)
        Next

        Call endgame()
        If gagner <> "" Then
            MsgBox(gagner)
            For i = 1 To 14
                Me.Controls("btn" & i).Enabled = False
            Next
        End If

        aupanierJ(n, nb)
        If aupanier = False Then
            m = 1
            Call lignevide()
            If vide = False Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            Else
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            End If
        ElseIf aupanier = True Then
            Call lignevide()
            If vide = False Then
                For i = 8 To 13
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
            ElseIf vide = True Then
                For i = 1 To 6
                    If tabjeton(i) <> 0 Then
                        Me.Controls("btn" & i).Enabled = True
                    Else
                        Me.Controls("btn" & i).Enabled = False
                    End If
                Next
                btn8.Enabled = False
                btn9.Enabled = False
                btn10.Enabled = False
                btn11.Enabled = False
                btn12.Enabled = False
                btn13.Enabled = False
            End If
        End If
    End Sub
End Class
