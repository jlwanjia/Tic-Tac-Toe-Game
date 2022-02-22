Public Class Form1
    Dim GameOn As Boolean = False
    Dim ImgNum As Integer
    Dim LastImgStr As String = ""
    Dim StepNum As Integer = 0
    Dim ImgX As Image = My.Resources.x
    Dim ImgO As Image = My.Resources.o
    Dim ImgPlain As Image = My.Resources.plain
    Dim GameAry() As String = {"", "", "", "", "", "", "", "", ""}

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox1.Image = ImgX
                LastImgStr = "x"
                GameAry(0) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox1.Image = ImgO
                LastImgStr = "o"
                GameAry(0) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 0)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox2.Image = ImgX
                LastImgStr = "x"
                GameAry(1) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox2.Image = ImgO
                LastImgStr = "o"
                GameAry(1) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 1)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox6.Image = ImgX
                LastImgStr = "x"
                GameAry(5) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox6.Image = ImgO
                LastImgStr = "o"
                GameAry(5) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 5)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox5.Image = ImgX
                LastImgStr = "x"
                GameAry(4) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox5.Image = ImgO
                LastImgStr = "o"
                GameAry(4) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 4)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox4.Image = ImgX
                LastImgStr = "x"
                GameAry(3) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox4.Image = ImgO
                LastImgStr = "o"
                GameAry(3) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 3)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox9.Image = ImgX
                LastImgStr = "x"
                GameAry(8) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox9.Image = ImgO
                LastImgStr = "o"
                GameAry(8) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 8)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox8.Image = ImgX
                LastImgStr = "x"
                GameAry(7) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox8.Image = ImgO
                LastImgStr = "o"
                GameAry(7) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 7)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox7.Image = ImgX
                LastImgStr = "x"
                GameAry(6) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox7.Image = ImgO
                LastImgStr = "o"
                GameAry(6) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 6)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub WinCheck()

        If (GameAry(0) = "x" And GameAry(1) = "x" AndAlso GameAry(2) = "x") Or
           (GameAry(3) = "x" And GameAry(4) = "x" AndAlso GameAry(5) = "x") Or
           (GameAry(6) = "x" And GameAry(7) = "x" AndAlso GameAry(8) = "x") Or
           (GameAry(0) = "x" And GameAry(5) = "x" AndAlso GameAry(8) = "x") Or
           (GameAry(1) = "x" And GameAry(4) = "x" AndAlso GameAry(7) = "x") Or
           (GameAry(2) = "x" And GameAry(3) = "x" AndAlso GameAry(6) = "x") Or
           (GameAry(0) = "x" And GameAry(4) = "x" AndAlso GameAry(6) = "x") Or
           (GameAry(2) = "x" And GameAry(4) = "x" AndAlso GameAry(8) = "x") Then
            Call Response("x win")
        ElseIf (GameAry(0) = "o" And GameAry(1) = "o" AndAlso GameAry(2) = "o") Or
           (GameAry(3) = "o" And GameAry(4) = "o" AndAlso GameAry(5) = "o") Or
           (GameAry(6) = "o" And GameAry(7) = "o" AndAlso GameAry(8) = "o") Or
           (GameAry(0) = "o" And GameAry(5) = "o" AndAlso GameAry(8) = "o") Or
           (GameAry(1) = "o" And GameAry(4) = "o" AndAlso GameAry(7) = "o") Or
           (GameAry(2) = "o" And GameAry(3) = "o" AndAlso GameAry(6) = "o") Or
           (GameAry(0) = "o" And GameAry(4) = "o" AndAlso GameAry(6) = "o") Or
           (GameAry(2) = "o" And GameAry(4) = "o" AndAlso GameAry(8) = "o") Then
            Call Response("o win")
        ElseIf StepNum = 9 Then
            Call Response("draw")
        End If
    End Sub

    Private Sub PictureBox3_Click(sender As Object, e As EventArgs) Handles PictureBox3.Click
        If GameOn = True Then
            If LastImgStr = "o" Then
                PictureBox3.Image = ImgX
                LastImgStr = "x"
                GameAry(2) = "x"
            ElseIf LastImgStr = "x" Then
                PictureBox3.Image = ImgO
                LastImgStr = "o"
                GameAry(2) = "o"
            End If
        Else
            Dim TempPB As PictureBox = CType(sender, PictureBox)
            Call CreateFirstImg(TempPB, 2)
        End If
        StepNum += 1
        If StepNum > 4 Then
            Call WinCheck()
        End If
    End Sub

    Private Sub CreateFirstImg(TempPB, Pointer)
        Randomize()
        ImgNum = CInt(Rnd())
        If ImgNum = 0 Then
            TempPB.Image = ImgO
            LastImgStr = "o"
            GameAry(Pointer) = "o"
        ElseIf ImgNum = 1 Then
            TempPB.Image = ImgX
            LastImgStr = "x"
            GameAry(Pointer) = "x"
        End If
        GameOn = True
    End Sub

    Private Sub Response(winner)
        Dim Response = MsgBox(winner & " " & "do you want to continue", vbYesNo)
        If Response = vbYes Then
            Dim PicBoxAry() As PictureBox = Me.Controls.OfType(Of PictureBox)().ToArray
            For i = 0 To 8
                PicBoxAry(i).Image = ImgPlain
                GameAry(i) = ""
            Next
            GameOn = False
            StepNum = 0
            LastImgStr = ""
            GameAry = {"", "", "", "", "", "", "", "", ""}
            Exit Sub
        Else
            End
        End If
    End Sub
End Class
