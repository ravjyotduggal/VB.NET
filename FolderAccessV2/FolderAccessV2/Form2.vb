Public Class Form2
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Label2.Text = "Please wait while we are loading the access control page..."
        Form1.Show()
        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click



    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label2.Text = "User Login Details: " & Format(Now, "dd-MM-yyyy hh:mm:ss")

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim i As Integer

        i = MsgBox(" Are you sure you want to exit?", vbYesNo + vbInformation, "Exit Application")

        If i = 6 Then
            Me.Close()
        Else
            Exit Sub
        End If

    End Sub
End Class