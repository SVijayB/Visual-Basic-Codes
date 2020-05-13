Imports System.Windows.Forms

Public Class frmNameEntry

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

        frmGame.PlayerName = TextBox1.Text
        Me.Dispose()

    End Sub

    Private Sub frmNameEntry_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed

        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        frmGame.Focus()
        Me.Dispose()

    End Sub

    Private Sub frmNameEntry_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Show()
        TextBox1.Focus()

    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyCode = Keys.Enter Then
            frmGame.PlayerName = TextBox1.Text
            Me.Dispose()
        End If

    End Sub

End Class
