Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1
    Private formClientes As clientes_cadastro = Nothing

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        If formClientes Is Nothing OrElse formClientes.IsDisposed Then
            formClientes = New clientes_cadastro()
            formClientes.Show()
        Else
            formClientes.BringToFront()
        End If
    End Sub
End Class
