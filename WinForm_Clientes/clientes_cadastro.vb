Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Messaging

Public Class clientes_cadastro
    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Me.Text = "Hello World!"
    End Sub

    Private Sub clientes_cadastro_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim sit As String() = {"ATIVA", "INATIVA"}
        Dim estados As String() = {"AC", "PA", "CE", "MA"}

        statusCb.Items.AddRange(sit)
        statusCb.SelectedIndex = 0

        ufCb.Items.AddRange(estados)
        ufCb.SelectedIndex = 1

        CarregaDadosDaPlanilha()

    End Sub

    Private Sub salvarBt_Click(sender As Object, e As EventArgs) Handles salvarBt.Click

        If String.IsNullOrWhiteSpace(nomeTx.Text) OrElse String.IsNullOrWhiteSpace(statusCb.Text) Then
            MessageBox.Show("Os campos 'Nome' e 'Situação' são obrigatórios!")
            Exit Sub
        End If

        cpfTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
        cepTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
        telefoneTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
        whatsAppTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals

        Dim pl As Excel.Worksheet = CType(Globals.ThisWorkbook.Application.Worksheets("Planilha1"), Excel.Worksheet)
        Dim ultLinha As Integer = pl.Cells(pl.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
        Dim proxId As Integer = 0

        If ultLinha > 1 Then
            Dim maxId As Integer = 0
            For i As Integer = 2 To ultLinha
                Dim idAtual As Integer
                If Integer.TryParse(pl.Cells(i, 1).Value.ToString(), idAtual) Then
                    If idAtual > maxId Then
                        maxId = idAtual
                    End If
                End If
            Next
            proxId = maxId + 1
        End If
        ultLinha += 1

        pl.Cells(ultLinha, 1).Value = proxId
        pl.Cells(ultLinha, 2).Value = nomeTx.Text
        pl.Cells(ultLinha, 3).Value = "'" & cpfTx.Text
        pl.Cells(ultLinha, 4).Value = cepTx.Text
        pl.Cells(ultLinha, 5).Value = ruaTx.Text
        pl.Cells(ultLinha, 6).Value = bairroTx.Text
        pl.Cells(ultLinha, 7).Value = numeroTx.Text
        pl.Cells(ultLinha, 8).Value = cidadeTx.Text
        pl.Cells(ultLinha, 9).Value = ufCb.Text
        pl.Cells(ultLinha, 10).Value = telefoneTx.Text
        pl.Cells(ultLinha, 11).Value = whatsAppTx.Text
        pl.Cells(ultLinha, 12).Value = emailTx.Text
        pl.Cells(ultLinha, 13).Value = statusCb.Text

        CarregaDadosDaPlanilha()
        LimapDados()
        MsgBox("Dados cadastrados com Sucesso!", vbInformation, "Cliente Cadastrado!")

    End Sub

    Private Sub CarregaDadosDaPlanilha()

        Dim tabela As New DataTable()

        DataGridView1.DataSource = Nothing
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

        Dim pl As Excel.Worksheet = CType(Globals.ThisWorkbook.Application.Worksheets("Planilha1"), Excel.Worksheet)
        Dim intervalousado As Excel.Range = pl.UsedRange
        Dim linhas As Integer = intervalousado.Rows.Count
        Dim colunas As Integer = intervalousado.Columns.Count

        For i As Integer = 1 To colunas
            Dim colunaNome As String = "Coluna" & i
            If intervalousado.Cells(1, i).value2 IsNot Nothing Then
                colunaNome = intervalousado.Cells(1, i).value2.ToString()
            End If
            tabela.Columns.Add(colunaNome)
        Next

        For i As Integer = 2 To linhas
            Dim novaLinha As DataRow = tabela.NewRow()
            For j As Integer = 1 To colunas
                If intervalousado.Cells(i, j).value2 IsNot Nothing Then
                    novaLinha(j - 1) = intervalousado.Cells(i, j).value2
                End If
            Next
            tabela.Rows.Add(novaLinha)
        Next

        DataGridView1.DataSource = tabela

    End Sub

    Private Sub LimapDados()
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is TextBox And ctrl.Name <> "idTx" Then
                CType(ctrl, TextBox).Clear()
            ElseIf TypeOf ctrl Is ComboBox And ctrl.Name <> "statusCb" Then
                CType(ctrl, ComboBox).SelectedIndex = -1
            ElseIf TypeOf ctrl Is MaskedTextBox Then
                CType(ctrl, MaskedTextBox).Clear()
            End If
        Next
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            Dim linha As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            idTx.Text = linha.Cells("ID").Value.ToString()
            nomeTx.Text = linha.Cells("Nome").Value.ToString()
            cpfTx.Text = linha.Cells("CPF").Value.ToString()
            cepTx.Text = linha.Cells("CEP").Value.ToString()
            ruaTx.Text = linha.Cells("rua").Value.ToString()
            numeroTx.Text = linha.Cells("numero").Value.ToString()
            bairroTx.Text = linha.Cells("bairro").Value.ToString()
            cidadeTx.Text = linha.Cells("cidade").Value.ToString()
            ufCb.Text = linha.Cells("uf").Value.ToString()
            statusCb.Text = linha.Cells("situação").Value.ToString()

        End If
    End Sub

    Private Sub editarBt_Click(sender As Object, e As EventArgs) Handles editarBt.Click

        If String.IsNullOrWhiteSpace(nomeTx.Text) OrElse String.IsNullOrWhiteSpace(statusCb.Text) Then
            MessageBox.Show("Os campos 'Nome' e 'Situação' são obrigatórios!")
            Exit Sub
        End If

        cpfTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
        cepTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
        telefoneTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
        whatsAppTx.TextMaskFormat = MaskFormat.ExcludePromptAndLiterals

        Dim pl As Excel.Worksheet = CType(Globals.ThisWorkbook.Application.Worksheets("Planilha1"), Excel.Worksheet)
        Dim ultLinha As Integer = pl.Cells(pl.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row

        For i As Integer = 2 To ultLinha
            If pl.Cells(i, 1).value = idTx.Text Then
                pl.Cells(i, 2).Value = nomeTx.Text
                pl.Cells(i, 3).Value = cpfTx.Text
                pl.Cells(i, 4).Value = cepTx.Text
                pl.Cells(i, 5).Value = ruaTx.Text
                pl.Cells(i, 6).Value = bairroTx.Text
                pl.Cells(i, 7).Value = numeroTx.Text
                pl.Cells(i, 8).Value = cidadeTx.Text
                pl.Cells(i, 9).Value = ufCb.Text
                pl.Cells(i, 10).Value = telefoneTx.Text
                pl.Cells(i, 11).Value = whatsAppTx.Text
                pl.Cells(i, 12).Value = emailTx.Text
                pl.Cells(i, 13).Value = statusCb.Text

                CarregaDadosDaPlanilha()
                LimapDados()
                MsgBox("Dados alterados com Sucesso!", vbInformation, "Cliente Alterado!")
            End If
        Next


    End Sub

    Private Sub excluirBt_Click(sender As Object, e As EventArgs) Handles excluirBt.Click
        If String.IsNullOrWhiteSpace(idTx.Text) Then
            MessageBox.Show("Por favor, selecione um cliente para excluir.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim confirm As DialogResult = MessageBox.Show("Tem certeza que deseja excluir este cliente?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If confirm = DialogResult.Yes Then

            Dim pl As Excel.Worksheet = CType(Globals.ThisWorkbook.Application.Worksheets("Planilha1"), Excel.Worksheet)
            Dim ultLinha As Integer = pl.Cells(pl.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row

            For i As Integer = 2 To ultLinha
                If pl.Cells(i, 1).Value.ToString() = idTx.Text Then
                    pl.Rows(i).Delete()
                    Exit For
                End If
            Next

            CarregaDadosDaPlanilha()
            LimapDados()
            MsgBox("Cliente excluído com sucesso!", vbInformation, "Cliente Excluído!")
        End If
    End Sub

End Class
