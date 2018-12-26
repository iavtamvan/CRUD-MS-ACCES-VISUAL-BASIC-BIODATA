Public Class Form1
    Public Databaru As Boolean
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Databaru = False
        isiGrid()
    End Sub

    Private Sub jalankansql(ByVal sQl As String)
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        Try
            objcmd.Connection = conn
            objcmd.CommandType = CommandType.Text
            objcmd.CommandText = sQl
            objcmd.ExecuteNonQuery()
            objcmd.Dispose()
            MsgBox("Data Sudah Disimpan", vbInformation)
        Catch ex As Exception
            MsgBox("Tidak Bisa Menyimpan data ke Database" & ex.Message)
        End Try
    End Sub
    Private Sub isiGrid()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM biodata", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "biodata")
        DataGridView1.DataSource = DS.Tables("biodata")
        DataGridView1.Enabled = True
        TextBox2.Focus()
        isiTextBox()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim simpan As String
        If TextBox1.Text = "" Then
            Databaru = True
        Else
            Databaru = False
        End If
        Me.Cursor = Cursors.WaitCursor

        If Databaru Then
            simpan = "INSERT INTO biodata(nama,nis,kelas,alamat) VALUES ('" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "') "
        Else
            simpan = "UPDATE biodata SET nama ='" & TextBox2.Text & "',nis ='" & TextBox3.Text & "',kelas ='" & TextBox4.Text & "',alamat ='" & TextBox5.Text & "' WHERE ID = " & TextBox1.Text & " "
        End If
        jalankansql(simpan)
        isiGrid()
        TextBox2.Focus()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub Bersih()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox2.Focus()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim hapussql As String
        Dim pesan As Integer
        pesan = MsgBox("Apakah anda yakin akan menghapus Data ID " + TextBox1.Text, vbExclamation + vbYesNo, "perhatian")
        If pesan = vbNo Then Exit Sub
        hapussql = "DELETE FROM biodata WHERE ID = " & TextBox1.Text & ""
        jalankansql(hapussql)
        Me.Cursor = Cursors.WaitCursor
        Bersih()
        TextBox2.Focus()
        Me.Cursor = Cursors.Default
    End Sub
    Private Sub isiTextBox()
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Databaru = False
        TextBox1.Text = DataGridView1.Item(0, i).Value
        TextBox2.Text = DataGridView1.Item(1, i).Value
        TextBox3.Text = DataGridView1.Item(2, i).Value
        TextBox4.Text = DataGridView1.Item(3, i).Value
        TextBox5.Text = DataGridView1.Item(4, i).Value
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        isiTextBox()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Bersih()
        Databaru = True
    End Sub

End Class
