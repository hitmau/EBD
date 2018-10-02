Imports System.Data.Common
Imports System.Data.SqlClient
Imports MySql.Data.MySqlClient

Public Class Form2
    'Conexão
    'Dim conn As New MySqlConnection
    Dim myCommand As New MySqlCommand
    'Dim myAdapter As New MySqlDataAdapter
    'Dim myData As New DataTable
    Dim SQL As String
    Dim sqlconection As MySqlConnection = New MySqlConnection
    Dim serverstring As String = "server=localhost;user id=root;password=;database=ebd"


    Dim _nome As String = ""
    Dim _Nascimento As String = ""
    Dim _telefone As String = ""
    Dim _sexo As String = ""
    Dim _classe As String = ""
    Dim _professor As Boolean = False
    Dim _especial As Boolean = False
    Dim _batismo As Boolean = False
    Dim _email As String = ""
    Dim _inativo As Boolean = False


    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sqlconection.ConnectionString = serverstring
        'Try
        '    If sqlconection.State = ConnectionState.Closed Then
        '        sqlconection.Open()
        '        MsgBox("Connection Successfully to DataBase")
        '    Else
        '        sqlconection.Close()
        '        MsgBox("Connection has closed")
        '    End If
        'Catch ex As Exception
        '    MsgBox(ex.ToString)

        'End Try

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox9.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        _sexo = "Homem" 'MetroComboBox1.SelectedIndex
        _classe = "teste" 'MetroComboBox2.SelectedItem
        MetroCheckBox1.Checked = False
        MetroCheckBox2.Checked = False
        MetroCheckBox3.Checked = False
        TextBox10.Text = ""
        MetroCheckBox4.Checked = False
    End Sub
End Class