Imports System.Drawing.Printing
Imports MySql.Data.MySqlClient
Imports System.Data.OleDb
Imports System.ComponentModel
Imports System.Convert
Imports System.Data.SqlClient
Imports System.IO

Public Class Form1
    Dim anoAtual As Integer = CInt(Format(Date.Now, "yyyy"))
    Dim dt2 As DateTime = DateTime.Now.AddMonths(1)
    Dim maxDominical As Integer

    Dim dtMesAtual As Date
    Dim intMesAtual As Integer
    Dim dtNiverAluno As Date
    Dim intNiverAluno As String

    Dim feitasMenos As Integer = 0
    Dim feitasNoResumo As Integer = 0
    Dim dt As New DataTable
    Dim imagem1 As String = ""
    Dim imagem2 As String = ""

    Dim _obs As String
    Dim _dInicio As String
    Dim _dFim As String
    Dim abc As String
    Dim selecao As Boolean = True

    'Variáveis Resumo
    Dim _rIdClasse As String = 0
    Dim _rPresentes As Integer
    Dim _rAusentes As Integer
    Dim _rVisitantes As Integer
    Dim _rCLasseId As Integer
    Dim _rCLasse As String

    'Variáveis Resumo
    Dim _ttAlunos As String
    Dim _ttIdClasse As String = 0
    Dim _ttPresentes As Integer
    Dim _ttAusentes As Integer
    Dim _ttVisitantes As Integer
    Dim _ttCLasseId As Integer
    Dim _ttCLasse As String
    Dim _Ttfertas As Double = 0.00

    Dim _Talunos As Integer = 0
    Dim _Tpresentes As Integer = 0
    Dim _Tausentes As Integer = 0
    Dim _Tvisitantes As Integer = 0
    Dim _Tofertas As Double = 0.00
    'Dim _TPresentes As Integer

    Dim totalDeAlunos As Integer = "0"
    Dim asteristico As String = ""
    Dim indices As Integer
    Dim novo1 As Integer
    Dim ativa3 As Boolean = True
    Dim dataAtual As Date = Date.Now
    Dim id As Integer
    Dim indice As Integer
    Dim indice1 As Integer
    Dim ids As Integer

    Dim _classes As String = ""
    Dim _dataini As Date
    Dim _datafim As Date
    Dim _especial As Boolean = False
    Dim _inativo As Boolean = False

    Dim stridCLasse As String = ""
    Dim stridPrincipal As String = ""
    Dim _nome2 As String = ""
    Dim _Nascimento2 As String = ""
    Dim _telefone2 As String = ""
    Dim _sexo2 As String = ""
    Dim _classe2 As String = ""
    Dim _professor2 As Boolean = False
    Dim _especial2 As Boolean = False
    Dim _batismo2 As Boolean = False
    Dim _email2 As String = ""
    Dim _inativo2 As Boolean = False

    Dim _nome3 As String = ""
    Dim _Nascimento3 As String = ""
    Dim _telefone3 As String = ""
    Dim _sexo3 As String = ""
    Dim _classe3 As String = ""
    Dim _professor3 As Boolean = False
    Dim _especial3 As Boolean = False
    Dim _batismo3 As Boolean = False
    Dim _email3 As String = ""
    Dim _inativo3 As Boolean = False

    Dim distancia1 As Integer
    Dim distancia2 As Integer
    Dim distancia3 As Integer
    Dim distancia4 As Integer
    Dim distancia5 As Integer
    Dim distancia6 As Integer

    Dim distancia7 As Integer
    Dim distancia8 As Integer
    Dim distancia9 As Integer
    Dim distancia10 As Integer
    Dim distancia11 As Integer
    Dim distancia12 As Integer

    Dim start As Boolean = True
    Dim start2 As Boolean = True
    Dim n As Integer = 1
    Dim iii As Integer = 0
    Dim inttamanhofontnormal As Integer = 10
    Dim inttamanhofontnegrito As Integer = 10

    'macoratti
    Dim cmd As OleDbCommand
    Private paginaAtual As Integer = 1
    Private nome As String
    Private data As String
    Private tel As String
    Private prof As String
    Dim ofertas As String
    Dim ofertass As String

    Private nome2 As String
    Private data2 As String
    Private tel2 As String
    Private prof2 As String


    Private MyConnection As OleDbConnection
    Private Leitor As OleDbDataReader
    Private RelatorioTitulo As String
    Private WithEvents m_PrintDocument As PrintDocument


    'FONTES DA DANFE
    'Private Big As New Font("Times New Roman", 20, FontStyle.Bold)

    'Private Font12_B As New Font("Times New Roman", 12, FontStyle.Bold)
    'Private Font12 As New Font("Times New Roman", 12, FontStyle.Regular)
    'Private Font6 As New Font("Times New Roman", 6, FontStyle.Regular)
    'Private Font6_Courier As New Font("Courier New", 6, FontStyle.Regular)
    'Private Font6_B As New Font("Times New Roman", 6, FontStyle.Bold)
    'Private Font5 As New Font("Times New Roman", 5, FontStyle.Regular)
    'Private Font5_B As New Font("Courier New", 5, FontStyle.Bold)
    'Private Font12_S As New Font("Times New Roman", 12, FontStyle.Underline)
    'Private Font8 As New Font("Times New Roman", 8, FontStyle.Regular)
    'Private Font7 As New Font("Times New Roman", 7, FontStyle.Regular)
    'Private Font10 As New Font("Times New Roman", 10, FontStyle.Regular)
    'Private Font10_B As New Font("Times New Roman", 10, FontStyle.Bold)
    'Private Font10_S As New Font("Times New Roman", 10, FontStyle.Underline)
    'Private Font8_B As New Font("Times New Roman", 8, FontStyle.Bold)
    'Private FontArial7 As New Font("Arial", 7, FontStyle.Bold)

    'Private WithEvents meuDataGridView As New MetroFramework.Controls.MetroGrid
    'Private Painel As New MetroFramework.Controls.MetroPanel
    'Private WithEvents incluiNovaLinhaButton As New MetroFramework.Controls.MetroButton
    'Private WithEvents deletaLinhaButton As New MetroFramework.Controls.MetroButton
    'Private WithEvents pesquisaNoGrid As New MetroFramework.Controls.MetroTextBox
    'Private WithEvents pesquisa As New MetroFramework.Controls.MetroButton
    Dim i As Integer = 2
    Dim nomeClasse As String



    Dim sqlconection As MySqlConnection = New MySqlConnection
    Dim serverstring As String = "server=localhost;user id=root;password=;database=ebd"

    'Restaurando configurações

    Public Sub CargaConfiguracoes()
        mgTotal.Refresh()
        mgTotal.MultiSelect = False
        TextBox9.MaxLength = MetroTextBox4.Text

        'Dim conn As New MySqlConnection
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter

        Dim SQL As String

        SQL = "Select * FROM config where id = (select max(g.id) from config g)"

        Try

            sqlconection.Open()

            Try
                Dim status As String = ""
                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(dt)
                End With
                Try
                    PictureBox1.Image = Image.FromFile(dt.Rows(0)(1))
                    imagem1 = dt.Rows(0)(1)
                    imagem1 = imagem1.Replace("\\", "\")
                Catch ex As Exception
                    status = "Imagem 1 não encontrada"
                End Try


                Try
                    PictureBox2.Image = Image.FromFile(dt.Rows(0)(2))
                    imagem2 = dt.Rows(0)(2)
                    imagem2 = imagem2.Replace("\\", "\")
                Catch ex As Exception
                    status = status & ", Imagem 2 não encontrada"
                    MsgBox(status)
                End Try

                Try


                    MetroToggle1.CheckState = dt.Rows(0)(6).ToString
                    cbCor.SelectedItem = dt.Rows(0)(3).ToString
                    cbTamanhoFonte.SelectedItem = dt.Rows(0)(4).ToString
                    MetroTextBox4.Text = dt.Rows(0)(5).ToString
                    SplitContainer4.SplitterDistance = dt.Rows(0)(12)
                    SplitContainer2.SplitterDistance = dt.Rows(0)(13)
                    SplitContainer3.SplitterDistance = dt.Rows(0)(14)
                    SplitContainer1.SplitterDistance = dt.Rows(0)(15)
                    MetroTextBox24.Text = dt.Rows(0)(16)
                    If dt.Rows(0)(17) = "True" Then
                        mtAniversariantes.Checked = True
                    Else
                        mtAniversariantes.Checked = False
                    End If
                    MetroTextBox27.Text = dt.Rows(0)(18)
                    MetroTextBox14.Text = dt.Rows(0)(19)
                    MetroTextBox15.Text = dt.Rows(0)(20)
                    MetroTextBox16.Text = dt.Rows(0)(21)
                    MetroTextBox17.Text = dt.Rows(0)(22)
                    MetroTextBox18.Text = dt.Rows(0)(23)
                    MetroToggle3.CheckState = dt.Rows(0)(24)
                    MetroToggle4.CheckState = dt.Rows(0)(25)
                Catch ex As Exception

                End Try
                'Dim z As Integer = 0
                'Dim arrayd(22) As Integer
                'Dim siteName As String
                'Dim singleChar As Char
                'siteName = dt.Rows(0)(7)
                'For Each singleChar In siteName
                '    arrayd(z) = Convert.ToInt32(singleChar)
                '    z += 1
                'Next
                'z = 0
                'While z <= arrayd.Count - 1
                '    If arrayd(i) <> "," Then
                '        MetroTrackBar1.Value = singleChar.ToString
                '    End If
                '    MetroTrackBar2.Value = singleChar.ToString
                '        MetroTrackBar3.Value = singleChar.ToString
                '        MetroTrackBar4.Value = singleChar.ToString
                '        MetroTrackBar5.Value = singleChar.ToString
                '        MetroTrackBar6.Value = singleChar.ToString
                '        MetroTrackBar7.Value = singleChar.ToString
                '        MetroTrackBar8.Value = singleChar.ToString
                '        MetroTrackBar9.Value = singleChar.ToString
                '        MetroTrackBar10.Value = singleChar.ToString
                '        MetroTrackBar11.Value = singleChar.ToString
                '        MetroTrackBar12.Value = singleChar.ToString


                'End While
                sqlconection.Clone()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
    End Sub

    'Carrega alunos do banco na grid
    Public Sub CargaBancoAlunos()
        mgTotal.Refresh()
        mgTotal.MultiSelect = False
        TextBox9.MaxLength = MetroTextBox4.Text

        'Dim conn As New MySqlConnection
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        Dim myData As New DataTable
        Dim SQL As String

        SQL = "Select CONTADOR, ALUNO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO, CLASSESUBESTAO FROM total"

        Try

            sqlconection.Open()

            Try

                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(myData)
                End With
                mgTotal.DataSource = myData

                sqlconection.Clone()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
        Try
            Me.mgTotal.Sort(Me.mgTotal.Columns(0), ListSortDirection.Ascending)

        Catch ex As Exception

        End Try
        PintaLinas()
    End Sub
    Public Sub CargaBancoLog()
        gridLog.Refresh()
        gridLog.MultiSelect = False

        'Dim conn As New MySqlConnection
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        Dim myData As New DataTable
        Dim SQL As String

        SQL = "Select CONTADOR, TIPO, DATA, ALUNO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO FROM log"

        Try

            sqlconection.Open()

            Try

                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(myData)
                End With
                gridLog.DataSource = myData

                sqlconection.Clone()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
    End Sub

    'Carrega as classes do banco na grid
    Public Sub CargaBancoClasses()
        mgClasses.Refresh()

        mgClasses.MultiSelect = False

        'Dim conn As New MySqlConnection
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        Dim myData As New DataTable
        Dim SQL As String

        SQL = "select c.contador as id, c.classe as Classe, c.dataini as inicio, c.DATAFIM as fim, c.especial as Especial, c.inativo as Inativado, a.nome as Categoria FROM classes c left join categoria a on a.ID = c.IDCATEGORIA"

        Try
            sqlconection.Open()
            Try

                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(myData)
                End With
                mgClasses.DataSource = myData
                sqlconection.Close()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
        Try
            Me.mgClasses.Sort(Me.mgClasses.Columns(0), ListSortDirection.Ascending)

        Catch ex As Exception
        End Try
    End Sub

    Public Sub CargaBancoCategorias()
        mgCategoria.Refresh()

        mgCategoria.MultiSelect = False

        'Dim conn As New MySqlConnection
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        Dim myData As New DataTable
        Dim SQL As String
        SQL = "select * FROM categoria"
        Try
            sqlconection.Open()
            Try

                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(myData)
                End With
                mgCategoria.DataSource = myData
                sqlconection.Close()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
    End Sub

    Public Sub CargaBancoResumo()
        'MetroGrid2.Refresh()
        'MetroGrid2.MultiSelect = False
        'TextBox9.MaxLength = MetroTextBox4.Text

        'Dim conn As New MySqlConnection
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        Dim myData As New DataTable
        Dim SQL As String

        SQL = "select r.id as 'N', r.data as 'Dara', c.CLASSE as 'Classe', r.totalalunos as 'Total Al.', r.presentes as 'Total Pr.', r.ausentes as 'Total Au.', r.visitantes as 'Visitantes', r.ofertas as 'Ofertas'from resumos r left join classes c on r.id_classes = c.contador where r.data = '" & dtRelatorio.Text & "';"

        Try

            sqlconection.Open()

            Try

                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(myData)
                End With
                mgResumo.DataSource = myData

                sqlconection.Clone()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
        feitasMenos = feitasNoResumo - mgResumo.Rows.Count
        GroupBox4.Text = "Relatório da CLasses Dominical - faltam (" & feitasMenos & ") classes."

    End Sub

    Public Sub CargaBancoResumoD()
        'MetroGrid2.Refresh()
        'MetroGrid2.MultiSelect = False
        'TextBox9.MaxLength = MetroTextBox4.Text

        'Dim conn As New MySqlConnection
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        Dim myData As New DataTable
        Dim SQL As String

        SQL = "SELECT IDRESUMOD as ID, data as Data, TALUNOS AS Alunos_Matriculados, TPRESENTES as Presentes, TAUSENTES as Ausentes, TVISITANTES AS Visitantes, TOTAL as Total, TOFERTAS AS Ofertas FROM resumosdominical order by IDRESUMOD;"

        Try

            sqlconection.Open()

            Try

                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(myData)
                End With
                mgRelHistorico.DataSource = myData

                sqlconection.Clone()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
        Try
            Me.mgRelHistorico.Sort(Me.mgRelHistorico.Columns(0), ListSortDirection.Descending)

        Catch ex As Exception

        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


        sqlconection.ConnectionString = serverstring
        CarregaTudo()
        Dim Data_hoje As DateTime = DateTime.Now
        bcMes.Text = Format(Data_hoje, "MMMM")
        maxDominical = MetroTextBox27.Text

        For p As Integer = 0 To 119
            mcbIdadeIni.Items.Add(p)
            mcbIdadeFim.Items.Add(p)
        Next
        mcbIdadeIni.SelectedIndex = 0
        mcbIdadeFim.SelectedIndex = 0
        mcbIdadeIni.DropDownHeight = mcbIdadeIni.ItemHeight * 10
        mcbIdadeFim.DropDownHeight = mcbIdadeFim.ItemHeight * 10



        cbCor.Items.Add("Lime")
        cbCor.Items.Add("Silver")
        cbCor.Items.Add("Orange")
        cbCor.Items.Add("Blue")
        cbCor.Items.Add("Green")
        cbCor.Items.Add("Red")
        cbCor.Items.Add("Purple")
        cbCor.SelectedIndex = 0

        MetroComboBox1.Items.Add("Sexo")
        MetroComboBox1.Items.Add("Homem")
        MetroComboBox1.Items.Add("Mulher")
        MetroComboBox1.SelectedIndex = 0

        MetroComboBox3.Items.Add("Sexo")
        MetroComboBox3.Items.Add("Homem")
        MetroComboBox3.Items.Add("Mulher")
        MetroComboBox3.SelectedIndex = 0

        'Tamanho da fonte da impressão
        For i As Integer = 1 To 20
            cbTamanhoFonte.Items.Add(i)
        Next
        cbTamanhoFonte.SelectedIndex = 12

        'Ajuste da grade dos meses
        MetroTextBox5.Text = MetroTrackBar1.Value
        MetroTextBox6.Text = MetroTrackBar2.Value
        MetroTextBox7.Text = MetroTrackBar3.Value
        MetroTextBox8.Text = MetroTrackBar4.Value
        MetroTextBox9.Text = MetroTrackBar5.Value
        MetroTextBox10.Text = MetroTrackBar6.Value

        MetroTextBox13.Text = MetroTrackBar12.Value
        MetroTextBox12.Text = MetroTrackBar11.Value
        MetroTextBox11.Text = MetroTrackBar10.Value
        MetroTextBox3.Text = MetroTrackBar9.Value
        MetroTextBox2.Text = MetroTrackBar8.Value
        MetroTextBox1.Text = MetroTrackBar7.Value

    End Sub

    'Public Shared Function diaUtil(ByVal dt As DateTime) As DateTime


    'Botão que atualiza os bancos
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            CarregaTudo()
            MsgBox("As tabelas do Banco de dados foram carregadas com sucesso!")
        Catch ex As Exception
            Return
        End Try
    End Sub
    Private Sub CarregaTudo()
        'Try
        'Carrega Configurações salvas
        CargaConfiguracoes()
        'Carrega Categoria/Departamentos
        CargaBancoCategorias()
        'Carrega Alunos
        CargaBancoAlunos()
        'Carrega Classes
        CargaBancoClasses()
        'Carrega Log
        CargaBancoLog()
        'carrega as classes nas ComboBox
        AbasClasses()
        'Carrega Duplicados
        CarregaDuplicados()
        'Carrega Relatórios
        CargaBancoClassesResumo()
        'Carrega Resumo (Resumo Geral)
        CargaBancoResumo()
        'Carrega Resumo Dominical (Relatório Simplificado)
        CargaBancoResumoD()

        'Consulta tabelas do Banco
        consultasqlTabelas()
        'carrega a grid de total de ofertas
        TotalOfertas()

        MetroComboBox5.Items.Clear()
        For k As Integer = 0 To mgCategoria.Rows.Count - 1
            If mgCategoria.Rows(k).Cells(2).Value.ToString = "False" Then
                MetroComboBox5.Items.Add(mgCategoria.Rows(k).Cells(1).Value.ToString)
            Else
                MetroComboBox5.Items.Add(mgCategoria.Rows(k).Cells(1).Value.ToString & "*")
            End If
        Next
        Try
            MetroComboBox5.SelectedIndex = 0

        Catch ex As Exception

        End Try

        PintaLinasClasses()
        PintaLinasCategorias()
        'Catch ex As Exception
        '    MsgBox("Erro ao carregar alguma tabela do banco de dados:")
        '    Return
        'End Try
    End Sub
    Private Sub ChecaCheckBox1()
        Dim nomeFeito As String
        Dim nomeComparado As String

        For i As Integer = 0 To mgResumo.Rows.Count - 1
            nomeFeito = mgResumo.Rows(i).Cells(1).Value.ToString

            For m As Integer = 0 To CheckedListBox1.SelectedItems.Count - 1
                nomeComparado = CheckedListBox1.CheckedIndices(m).ToString
                If nomeFeito = nomeComparado Then
                    CheckedListBox1.SetItemCheckState(m, CheckState.Checked)
                End If
            Next
        Next
    End Sub
    Private Sub PintaLinas()
        For i As Integer = 0 To mgTotal.Rows.Count - 1
            If ((i Mod 2) = 0) Then
                mgTotal.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
                If mgTotal.Rows(i).Cells(10).Value.ToString = "True" Then
                    mgTotal.Rows(i).DefaultCellStyle.BackColor = Color.Red
                    mgTotal.Rows(i).DefaultCellStyle.ForeColor = Color.White
                End If
            ElseIf mgTotal.Rows(i).Cells(10).Value.ToString = "True" Then
                mgTotal.Rows(i).DefaultCellStyle.BackColor = Color.DarkRed
                mgTotal.Rows(i).DefaultCellStyle.ForeColor = Color.White
            End If

        Next
    End Sub
    Private Sub PintaLinasClasses()
        For i As Integer = 0 To mgClasses.Rows.Count - 1
            If ((i Mod 2) = 0) Then
                mgClasses.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
            If mgClasses.Rows(i).Cells(5).Value.ToString = "True" Then
                mgClasses.Rows(i).DefaultCellStyle.BackColor = Color.Red
                mgClasses.Rows(i).DefaultCellStyle.ForeColor = Color.White
            End If
        Next
    End Sub
    Private Sub PintaLinasCategorias()
        For i As Integer = 0 To mgCategoria.Rows.Count - 1
            If ((i Mod 2) = 0) Then
                mgCategoria.Rows(i).DefaultCellStyle.BackColor = Color.LightGray
            End If
        Next
    End Sub
    'Private Sub MetroGrid1_RowPrePaint(ByVal e As System.Windows.Forms.DataGridViewRowPrePaintEventArgs) 'Handles MetroGrid1.RowPrePaint
    '    If selecao = True Then
    '        If e.RowIndex >= 0 Then

    '            MetroGrid1.Rows(e.RowIndex).Cells(0).Value = e.RowIndex + 1
    '            If ((e.RowIndex Mod 2) = 0) Then
    '                MetroGrid1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGray
    '            ElseIf MetroGrid1.Rows(e.RowIndex).Cells(10).Value.ToString = "True" Then
    '                MetroGrid1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
    '                MetroGrid1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
    '            End If
    '        Else
    '            selecao = False
    '        End If
    '    End If
    'End Sub
    'Private Sub MetroGrid3_RowPrePaint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewRowPrePaintEventArgs) Handles MetroGrid3.RowPrePaint

    '    If e.RowIndex >= 0 Then
    '        MetroGrid3.Rows(e.RowIndex).Cells(0).Value = e.RowIndex + 1
    '        If ((e.RowIndex Mod 2) = 0) Then
    '            MetroGrid3.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightGray
    '        ElseIf MetroGrid1.Rows(e.RowIndex).Cells(5).Value.ToString = "True" Then
    '            MetroGrid3.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
    '            MetroGrid3.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.White
    '        End If
    '    End If
    'End Sub
    Private Sub AbasClasses()
        'Limpa as ComboBox
        MetroComboBox2.Items.Clear()
        ComboBox1.Items.Clear()
        MetroComboBox4.Items.Clear()
        'Carrega as BomboBox com as classes
        ComboBox1.Items.Add("Escolha uma Classe")
        For i As Integer = 0 To mgClasses.Rows.Count - 1
            Dim nome As String = mgClasses.Rows(i).Cells(1).Value.ToString
            If mgClasses.Rows(i).Cells(5).Value.ToString = "True" Then
                MetroComboBox2.Items.Add(nome & " (Inativo)")
                ComboBox1.Items.Add(nome & " (Inativo)")
                MetroComboBox4.Items.Add(nome & " (Inativo)")
            ElseIf mgClasses.Rows(i).Cells(4).Value.ToString = "True" Then
                MetroComboBox2.Items.Add(nome & "*")
                ComboBox1.Items.Add(nome & "*")
                MetroComboBox4.Items.Add(nome & "*")
            Else
                MetroComboBox2.Items.Add(nome)
                ComboBox1.Items.Add(nome)
                MetroComboBox4.Items.Add(nome)
            End If
        Next
        'Seleciona os primeiros indices no combobox caso exista.
        If mgClasses.Rows.Count - 1 > 0 Then
            MetroComboBox2.SelectedIndex = 0
            ComboBox1.SelectedIndex = 0
            MetroComboBox4.SelectedIndex = 0
        End If
        If MetroToggle1.Checked = True Then
            SplitContainer1.Orientation = Orientation.Horizontal
        Else
            SplitContainer1.Orientation = Orientation.Vertical
        End If
    End Sub
    'Consultas sql
    Private Sub consultasql()
        Dim conn As New MySqlConnection
        Dim myCommandC As New MySqlCommand
        Dim myAdapterC As New MySqlDataAdapter
        Dim myDataC As New DataTable
        Dim SQLC As String

        conn = New MySqlConnection
        'conn.ConnectionString = "Server=mysql.hostinger.com.br;Database=u918624441_banco;Uid=u918624441_root;Pwd=fx74com.;"
        'conn.ConnectionString = "server=mysql.hostinger.com.br;user id=u918624441_root;password=fx74com.;database=u918624441_banco"

        conn.ConnectionString = "server=localhost;user id=root;password=;database=ebd"

        SQLC = RichTextBox1.Text.Replace("\r\n", " ")

        Try
            conn.Open()
            Try
                myCommandC.Connection = conn
                myCommandC.CommandText = SQLC.Trim()
                myAdapterC.SelectCommand = myCommandC
                myAdapterC.Fill(myDataC)
                MetroGrid6.DataSource = myDataC
                conn.Close()
            Catch ex As Exception
                MsgBox("Erro")
            End Try
        Catch ex As Exception

        End Try
    End Sub

    'Consulta tabelas do bando de dados
    Private Sub consultasqlTabelas()
        Dim conn As New MySqlConnection
        Dim myCommandC As New MySqlCommand
        Dim myAdapterC As New MySqlDataAdapter
        Dim myDataC As New DataTable
        Dim SQLC As String

        conn = New MySqlConnection
        'conn.ConnectionString = "Server=mysql.hostinger.com.br;Database=u918624441_banco;Uid=u918624441_root;Pwd=fx74com.;"
        'conn.ConnectionString = "server=mysql.hostinger.com.br;user id=u918624441_root;password=fx74com.;database=u918624441_banco"

        conn.ConnectionString = "server=localhost;user id=root;password=;database=ebd"

        SQLC = "show tables from ebd;"

        Try
            conn.Open()
            Try
                myCommandC.Connection = conn
                myCommandC.CommandText = SQLC.Trim()
                myAdapterC.SelectCommand = myCommandC
                myAdapterC.Fill(myDataC)
                mgTabelas.DataSource = myDataC
                conn.Close()
            Catch ex As Exception
                MsgBox("Erro")
            End Try
        Catch ex As Exception

        End Try
    End Sub
    'Private Sub incluiNovaLinhaButton_Click(ByVal sender As Object, ByVal e As EventArgs) Handles incluiNovaLinhaButton.Click
    '    'inclui uma nova linha no grid
    '    Me.meuDataGridView.Rows.Add()
    'End Sub

    'Public Shared Sub Main()
    '    Application.EnableVisualStyles()
    '    Application.Run(New Form1())
    'End Sub
    'Private Sub defineLayout()

    '    'define o tamanho do painel e inclui os botões : deletar e incluir linha
    '    'Me.Size = New Size(450, 250)

    '    With incluiNovaLinhaButton
    '        .Text = "Inclui Linha"
    '        .Location = New Point(10, 10)
    '    End With

    '    With deletaLinhaButton
    '        .Text = "Deleta Linha"
    '        .Location = New Point(100, 10)
    '    End With

    '    With pesquisaNoGrid
    '        .Style = MetroFramework.MetroColorStyle.Black
    '        .Width = 150
    '        .Location = New Point(200, 10)
    '    End With

    '    With pesquisa
    '        .Text = "Procura"
    '        .Location = New Point(350, 10)
    '    End With

    '    With Painel
    '        .BackColor = Color.Aqua
    '        .Controls.Add(incluiNovaLinhaButton)
    '        .Controls.Add(deletaLinhaButton)
    '        .Controls.Add(pesquisaNoGrid)
    '        .Controls.Add(pesquisa)
    '        .Height = 100
    '        .Dock = DockStyle.Bottom

    '    End With
    '    'um1.Controls.Add(Painel)
    '    'um1.Controls.Add(meuDataGridView)

    '    'Me.Controls.Add(Me.Painel)
    'End Sub
    'Private Sub configuraDataGridView()

    '    'um1.Controls.Add(meuDataGridView)
    '    meuDataGridView.ColumnCount = 3

    '    With meuDataGridView.ColumnHeadersDefaultCellStyle
    '        .BackColor = Color.Tomato
    '        .ForeColor = Color.White
    '        .Font = New Font(meuDataGridView.Font, FontStyle.Bold)
    '    End With

    '    'define o nome, tamanho , inclui colunas e linha no gridview
    '    With meuDataGridView
    '        .Name = "meuDataGridView"
    '        .Location = New Point(8, 8)
    '        .Size = New Size(250, 150)
    '        .AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCellsExceptHeaders
    '        .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
    '        .CellBorderStyle = DataGridViewCellBorderStyle.Single
    '        .GridColor = Color.Black
    '        .RowHeadersVisible = False

    '        'define 3 colunas : codigo, nome e nascimento
    '        .Columns(0).Name = "Codigo"
    '        .Columns(1).Name = "Nome"
    '        .Columns(2).Name = "Nascimento"
    '        .Columns(2).Width = 200
    '        .Columns(2).DefaultCellStyle.Font = New Font(Me.meuDataGridView.DefaultCellStyle.Font, FontStyle.Italic)
    '        .SelectionMode = DataGridViewSelectionMode.FullRowSelect
    '        .MultiSelect = False
    '        .Dock = DockStyle.Fill
    '    End With

    'End Sub

    Private Sub cbCor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCor.SelectedIndexChanged
        If cbCor.SelectedItem = "Lime" Then
            MetroTabControl1.Style = MetroFramework.MetroColorStyle.Lime
            tcSecundario.Style = MetroFramework.MetroColorStyle.Lime
            mgTotal.Style = MetroFramework.MetroColorStyle.Lime
            mgProfessores.Style = MetroFramework.MetroColorStyle.Lime
            mgAlunos.Style = MetroFramework.MetroColorStyle.Lime
        ElseIf cbCor.SelectedItem = "Silver" Then
            MetroTabControl1.Style = MetroFramework.MetroColorStyle.Silver
            tcSecundario.Style = MetroFramework.MetroColorStyle.Silver
            mgTotal.Style = MetroFramework.MetroColorStyle.Silver
            mgProfessores.Style = MetroFramework.MetroColorStyle.Silver
            mgAlunos.Style = MetroFramework.MetroColorStyle.Silver
        ElseIf cbCor.SelectedItem = "Orange" Then
            MetroTabControl1.Style = MetroFramework.MetroColorStyle.Orange
            tcSecundario.Style = MetroFramework.MetroColorStyle.Orange
            mgTotal.Style = MetroFramework.MetroColorStyle.Orange
            mgProfessores.Style = MetroFramework.MetroColorStyle.Orange
            mgAlunos.Style = MetroFramework.MetroColorStyle.Orange
        ElseIf cbCor.SelectedItem = "Blue" Then
            MetroTabControl1.Style = MetroFramework.MetroColorStyle.Blue
            tcSecundario.Style = MetroFramework.MetroColorStyle.Blue
            mgTotal.Style = MetroFramework.MetroColorStyle.Blue
            mgProfessores.Style = MetroFramework.MetroColorStyle.Blue
            mgAlunos.Style = MetroFramework.MetroColorStyle.Blue

        ElseIf cbCor.SelectedItem = "Green" Then
            MetroTabControl1.Style = MetroFramework.MetroColorStyle.Green
            tcSecundario.Style = MetroFramework.MetroColorStyle.Green
            mgTotal.Style = MetroFramework.MetroColorStyle.Green
            mgProfessores.Style = MetroFramework.MetroColorStyle.Green
            mgAlunos.Style = MetroFramework.MetroColorStyle.Green

        ElseIf cbCor.SelectedItem = "Red" Then
            MetroTabControl1.Style = MetroFramework.MetroColorStyle.Red
            tcSecundario.Style = MetroFramework.MetroColorStyle.Red
            mgTotal.Style = MetroFramework.MetroColorStyle.Red
            mgProfessores.Style = MetroFramework.MetroColorStyle.Red
            mgAlunos.Style = MetroFramework.MetroColorStyle.Red
        ElseIf cbCor.SelectedItem = "Purple" Then
            MetroTabControl1.Style = MetroFramework.MetroColorStyle.Purple
            tcSecundario.Style = MetroFramework.MetroColorStyle.Purple
            mgTotal.Style = MetroFramework.MetroColorStyle.Purple
            mgProfessores.Style = MetroFramework.MetroColorStyle.Purple
            mgAlunos.Style = MetroFramework.MetroColorStyle.Purple
        End If
        MetroTabControl1.Refresh()
        tcSecundario.Refresh()
    End Sub
    'Adicinar novo menu principal
    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        txtIdPrincipal.Text = ""
        If TextBox9.Text = "" Or MaskedTextBox1.Text = "" Or MetroComboBox2.Text = "" Then
            MsgBox("Campos em Branco!")
        Else
            'Dim convertedDate As Date
            Dim myCommand2 As MySqlCommand

            Try

                sqlconection.Open()
                _nome3 = TextBox9.Text
                _Nascimento3 = MaskedTextBox1.Text
                _telefone3 = MaskedTextBox2.Text
                _sexo3 = MetroComboBox1.SelectedItem
                _classe3 = MetroComboBox2.SelectedItem
                _professor3 = MetroCheckBox1.CheckState
                _especial3 = MetroCheckBox2.CheckState
                _batismo3 = MetroCheckBox3.CheckState
                _email3 = TextBox10.Text
                _inativo3 = MetroCheckBox4.CheckState
                _obs = ""
                'convertedDate = Convert.ToDateTime(_Nascimento3)

                '_Nascimento3 = CDate(_Nascimento3).ToString("yyyy-MM-dd")
                'Sugestão de classe certa
                Dim dataAtualPessoa As String = _Nascimento3
                _Nascimento3 = anoAtual - _Nascimento3.Remove(0, 6)
                Dim data_1 As Integer
                Dim data_2 As Integer
                Dim _sugestao As String = ""
                For x As Integer = 0 To mgClasses.Rows.Count - 1
                    data_1 = CInt(mgClasses.Rows(x).Cells(2).Value)
                    'data_1 = CDate(data_1).ToString("yyyy-MM-dd")
                    data_2 = CInt(mgClasses.Rows(x).Cells(3).Value)
                    'data_2 = CDate(data_2).ToString("yyyy-MM-dd")
                    If data_1 <= _Nascimento3 And _Nascimento3 <= data_2 Then
                        MetroLabel13.Text = mgClasses.Rows(x).Cells(1).Value.ToString
                        MetroLabel13.Refresh()
                        _sugestao = mgClasses.Rows(x).Cells(1).Value.ToString
                        Exit For
                    End If
                Next
                'Fim da sugestão
                'Adicionar _sugestao na QUERY
                _Nascimento3 = MaskedTextBox1.Text
                _Nascimento3 = CDate(_Nascimento3).ToString("yyyy-MM-dd")
                If (_especial3 = "False") And (_professor3 = "False") Then
                    _classe3 = _sugestao
                End If

                Dim codDuplo As String = ""
                For m As Integer = 0 To mgTotal.Rows.Count - 1
                    If (TextBox9.Text = mgTotal.Rows(m).Cells(1).Value.ToString) Then
                        codDuplo = codDuplo & " " & mgTotal.Rows(m).Cells(0).Value.ToString
                    End If
                Next
                If codDuplo <> "" Then
                    If MsgBox("Encontramos Alunos com nomes parecidos, deseja cadastrar? " & codDuplo, MsgBoxStyle.YesNo, "Inserindo Aluno:") = MsgBoxResult.No Then
                        sqlconection.Close()
                        LimpaPrincipal()
                        Return
                    End If
                End If
                Dim Sql3 As String = "INSERT INTO TOTAL (ALUNO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO, CLASSESUBESTAO, OBS) VALUES ('" _
                    & _nome3 & "', '" _
                    & _Nascimento3 & "', '" _
                    & _telefone3 & "', '" _
                    & _sexo3 & "', '" _
                    & _classe3 & "', '" _
                    & _professor3 & "', '" _
                    & _especial3 & "', '" _
                    & _batismo3 & "', '" _
                    & _email3 & "', '" _
                    & _inativo3 & "', '" _
                    & _sugestao & "', '" _
                    & _obs & "');"


                'Dim cmd As New MySqlCommand
                myCommand2 = New MySqlCommand(Sql3, sqlconection)
                With myCommand2
                    '.CommandText = SQL
                    .CommandType = CommandType.Text
                    '.Connection = sqlconection
                    .ExecuteNonQuery()
                End With
                MsgBox("Cadastrado com sucesso")

            Catch ex As Exception
                MsgBox("Erro : " & ex.Message)
            End Try
            sqlconection.Close()
            'Adiciona log
            AddLog()
            'Apaga os campos
            TextBox9.Text = ""
            MaskedTextBox1.Text = ""
            MaskedTextBox2.Text = ""
            MetroComboBox3.SelectedIndex = 0
            MetroComboBox4.SelectedIndex = 0
            MetroCheckBox1.Checked = False
            MetroCheckBox2.Checked = False
            MetroCheckBox3.Checked = False
            TextBox10.Text = ""
            MetroCheckBox4.Checked = False
            'Atualiza a grid para aparecer o novo dado
            CargaBancoAlunos()
            CargaBancoLog()
            CarregaDuplicados()
            PintaLinas()
            'Vai para a ultima linha onde se encontra o novo dado.
            mgTotal.CurrentCell = mgTotal.Rows(mgTotal.Rows.Count - 1).Cells(0)
            LimpaPrincipal()
        End If
    End Sub
    Private Sub AddLog()
        Dim convertedDate As Date
        Dim myCommand2 As MySqlCommand

        Try
            sqlconection.Open()
            _nome3 = "(" & txtIdPrincipal.Text & ") " & TextBox9.Text
            _Nascimento3 = MaskedTextBox1.Text
            _telefone3 = MaskedTextBox2.Text
            _sexo3 = MetroComboBox1.SelectedItem
            _classe3 = MetroComboBox2.SelectedItem
            _professor3 = MetroCheckBox1.CheckState
            _especial3 = MetroCheckBox2.CheckState
            _batismo3 = MetroCheckBox3.CheckState
            _email3 = TextBox10.Text
            _inativo3 = MetroCheckBox4.CheckState

            convertedDate = Convert.ToDateTime(_Nascimento3)

            _Nascimento3 = CDate(_Nascimento3).ToString("yyyy-MM-dd")

            Dim _tipo As String = "Novo Aluno"

            Dim Sql3 As String = "INSERT INTO LOG (ALUNO, TIPO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO, DATA) VALUES ('" _
            & _nome3 & "', '" _
            & _tipo & "', '" _
            & _Nascimento3 & "', '" _
            & _telefone3 & "', '" _
            & _sexo3 & "', '" _
            & _classe3 & "', '" _
            & _professor3 & "', '" _
            & _especial3 & "', '" _
            & _batismo3 & "', '" _
            & _email3 & "', '" _
            & _inativo3 & "', '" _
            & dataAtual & "');"

            'Dim cmd As New MySqlCommand
            myCommand2 = New MySqlCommand(Sql3, sqlconection)
            With myCommand2
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            'MsgBox("Log registrado com sucesso")
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
        sqlconection.Close()
    End Sub
    Private Sub AltLog()
        Dim convertedDate As Date
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim _tipo As String = "Altera Aluno"
        Try
            sqlconection.Open()
            _nome2 = "(" & txtIdPrincipal.Text & ") " & TextBox9.Text
            _Nascimento2 = MaskedTextBox1.Text
            _telefone2 = MaskedTextBox2.Text
            _sexo2 = MetroComboBox1.SelectedItem
            _classe2 = MetroComboBox2.SelectedItem
            _professor2 = MetroCheckBox1.CheckState
            _especial2 = MetroCheckBox2.CheckState
            _batismo2 = MetroCheckBox3.CheckState
            _email2 = TextBox10.Text
            _inativo2 = MetroCheckBox4.CheckState

            convertedDate = Convert.ToDateTime(_Nascimento2)

            _Nascimento2 = CDate(_Nascimento2).ToString("yyyy-MM-dd")

            SQL2 = "INSERT INTO LOG (ALUNO, TIPO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO, DATA) VALUES ('" _
            & _nome2 & "', '" _
            & _tipo & "', '" _
            & _Nascimento2 & "', '" _
            & _telefone2 & "', '" _
            & _sexo2 & "', '" _
            & _classe2 & "', '" _
            & _professor2 & "', '" _
            & _especial2 & "', '" _
            & _batismo2 & "', '" _
            & _email2 & "', '" _
            & _inativo2 & "', '" _
            & dataAtual & "');"


            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            'MsgBox("Cadastrado alterado com sucesso")
            sqlconection.Close()

        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
    End Sub
    Private Sub DelLog()
        Dim convertedDate As Date
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim _tipo As String = "Deleta Aluno"
        Try
            sqlconection.Open()
            _nome2 = "(" & txtIdPrincipal.Text & ") " & TextBox9.Text
            _Nascimento2 = MaskedTextBox1.Text
            _telefone2 = MaskedTextBox2.Text
            _sexo2 = MetroComboBox1.SelectedItem
            _classe2 = MetroComboBox2.SelectedItem
            _professor2 = MetroCheckBox1.CheckState
            _especial2 = MetroCheckBox2.CheckState
            _batismo2 = MetroCheckBox3.CheckState
            _email2 = TextBox10.Text
            _inativo2 = MetroCheckBox4.CheckState

            convertedDate = Convert.ToDateTime(_Nascimento2)

            _Nascimento2 = CDate(_Nascimento2).ToString("yyyy-MM-dd")

            SQL2 = "INSERT INTO LOG (ALUNO, TIPO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO, DATA) VALUES ('" _
            & _nome2 & "', '" _
            & _tipo & "', '" _
            & _Nascimento2 & "', '" _
            & _telefone2 & "', '" _
            & _sexo2 & "', '" _
            & _classe2 & "', '" _
            & _professor2 & "', '" _
            & _especial2 & "', '" _
            & _batismo2 & "', '" _
            & _email2 & "', '" _
            & _inativo2 & "', '" _
            & dataAtual & "');"


            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            sqlconection.Close()

        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
    End Sub
    Private Sub AddClasseLog()

    End Sub
    Private Sub AltClasseLog()
        Dim convertedDate As Date
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim _tipo As String = "Altera Classe"
        Try
            sqlconection.Open()
            _nome2 = TextBox11.Text
            _Nascimento2 = mcbIdadeIni.Text & " - " & mcbIdadeFim.Text
            _telefone2 = "" ''MaskedTextBox2.Text
            _sexo2 = "" 'MetroComboBox1.SelectedItem
            _classe2 = TextBox11.Text
            _professor2 = "" 'MetroCheckBox1.CheckState
            _especial2 = CheckBox1.CheckState
            _batismo2 = CheckBox2.CheckState
            _email2 = ""


            convertedDate = Convert.ToDateTime(_Nascimento2)

            _Nascimento2 = CDate(_Nascimento2).ToString("yyyy-MM-dd")

            SQL2 = "INSERT INTO LOG (ALUNO, TIPO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO, DATA) VALUES ('" _
            & _nome2 & "', '" _
            & _tipo & "', '" _
            & _Nascimento2 & "', '" _
            & _telefone2 & "', '" _
            & _sexo2 & "', '" _
            & _classe2 & "', '" _
            & _professor2 & "', '" _
            & _especial2 & "', '" _
            & _batismo2 & "', '" _
            & _email2 & "', '" _
            & _inativo2 & "', '" _
            & dataAtual & "');"


            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            sqlconection.Close()

        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
    End Sub
    Private Sub DelClasseLog()

    End Sub

    'gerar a pagina e imprimir
    Private Sub m_PrintDocument_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles m_PrintDocument.PrintPage

        Using caneta As New Pen(Color.Black, 20)
            e.Graphics.DrawRectangle(caneta, e.MarginBounds)
            caneta.DashStyle = Drawing2D.DashStyle.Dash
            caneta.Alignment = Drawing2D.PenAlignment.Outset
            e.Graphics.DrawRectangle(caneta, e.PageBounds)
        End Using

        '¡ndica que nao ha  mais paginas a serem impressas
        e.HasMorePages = False
    End Sub
    'Layout da(s) p gina(s) a imprimir
    Private Sub ClasseUm(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        'Armazena data atual
        dtMesAtual = Date.Now
        'Obtem apena o Mês atual da data Atual
        intMesAtual = Month(dtMesAtual)
        'Variaveis das linhas
        Dim LinhasPorPagina As Single = 0
        Dim PosicaoDaLinha As Single = 0
        Dim PosicaoDaLinha2 As Single = 0
        Dim LinhaAtual As Integer = 0

        'Variaveis das margens
        Dim MargemEsquerda As Single = e.MarginBounds.Left - 10
        Dim MargemSuperior As Single = e.MarginBounds.Top + 60
        Dim MargemSuperior2 As Single = e.MarginBounds.Top + 60
        Dim MargemDireita As Single = e.MarginBounds.Right + 80
        Dim MargemInferior As Single = e.MarginBounds.Bottom + 30
        Dim CanetaDaImpressora As Pen = New Pen(Color.Black, 1)
        'Dim codigo As Integer

        'Variaveis das fontes
        Dim FonteNegrito As Font
        Dim FonteTitulo As Font
        Dim FonteSubTitulo As Font
        Dim FonteRodape As Font
        Dim FonteNormal As Font
        Dim FonteNormalProf As Font
        Dim FonteNormaltel As Font
        Dim FonteNormaltel2 As Font
        Dim totalPaginas As Integer
        If ativa3 = True Then
            totalPaginas = (mgAlunos.Rows.Count - 1) + (mgProfessores.Rows.Count - 1)
            novo1 = totalPaginas
            ativa3 = False
        End If

        'define efeitos em fontes
        FonteNegrito = New Font("Arial", 10, FontStyle.Bold)
        FonteTitulo = New Font("Century Gothic", 20, FontStyle.Bold)
        FonteSubTitulo = New Font("Century Gothic", 12, FontStyle.Bold)
        FonteRodape = New Font("Arial", 10)
        FonteNormal = New Font("Arial", inttamanhofontnormal)
        FonteNormalProf = New Font("Arial", inttamanhofontnormal, FontStyle.Bold)
        FonteNormaltel = New Font("Arial", 7, FontStyle.Bold)
        FonteNormaltel2 = New Font("Arial", 7)

        'define valores para linha atual e para linha da impressao
        LinhaAtual = 0
        'Cabecalho
        e.Graphics.DrawLine(CanetaDaImpressora, 10, 10, MargemDireita, 10)

        'Imagem
        Try
            e.Graphics.DrawImage(Image.FromFile(imagem1), 20, 20)
            e.Graphics.DrawImage(Image.FromFile(imagem2), e.MarginBounds.Right - 140, 20)
            'e.Graphics.DrawString(RelatorioTitulo & System.DateTime.Today, FonteSubTitulo, Brushes.Black, MargemEsquerda + 250, 120, New StringFormat())
        Catch ex As Exception
        End Try
        'nome da Classe
        e.Graphics.DrawString(nomeClasse, FonteTitulo, Brushes.Black, distancia7, 100, New StringFormat())
        'Linha 2
        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, 130, MargemEsquerda + distancia6, 130)
        'campos a serem impressos: Codigo e Nome
        e.Graphics.DrawString("N.", FonteNegrito, Brushes.Black, MetroTrackBar12.Value + 5, 137, New StringFormat())
        e.Graphics.DrawString("Nome", FonteNegrito, Brushes.Black, MargemEsquerda - 20, 137, New StringFormat())
        e.Graphics.DrawString("Nascimento", FonteNegrito, Brushes.Black, MargemEsquerda + 300, 137, New StringFormat())
        e.Graphics.DrawString("Telefone", FonteNegrito, Brushes.Black, MargemEsquerda + 400, 137, New StringFormat())
        'Busca Mes em configurações
        e.Graphics.DrawString(bcMes.Text.Trim, FonteNegrito, Brushes.Black, MargemDireita - 122, 110, New StringFormat())
        'Busca dias de domingo em configurações
        e.Graphics.DrawString(MetroTextBox14.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia1 + 10, 137, New StringFormat())
        e.Graphics.DrawString(MetroTextBox15.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia2 + 10, 137, New StringFormat())
        e.Graphics.DrawString(MetroTextBox16.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia3 + 10, 137, New StringFormat())
        e.Graphics.DrawString(MetroTextBox17.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia4 + 10, 137, New StringFormat())
        e.Graphics.DrawString(MetroTextBox18.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia5 + 10, 137, New StringFormat())

        'Culunas do indice
        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, 130, MargemEsquerda + distancia1, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, 130, MargemEsquerda + distancia2, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, 130, MargemEsquerda + distancia3, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, 130, MargemEsquerda + distancia4, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, 130, MargemEsquerda + distancia5, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, 130, MargemEsquerda + distancia6, 160)

        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, 130, distancia7, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, distancia8, 130, distancia8, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, distancia9, 130, distancia9, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, distancia10, 130, distancia10, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, distancia11, 130, distancia11, 160)
        e.Graphics.DrawLine(CanetaDaImpressora, distancia12, 130, distancia12, 160)
        'Linha 3
        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, 1600, MargemEsquerda + distancia6, 1600)

        LinhasPorPagina = CInt(e.MarginBounds.Height / FonteNormalProf.GetHeight(e.Graphics) - 9) + 6

        '================================================================================================================
        '               Inicia carga na folha
        '================================================================================================================
        While ((LinhaAtual < LinhasPorPagina) AndAlso (iii <= mgAlunos.Rows.Count - 1))
            'obtem os valores da grid
            Try
                nome = mgAlunos.Rows(iii).Cells(1).Value.ToString
                'If nome.Length >= "25" Then
                '    nome = nome.Remove(nome.Length - MetroTextBox4.Text, MetroTextBox4.Text)
                'End If
                data = mgAlunos.Rows(iii).Cells(2).Value
                tel = mgAlunos.Rows(iii).Cells(3).Value
                'prof = MetroGrid5.Rows(iii).Cells(4).Value

            Catch ex As Exception
            End Try
            'inicia a impressao
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            PosicaoDaLinha2 = MargemSuperior2 + (LinhaAtual * FonteNormal.GetHeight)
            Try
                '----------------------------------------------------------------------
                'Testa se é professor e executa inserção                              |
                '----------------------------------------------------------------------
                If start = True Then
                    For i As Integer = 0 To mgProfessores.Rows.Count - 1
                        'Pega data do aluno e converte para "Date"
                        dtNiverAluno = Date.Parse(mgProfessores.Rows(i).Cells(2).Value)
                        'Converte "Data" do aluno para Inteiro(Somente o Mês)
                        intNiverAluno = Month(dtNiverAluno)

                        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                        PosicaoDaLinha2 = MargemSuperior2 + (LinhaAtual * FonteNormal.GetHeight)

                        nome2 = mgProfessores.Rows(i).Cells(1).Value
                        data2 = mgProfessores.Rows(i).Cells(2).Value
                        tel2 = mgProfessores.Rows(i).Cells(3).Value

                        'prof2 = gridProf2.Rows(i).Cells(4).Value

                        e.Graphics.DrawString(n.ToString(), FonteNormalProf, Brushes.Black, MetroTrackBar12.Value + 5, PosicaoDaLinha, New StringFormat())


                        'Compara os meses.
                        If intNiverAluno = intMesAtual Then
                            e.Graphics.DrawString(nome2.ToString() & " \o/", FonteNormalProf, Brushes.Black, MargemEsquerda - 20, PosicaoDaLinha, New StringFormat())
                        Else
                            e.Graphics.DrawString(nome2.ToString(), FonteNormalProf, Brushes.Black, MargemEsquerda - 20, PosicaoDaLinha, New StringFormat())
                        End If
                        e.Graphics.DrawString(data2.ToString, FonteNormalProf, Brushes.Black, MargemEsquerda + 300, PosicaoDaLinha, New StringFormat())

                        If tel2.Length > 10 Then
                            e.Graphics.DrawString(tel2.ToString, FonteNormaltel2, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
                        Else
                            e.Graphics.DrawString(tel2.ToString, FonteNormalProf, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
                        End If
                        'e.Graphics.DrawString(prof2.ToString, FonteNormalProf, Brushes.Black, MargemEsquerda + 580, PosicaoDaLinha, New StringFormat())
                        LinhaAtual += 1
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2)
                        'Linhas na vertical
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)

                        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, distancia7, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia8, PosicaoDaLinha2, distancia8, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia9, PosicaoDaLinha2, distancia9, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia10, PosicaoDaLinha2, distancia10, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia11, PosicaoDaLinha2, distancia11, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia12, PosicaoDaLinha2, distancia12, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)

                        n += 1
                        If i = mgProfessores.Rows.Count - 1 Then
                            start = False
                        End If
                    Next
                    'Ao terminar carga dos professores ele executa inserção dos alunos

                    'Pega data do aluno e converte para "Date"



                    PosicaoDaLinha2 = MargemSuperior2 + (LinhaAtual * FonteNormal.GetHeight)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2)
                    PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                End If
                e.Graphics.DrawString(n.ToString(), FonteNormal, Brushes.Black, MetroTrackBar12.Value + 5, PosicaoDaLinha, New StringFormat())
                'dtNiverAluno = Date.Parse(mgAlunos.Rows(iii).Cells(2).Value)
                'Converte "Data" do aluno para Inteiro(Somente o Mês)
                intNiverAluno = data.Remove(0, 3) 'Month(dtNiverAluno)
                intNiverAluno = intNiverAluno.Remove(2, 5)
                If asteristico = True Then
                    nome = nome.Replace("*", "")
                    'Compara os meses.
                    If (intNiverAluno = intMesAtual) And (mtAniversariantes.Checked = "True") Then
                        e.Graphics.DrawString(nome.ToString() & " \o/", FonteNormalProf, Brushes.Black, MargemEsquerda - 20, PosicaoDaLinha, New StringFormat())
                    Else
                        e.Graphics.DrawString(nome.ToString(), FonteNormalProf, Brushes.Black, MargemEsquerda - 20, PosicaoDaLinha, New StringFormat())
                    End If
                Else
                    'Compara os meses.
                    If (intNiverAluno = intMesAtual) And (mtAniversariantes.Checked = "True") Then
                        e.Graphics.DrawString(nome.ToString() & " \o/", FonteNormalProf, Brushes.Black, MargemEsquerda - 20, PosicaoDaLinha, New StringFormat())
                    Else
                        e.Graphics.DrawString(nome.ToString(), FonteNormalProf, Brushes.Black, MargemEsquerda - 20, PosicaoDaLinha, New StringFormat())
                    End If
                End If
                If data.ToString Like "*2016" Or data.ToString Like "2017" Then
                    e.Graphics.DrawString(data.ToString().ToString, FonteNormal, Brushes.Silver, MargemEsquerda + 300, PosicaoDaLinha, New StringFormat())
                Else
                    e.Graphics.DrawString(data.ToString().ToString, FonteNormal, Brushes.Black, MargemEsquerda + 300, PosicaoDaLinha, New StringFormat())
                End If

                If tel.Length > 10 Then
                    e.Graphics.DrawString(tel.ToString().ToString, FonteNormaltel2, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
                Else
                    e.Graphics.DrawString(tel.ToString().ToString, FonteNormal, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
                End If

                'e.Graphics.DrawString(prof.ToString, FonteNormal, Brushes.Black, MargemEsquerda + 580, PosicaoDaLinha, New StringFormat())
            Catch ex As Exception
                MsgBox("sfsdf", ex.Message)
            End Try
            LinhaAtual += 1

            'Insere linha dos alunos (DIas)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2)
            'Insere coluna para alunos
            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            'Colunas dos alunos (inicio)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, distancia7, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia8, PosicaoDaLinha2, distancia8, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia9, PosicaoDaLinha2, distancia9, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia10, PosicaoDaLinha2, distancia10, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia11, PosicaoDaLinha2, distancia11, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia12, PosicaoDaLinha2, distancia12, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)

            n += 1
            iii += 1
        End While
        '================================================================================================================
        '               Finaliza carga na folha
        '================================================================================================================
        If iii - 1 = mgAlunos.Rows.Count - 1 Then
            If start2 = True Then
                While (LinhaAtual + 3 < LinhasPorPagina)
                    PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                    PosicaoDaLinha2 = MargemSuperior2 + (LinhaAtual * FonteNormal.GetHeight)
                    'e.Graphics.DrawString(prof2.ToString, FonteNormalProf, Brushes.Black, MargemEsquerda + 580, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
                    'Linhas na vertical
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)

                    e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, distancia7, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia8, PosicaoDaLinha2, distancia8, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia9, PosicaoDaLinha2, distancia9, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia10, PosicaoDaLinha2, distancia10, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia11, PosicaoDaLinha2, distancia11, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia12, PosicaoDaLinha2, distancia12, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    LinhaAtual += 1
                    e.Graphics.DrawString(n.ToString(), FonteNormal, Brushes.Black, MetroTrackBar12.Value + 5, PosicaoDaLinha, New StringFormat())
                    n += 1
                End While
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

                'Parabens
                If (mtAniversariantes.Checked = "True") Then
                    e.Graphics.DrawString("\o/ Dê os parabéns ao(s) aniversariante(s) do mês!", FonteRodape, Brushes.Tomato, MargemEsquerda - 60, PosicaoDaLinha + 5, New StringFormat())
                End If


                e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
                start2 = False
                LinhaAtual += 2

                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                PosicaoDaLinha2 = MargemSuperior2 + (LinhaAtual * FonteNormal.GetHeight)
                e.Graphics.DrawString("Total de Presentes:", FonteNegrito, Brushes.Black, MargemEsquerda + distancia1 - 143, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                PosicaoDaLinha2 = MargemSuperior2 + (LinhaAtual * FonteNormal.GetHeight)

                e.Graphics.DrawString("Total de Visitantes:", FonteNegrito, Brushes.Black, MargemEsquerda + distancia1 - 143, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                PosicaoDaLinha2 = MargemSuperior2 + (LinhaAtual * FonteNormal.GetHeight)

                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
                e.Graphics.DrawString("Total das Ofertas:", FonteNegrito, Brushes.Black, MargemEsquerda + distancia1 - 134, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)

            End If
            n = 1
        End If
        'Rodape
        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, MargemInferior, MargemEsquerda + distancia6, MargemInferior)
        e.Graphics.DrawString(System.DateTime.Now.ToString("MMMM - yyyy"), FonteRodape, Brushes.Black, MargemEsquerda - 60, MargemInferior, New StringFormat())

        e.Graphics.DrawString("Página :" & paginaAtual, FonteRodape, Brushes.Black, MargemDireita - 70, MargemInferior, New StringFormat())

        novo1 = novo1 - LinhaAtual
        'verifica se continua imprimindo
        If (LinhaAtual >= LinhasPorPagina And novo1 > 0) Then
            'If (MetroGrid5.Rows.Count - 1 < LinhaAtual) Then
            e.HasMorePages = True
            paginaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
        Else
            start = True
            start2 = True
            ativa3 = True
            iii = 0
            e.HasMorePages = False
        End If
    End Sub

    Private ControlaImpressao As Integer = 0
    Private ContaProdutos As Integer = 0
    Private TotalFolha As Integer = 1
    Private ContaProdutos1 As Integer = 0
    Private Sub InicioImpressao()
        'ZERA VARIAVEIS DE CONTROLE DE IMPRESSÃO
        ControlaImpressao = 0
        ContaProdutos = 0
        TotalFolha = 1
    End Sub
    'Private Sub ImprimirDanfe(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
    '    'VERIFICA SE É PRIMEIRA FOLHA OU DEMAIS PARA FORMATAR A IMPRESSÃO
    '    If ControlaImpressao = 0 Then
    '        InserirCabecalho(e)
    '        'DesenharRetangulos(e)
    '        'InserirTextos(e)
    '        'Preencher_Dados_Danfe(e)
    '        'Imprimir_Produtos(e)
    '    Else
    '        InserirCabecalho(e)
    '        'Imprimir_Produtos(e)
    '        'CamposdemaisPaginas(e)
    '    End If
    'End Sub


    'Private Sub InserirCabecalho(ByVal Gra_Saida As System.Drawing.Printing.PrintPageEventArgs)

    '    Dim pen As New Pen(Brushes.Black, 0.2)
    '    Gra_Saida.Graphics.PageUnit = GraphicsUnit.Millimeter

    '    'RETANGULOS ENTRADA E SAÍDA E FRETE POR CONTA
    '    Gra_Saida.Graphics.DrawRectangle(pen, 157, 14, 5, 5)

    '    'CONJUNTO DE LINHAS LATERAIS ESQUERDAS
    '    'Gra_Saida.Graphics.DrawLine(pen, 2, 201, 2, 2)
    '    'Gra_Saida.Graphics.DrawLine(pen, 23, 201, 23, 2)
    '    'Gra_Saida.Graphics.DrawLine(pen, 12, 201, 12, 36)
    '    'Gra_Saida.Graphics.DrawLine(pen, 2, 36, 23, 36)
    '    'Gra_Saida.Graphics.DrawLine(pen, 12, 137, 23, 137)
    '    'Ultima Linha Direita
    '    'Gra_Saida.Graphics.DrawLine(pen, 287, 201, 287, 2)

    '    'CONJUNTO DE LINHAS HORIZONTAIS DADOS EMPRESA
    '    'Gra_Saida.Graphics.DrawLine(pen, 23, 33, 287, 33)
    '    'Gra_Saida.Graphics.DrawLine(pen, 23, 40, 287, 40)
    '    'Gra_Saida.Graphics.DrawLine(pen, 23, 47, 287, 47)

    '    'CONJUNTO DE LINHAS ESQUADRAR A PAGINA
    '    'Gra_Saida.Graphics.DrawLine(pen, 2, 2, 287, 2)
    '    'Gra_Saida.Graphics.DrawLine(pen, 2, 201, 287, 201)

    '    'CONJUNTO DE LINHAS VERTICAIS DADOS DA NOTA
    '    'Gra_Saida.Graphics.DrawLine(pen, 130, 2, 130, 33)
    '    'Gra_Saida.Graphics.DrawLine(pen, 166, 2, 166, 40)

    '    'INSCRIÇÃO ESTADUAL / INSCR. SUBST / CNPJ
    '    'Gra_Saida.Graphics.DrawLine(pen, 112, 40, 112, 47)
    '    'Gra_Saida.Graphics.DrawLine(pen, 201, 40, 201, 47)

    '    'CONJUNTO DE LINHAS HORIZONTAIS CODIGO DE BARRAS
    '    'Gra_Saida.Graphics.DrawLine(pen, 164, 15, 287, 15)
    '    'Gra_Saida.Graphics.DrawLine(pen, 164, 22, 287, 22)

    '    'IDENTIFICAÇÃO DO DOCUMENTO
    '    Try
    '        Gra_Saida.Graphics.DrawImage(Image.FromFile("Logo1.jpg"), 4, 4, 75, 25)
    '    Catch ex As Exception
    '    End Try
    '    'Gra_Saida.Graphics.DrawString("NOTA", Font12, Brushes.BlueViolet, 6.3, 2)
    '    'Gra_Saida.Graphics.DrawString("NOTA", Font12, Brushes.Black, 60.3, 2)
    '    ''------------------------------------------------------------primeiro left
    '    'Gra_Saida.Graphics.DrawString("NOTA", Font12, Brushes.Black, 60, 0)
    '    ''------------------------------------------------------------segundo top
    '    'Gra_Saida.Graphics.DrawString("FISCAL", Font12, Brushes.Black, 5.3, 6)
    '    'Gra_Saida.Graphics.DrawString("Nº", Font12_B, Brushes.Black, 10.3, 11)

    '    'Título da Classe
    '    Gra_Saida.Graphics.DrawString(nomeClasse, Big, Brushes.Black, 4, 55)
    '    Gra_Saida.Graphics.DrawString("N.", Font12_B, Brushes.Black, 4, 60)
    '    Gra_Saida.Graphics.DrawString("Aluno", Font12_B, Brushes.Black, 20, 60)

    '    Gra_Saida.Graphics.DrawString("Nascimento", Font12_B, Brushes.Black, 120, 60)
    '    'Gra_Saida.Graphics.DrawString("SÉRIE 1", Font12, Brushes.Black, 5, 24)
    '    Try
    '        Gra_Saida.Graphics.DrawImage(Image.FromFile("Logo.jpg"), 27, 4, 95, 28)
    '    Catch ex As Exception
    '    End Try

    '    Gra_Saida.Graphics.DrawString("Identificação do Emitente", Font10_B, Brushes.Black, 25, 2)
    '    'Gra_Saida.Graphics.DrawString("Identificação do Emitente", Font10_B, Brushes.Black, 25, 2)
    '    'Gra_Saida.Graphics.DrawString("DANFE", Font12_B, Brushes.Black, 140, 2)
    '    'Gra_Saida.Graphics.DrawString("DOCUMENTO AUXILIAR DA", Font7, Brushes.Black, 131.2, 7)
    '    'Gra_Saida.Graphics.DrawString("NOTA FISCAL ELETRÔNICA", Font7, Brushes.Black, 131.5, 10)
    '    'Gra_Saida.Graphics.DrawString("0 - ENTRADA", Font7, Brushes.Black, 134.5, 13.5)
    '    'Gra_Saida.Graphics.DrawString("1 - SAÍDA", Font7, Brushes.Black, 134.5, 17)
    '    ''Gra_Saida.Graphics.DrawString("N.º " & Int32.Parse(I_DDadosNfe.NUMERO_NFE).ToString("000000000"), Font10_B, Brushes.Black, 131.5, 21)
    '    'Gra_Saida.Graphics.DrawString("SÉRIE 1", Font10_B, Brushes.Black, 131.5, 25)

    '    ''CALCULA TOTAL DE FOLHAS
    '    'Dim Resto As Integer
    '    ''If V_PRODUTOS.Count > 3 Then
    '    ''    Resto = (V_PRODUTOS.Count - 3) Mod 8
    '    ''    If Resto > 0 Then
    '    ''        TotalFolha = 2 + ((V_PRODUTOS.Count - 3) - Resto) / 8
    '    ''    Else
    '    ''        TotalFolha = 1 + ((V_PRODUTOS.Count - 3) - Resto) / 8
    '    ''    End If
    '    ''End If

    '    Dim LimitePagina As Integer
    '    Dim AlturaLinha As Integer

    '    LimitePagina = 150
    '    AlturaLinha = 115

    '    Dim Conte As Integer = 1
    '    Dim ContaProd As Integer = 0
    '    ContaProdutos1 = 0

    '    Gra_Saida.Graphics.DrawString("FOLHA " & ControlaImpressao + 1 & "/" & Conte, Font10_B, Brushes.Black, 131.5, 29)

    '    'PRIMEIRA LINHA
    '    'Dim CodigoBarra As CodigodeBarra
    '    'CodigoBarra = New CodigodeBarra(CodigodeBarra.BCEncoding.Code128C)
    '    'Gra_Saida.Graphics.DrawImage(CodigoBarra.DrawBarCode, 184, 2, 100, 11)
    '    'Gra_Saida.Graphics.DrawString("CHAVE DE ACESSO", Font6, Brushes.Black, 167, 15.2)
    '    'Gra_Saida.Graphics.DrawString("Consulta de autenticidade no portal nacional da NF-e", Font12, Brushes.Black, 167, 22)
    '    'Gra_Saida.Graphics.DrawString("www.nfe.fazenda.gov.br/portal", Font12_S, Brushes.Black, 167, 26)
    '    'Gra_Saida.Graphics.DrawString("ou no site da Sefaz Autorizadora", Font12, Brushes.Black, 220.9, 26)

    '    'Gra_Saida.Graphics.DrawString("NATUREZA DA OPERAÇÃO", Font6, Brushes.Black, 25, 33.2)
    '    'Gra_Saida.Graphics.DrawString("PROTOCOLO DE AUTORIZAÇÃO DE USO", Font6, Brushes.Black, 167, 33.2)

    '    ''SEGUNDA LINHA
    '    'Gra_Saida.Graphics.DrawString("INSCRIÇÃO ESTADUAL", Font6, Brushes.Black, 25, 40.2)
    '    'Gra_Saida.Graphics.DrawString("INSCR. ESTADUAL DO SUBST. TRIBUT.", Font6, Brushes.Black, 113, 40.2)
    '    'Gra_Saida.Graphics.DrawString("CNPJ", Font6, Brushes.Black, 202, 40.2)

    '    'Dim NChave As String = ""
    '    'Dim ContarS As Integer = 0
    '    ''For x As Int16 = 0 To I_DDadosNfe.CHAVEACESSO_NFE.Length - 1
    '    ''    ContarS = ContarS + 1
    '    ''    NChave = NChave & I_DDadosNfe.CHAVEACESSO_NFE.Substring(x, 1)
    '    ''    If ContarS = 4 Then
    '    ''        ContarS = 0
    '    ''        NChave = NChave & " "
    '    ''    End If
    '    ''Next
    '    'Gra_Saida.Graphics.DrawString(NChave, Font8_B, Brushes.Black, 167, 18)

    '    'NATUREZA, PROTOCOLO E TIPO DE NOTA
    '    'Gra_Saida.Graphics.DrawString(I_DDadosNfe.TIPONOTA_NFE, Font10, Brushes.Black, 158, 14.5)
    '    'Gra_Saida.Graphics.DrawString(I_DDadosNfe.NATUREZA_NFE, Font10, Brushes.Black, 25, 36.2)
    '    'Gra_Saida.Graphics.DrawString(I_DDadosNfe.PROTOCOLO_NFE & "  " & I_DDadosNfe.DHRECBTO_NFE, Font10, Brushes.Black, 167, 36.2)

    '    'DADOS EMITENTE
    '    'Gra_Saida.Graphics.DrawString(V_DEmitente.NOME, Font12_B, Brushes.Black, 25, 5)
    '    'Gra_Saida.Graphics.DrawString(V_DEmitente.ENDERECO_COMPLETO, Font8_B, Brushes.Black, 70, 13)
    '    'Gra_Saida.Graphics.DrawString(V_DEmitente.MUNICIPIO & ", " & V_DEmitente.UF, Font8_B, Brushes.Black, 70, 16)
    '    'Gra_Saida.Graphics.DrawString("FONE: " & V_DEmitente.TELEFONE & " CEP " & V_DEmitente.CEP, Font8_B, Brushes.Black, 70, 19)

    '    'INICIO DO PREENCHIMENTO NOTA EMITENTE
    '    'Gra_Saida.Graphics.DrawString(V_DEmitente.IE, Font10, Brushes.Black, 25, 43.2)
    '    'Gra_Saida.Graphics.DrawString(V_DEmitente.IESUBS, Font10, Brushes.Black, 113, 43.2)
    '    'Gra_Saida.Graphics.DrawString(V_DEmitente.CNPJ, Font10, Brushes.Black, 202, 43.2)

    '    ''INSERIR BLOCOS DE CAMPOS
    '    'Dim Formato_Vertical As System.Drawing.StringFormat
    '    'Dim Posição_Texto As System.Drawing.Point
    '    'Dim Rotacionador_de_Texto As System.Drawing.Drawing2D.Matrix = Gra_Saida.Graphics.Transform()

    '    'Formato_Vertical = New StringFormat(StringFormatFlags.DirectionVertical)
    '    'Posição_Texto = New Point(27.5, 65)

    '    'Rotacionador_de_Texto.RotateAt(180, Posição_Texto)
    '    'Gra_Saida.Graphics.Transform = Rotacionador_de_Texto

    '    ''GUIA DO CLIENTE
    '    'Posição_Texto = New Point(47.5, -71)
    '    'Gra_Saida.Graphics.Transform = Rotacionador_de_Texto
    '    'Gra_Saida.Graphics.DrawString("RECEBEMOS DE", Font6, Brushes.Black, Posição_Texto, Formato_Vertical)

    '    'Posição_Texto = New Point(47.5, -53)
    '    'Gra_Saida.Graphics.Transform = Rotacionador_de_Texto
    '    ''Gra_Saida.Graphics.DrawString(V_DEmitente.NOME, Font6_B, Brushes.Black, Posição_Texto, Formato_Vertical)

    '    'Posição_Texto = New Point(47.5, -6)
    '    'Gra_Saida.Graphics.Transform = Rotacionador_de_Texto
    '    'Gra_Saida.Graphics.DrawString(", OS PRODUTOS OU SERVIÇOS CONSTANTES NA NOTA FISCAL ELETRÔNICA INDICADA AO LADO", Font6, Brushes.Black, Posição_Texto, Formato_Vertical)

    '    'Posição_Texto = New Point(41, -71)
    '    'Gra_Saida.Graphics.Transform = Rotacionador_de_Texto
    '    'Gra_Saida.Graphics.DrawString("DATA E HORA", Font6, Brushes.Black, Posição_Texto, Formato_Vertical)

    '    'Posição_Texto = New Point(41, -6)
    '    'Gra_Saida.Graphics.Transform = Rotacionador_de_Texto
    '    'Gra_Saida.Graphics.DrawString("IDENTIFICAÇÃO DO RECEBEDOR", Font6, Brushes.Black, Posição_Texto, Formato_Vertical)

    '    ''VOLTA TEXTO PARA POSIÇÃO ORIGINAL
    '    'Gra_Saida.Graphics.ResetTransform()

    'End Sub
    'Imagem do cabeçalho1
    Private Sub cbimagemClasse1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbimagemClasse1.SelectedIndexChanged
        If cbimagemClasse1.SelectedItem = "AutoSize" Then
            PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize
        ElseIf cbimagemClasse1.SelectedItem = "CenterImage" Then
            PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
        ElseIf cbimagemClasse1.SelectedItem = "Normal" Then
            PictureBox1.SizeMode = PictureBoxSizeMode.Normal
        ElseIf cbimagemClasse1.SelectedItem = "StretchImage" Then
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        ElseIf cbimagemClasse1.SelectedItem = "Zoom" Then
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        End If
        PictureBox1.Refresh()

    End Sub
    'Imagem do cabeçalho 2
    Private Sub cbimagem1Classe1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbimagem1Classe1.SelectedIndexChanged
        If cbimagem1Classe1.SelectedItem = "AutoSize" Then
            PictureBox2.SizeMode = PictureBoxSizeMode.AutoSize
        ElseIf cbimagem1Classe1.SelectedItem = "CenterImage" Then
            PictureBox2.SizeMode = PictureBoxSizeMode.CenterImage
        ElseIf cbimagem1Classe1.SelectedItem = "Normal" Then
            PictureBox2.SizeMode = PictureBoxSizeMode.Normal
        ElseIf cbimagem1Classe1.SelectedItem = "StretchImage" Then
            PictureBox2.SizeMode = PictureBoxSizeMode.StretchImage
        ElseIf cbimagem1Classe1.SelectedItem = "Zoom" Then
            PictureBox2.SizeMode = PictureBoxSizeMode.Zoom
        End If
        PictureBox2.Refresh()
    End Sub


    'A conexÆo e o DataReader ‚ aberto aqui
    'Private Function GetLinhasSelecionadas() As List(Of String)


    '    'Dim dgvColecaoLinhasSelecionadas As DataGridViewSelectedRowCollection = MetroGrid2.MetroGrid2.Rows


    '    Dim ids As New List(Of String)
    '    'For i As Integer = 0 To MetroGrid2.Rows.Count - 1
    '    '    Dim id As String = MetroGrid2.Rows(i).Cells(0).Value
    '    '    Dim nome As String = MetroGrid2.Rows(i).Cells(1).Value
    '    '    Dim endereco As String = MetroGrid2.Rows(i).Cells(2).Value
    '    '    MsgBox(id.ToString & "," & nome & " " & endereco)
    '    'Next
    '    Return ids


    'End Function
    Private Sub Begin_Print(ByVal sender As Object, ByVal e As Printing.PrintEventArgs)
        paginaAtual = 1
    End Sub
    'Configurações da grid vertical
    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton5.Click
        'define o objeto para visualizar a impressao
        Dim objPrintPreview As New PrintPreviewDialog
        Dim pd As Printing.PrintDocument = New Printing.PrintDocument()
        Impressao()

    End Sub

    Private Sub Impressao()

        Dim objPrintPreview As New PrintPreviewDialog

        nomeClasse = ComboBox1.SelectedItem

        distancia1 = MetroTextBox5.Text
        distancia2 = MetroTextBox6.Text
        distancia3 = MetroTextBox7.Text
        distancia4 = MetroTextBox8.Text
        distancia5 = MetroTextBox9.Text
        distancia6 = MetroTextBox10.Text

        distancia7 = MetroTextBox13.Text
        distancia8 = MetroTextBox12.Text
        distancia9 = MetroTextBox11.Text
        distancia10 = MetroTextBox3.Text
        distancia11 = MetroTextBox2.Text
        distancia12 = MetroTextBox1.Text

        'Tamanho da fonte selecinada em configurações
        inttamanhofontnormal = cbTamanhoFonte.SelectedItem
        'Título do Relatório
        RelatorioTitulo = "Lista de Alunos - "

        'Define os objetos printdocument e os eventos associados
        Dim pd As Printing.PrintDocument = New Printing.PrintDocument()

        'IMPORTANTE - definimos 2 eventos para tratar a impressão : PringPage e BeginPrint.
        AddHandler pd.PrintPage, New Printing.PrintPageEventHandler(AddressOf Me.ClasseUm)
        AddHandler pd.BeginPrint, New Printing.PrintEventHandler(AddressOf Me.Begin_Print)
        Try
            'define o formulário como maximizado e com Zoom
            With objPrintPreview
                .WindowState = FormWindowState.Maximized
                .Document = pd
                .PrintPreviewControl.Zoom = 0.65
                .Text = "Relacao de Alunos"
                .ShowDialog()
            End With
            'start = True
            'iii = 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Private Sub MetroButton20_Click(sender As Object, e As EventArgs) Handles MetroButton20.Click
        If ComboBox1.SelectedIndex = 0 Then
            MsgBox("Escolha uma Classe!", MsgBoxStyle.Information, "Seleção incorreta")
        Else
            indices = ComboBox1.SelectedIndex
            indices -= 1
            CarregaClasses()
        End If
    End Sub

    Private Sub CarregaClasses()
        Dim contProf As Integer = 0
        Dim contAlunos As Integer = 0
        totalDeAlunos = 0
        CargaBancoAlunos()
        mgProfessores.Rows.Clear()
        mgAlunos.Rows.Clear()


        'Campo vazio para ser carregar 
        Try
            asteristico = mgClasses.Rows(indices).Cells(4).Value.ToString
            Label1.Text = "Faixa etária de " &
                        mgClasses.Rows(indices).Cells(3).Value.ToString.Replace("00:00:00", "") &
                        " a " &
                        mgClasses.Rows(indices).Cells(2).Value.ToString.Replace("00:00:00", "")
        Catch ex As Exception
            MsgBox("Idades com problemas, verifique a tabela de CLasses no Menu Principal!")
        End Try

        For i As Integer = 0 To mgTotal.Rows.Count - 1
            Dim contador As String = mgTotal.Rows(i).Cells(0).Value
            Dim nomes As String = mgTotal.Rows(i).Cells(1).Value.ToString
            Dim datas As String = mgTotal.Rows(i).Cells(2).Value
            Dim tels As String = mgTotal.Rows(i).Cells(3).Value.ToString
            Dim _sexo4 As String = mgTotal.Rows(i).Cells(4).Value.ToString
            Dim _classe4 As String = mgTotal.Rows(i).Cells(5).Value.ToString().Replace(" (Inativo)", "").Replace("*", "")
            Dim prof As String = mgTotal.Rows(i).Cells(6).Value.ToString
            Dim _especial4 As String = mgTotal.Rows(i).Cells(7).Value.ToString
            Dim _batismo4 As String = mgTotal.Rows(i).Cells(8).Value.ToString
            Dim _email4 As String = mgTotal.Rows(i).Cells(9).Value.ToString
            Dim _inativo3 As String = mgTotal.Rows(i).Cells(10).Value.ToString
            Dim convdata As String = mgTotal.Rows(i).Cells(2).Value


            'Testes
            'If contador = 529 Then
            '    MsgBox(contador)
            'End If


            'Se não tiver data 
            'Se FOR INATIVO, é pulado.
            If mgTotal.Rows(i).Cells(10).Value.ToString <> True Then
                Dim dada As Date = Date.Parse(convdata)
                Dim Arquivonovo As String = Convert.ToDateTime(convdata)
                Arquivonovo = anoAtual - Arquivonovo.Remove(0, 6)
                'dada = Convert.ToDateTime(convdata)
                Dim ini As Integer
                Dim fim As Integer
                Try
                    ini = CInt(mgClasses.Rows(indices).Cells(2).Value.ToString)
                    fim = CInt(mgClasses.Rows(indices).Cells(3).Value.ToString)
                Catch ex As Exception
                    MsgBox("Valor inválido!")
                End Try
                'ini = Convert.ToDateTime(ini)
                ' Dim Arquivonovoini As String = Convert.ToDateTime(mgClasses.Rows(indices).Cells(2).Value.ToString)
                'Dim fim As Date = Date.Parse
                '(mgClasses.Rows(indices).Cells(3).Value.ToString = Convert.ToDateTime(fim)
                'Dim Arquivonovofim As String = Convert.ToDateTime(mgClasses.Rows(indices).Cells(3).Value.ToString)
                'Se for PROFESSOR e da classe que está no cadastro dele
                If mgTotal.Rows(i).Cells(6).Value.ToString = True And _classe4 = ComboBox1.SelectedItem.Replace(" (Inativo)", "").Replace("*", "") Then
                    mgProfessores.ColumnCount = 11
                    mgProfessores.Columns(0).Name = "ID"
                    mgProfessores.Columns(1).Name = "Name"
                    mgProfessores.Columns(2).Name = "Data"
                    mgProfessores.Columns(3).Name = "tels"
                    mgProfessores.Columns(4).Name = "Sexo"
                    mgProfessores.Columns(5).Name = "Classe"
                    mgProfessores.Columns(6).Name = "Professor"
                    mgProfessores.Columns(7).Name = "Especial"
                    mgProfessores.Columns(8).Name = "Batismo"
                    mgProfessores.Columns(9).Name = "E-mail"
                    mgProfessores.Columns(10).Name = "Inativo"
                    Dim row As String() = New String() {contador, nomes, datas, tels, _sexo4, _classe4, prof, _especial4, _batismo4, _email4, _inativo3}
                    mgProfessores.Rows.Add(row)
                    'Oculta as colunas abaixo
                    'gridProf2.Columns(0).Visible = False
                    mgProfessores.Columns(4).Visible = False
                    mgProfessores.Columns(5).Visible = False
                    mgProfessores.Columns(6).Visible = False
                    mgProfessores.Columns(7).Visible = False
                    mgProfessores.Columns(8).Visible = False
                    mgProfessores.Columns(9).Visible = False
                    mgProfessores.Columns(10).Visible = False
                    contProf += 1
                    'Se for aluno especial é da classe que está no cadastro dele
                ElseIf mgTotal.Rows(i).Cells(7).Value.ToString = "True" Then
                    If _classe4 = ComboBox1.SelectedItem.Replace(" (Inativo)", "").Replace("*", "") Then
                        mgAlunos.ColumnCount = 11
                        mgAlunos.Columns(0).Name = "ID"
                        mgAlunos.Columns(1).Name = "Name"
                        mgAlunos.Columns(2).Name = "Data"
                        mgAlunos.Columns(3).Name = "tels"
                        mgAlunos.Columns(4).Name = "Sexo"
                        mgAlunos.Columns(5).Name = "Classe"
                        mgAlunos.Columns(6).Name = "Professor"
                        mgAlunos.Columns(7).Name = "Especial"
                        mgAlunos.Columns(8).Name = "Batismo"
                        mgAlunos.Columns(9).Name = "E-mail"
                        mgAlunos.Columns(10).Name = "Inativo"
                        Dim row2 As String() = New String() {contador, nomes & "*", datas, tels, _sexo4, _classe4, prof, _especial4, _batismo4, _email4, _inativo3}
                        mgAlunos.Rows.Add(row2)
                        'Oculta as colunas abaixo
                        'MetroGrid5.Columns(0).Visible = False
                        mgAlunos.Columns(4).Visible = False
                        mgAlunos.Columns(5).Visible = False
                        mgAlunos.Columns(6).Visible = False
                        mgAlunos.Columns(7).Visible = False
                        mgAlunos.Columns(8).Visible = False
                        mgAlunos.Columns(9).Visible = False
                        mgAlunos.Columns(10).Visible = False
                        contAlunos += 1

                    End If

                    'Se não for professor 
                    'Nem aluno especial
                    'É separado pela idade
                Else
                    If ini <= Arquivonovo And Arquivonovo <= fim And mgTotal.Rows(i).Cells(6).Value.ToString <> "True" Then
                        mgAlunos.ColumnCount = 11
                        mgAlunos.Columns(0).Name = "ID"
                        mgAlunos.Columns(1).Name = "Name"
                        mgAlunos.Columns(2).Name = "Data"
                        mgAlunos.Columns(3).Name = "tels"
                        mgAlunos.Columns(4).Name = "Sexo"
                        mgAlunos.Columns(5).Name = "Classe"
                        mgAlunos.Columns(6).Name = "Professor"
                        mgAlunos.Columns(7).Name = "Especial"
                        mgAlunos.Columns(8).Name = "Batismo"
                        mgAlunos.Columns(9).Name = "E-mail"
                        mgAlunos.Columns(10).Name = "Inativo"
                        Dim row3 As String() = New String() {contador, nomes, datas, tels, _sexo4, _classe4, prof, _especial4, _batismo4, _email4, _inativo3}
                        mgAlunos.Rows.Add(row3)
                        'Oculta as colunas abaixo
                        'MetroGrid5.Columns(0).Visible = False
                        mgAlunos.Columns(4).Visible = False
                        mgAlunos.Columns(5).Visible = False
                        mgAlunos.Columns(6).Visible = False
                        mgAlunos.Columns(7).Visible = False
                        mgAlunos.Columns(8).Visible = False
                        mgAlunos.Columns(9).Visible = False
                        mgAlunos.Columns(10).Visible = False
                        contAlunos += 1
                    End If
                End If
            End If
            'End If
        Next
        mgProfessores.MultiSelect = False
        mgProfessores.AutoResizeColumns()
        mgAlunos.AutoResizeColumns()
        Try
            'Define coluna NOME (ALUNO) para preencher toda a grid
            mgProfessores.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            mgAlunos.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            'Ordena as colunas
            Me.mgProfessores.Sort(Me.mgProfessores.Columns(1), ListSortDirection.Ascending)
            Me.mgAlunos.Sort(Me.mgAlunos.Columns(1), ListSortDirection.Ascending)
        Catch ex As Exception
            MsgBox("Não existe informações no banco de dados", MsgBoxStyle.Exclamation)
        End Try
        tcSecundario.Refresh()
        Label1.Text = Label1.Text & " - Professores: " & contProf & " | Alunos: " & contAlunos
        totalDeAlunos = mgProfessores.Rows.Count - 1
        totalDeAlunos += mgAlunos.Rows.Count - 1
    End Sub

    Private Sub CarregaClassesResumo()
        totalDeAlunos = 0
        'CargaBancoAlunos()
        'Campo vazio para ser carregar 
        Try
            asteristico = CheckedListBox1.SelectedItem().ToString
        Catch ex As Exception
            MsgBox("Idades com problemas, verifique a tabela de CLasses no Menu Principal!")
        End Try

        For i As Integer = 0 To mgTotal.Rows.Count - 1
            Dim _classe4 As String = mgTotal.Rows(i).Cells(5).Value.ToString().Replace(" (Inativo)", "").Replace("*", "")
            'Obtem a data de nascimento
            'elimina dia e mes e hora
            'subtrai pelo ano atual
            'resultado convdata = idade!
            Dim convdata As String = mgTotal.Rows(i).Cells(2).Value.ToString
            convdata = convdata.Remove(0, 6).Replace(" 00:00:00", "")
            convdata = anoAtual - convdata
            'Se não tiver data 
            'Se FOR INATIVO, é pulado.
            If mgTotal.Rows(i).Cells(10).Value.ToString <> "True" Then
                'Dim dada As Date = Date.Parse(convdata)
                'Dim Arquivonovo As String = Convert.ToDateTime(convdata)
                'dada = Convert.ToDateTime(convdata)
                Dim ini As Integer = CInt(mgClasses.Rows(indices).Cells(2).Value.ToString)
                ' = Convert.ToDateTime(ini)
                'Dim Arquivonovoini As String = Convert.ToDateTime(mgClasses.Rows(indices).Cells(2).Value.ToString)
                Dim fim As Integer = CInt(mgClasses.Rows(indices).Cells(3).Value.ToString)
                'fim = Convert.ToDateTime(fim)
                'Dim Arquivonovofim As String = Convert.ToDateTime(mgClasses.Rows(indices).Cells(3).Value.ToString)

                'Se for PROFESSOR e da classe que está no cadastro dele
                If mgTotal.Rows(i).Cells(6).Value.ToString = "True" And _classe4 = asteristico Then
                    totalDeAlunos += 1
                    'Se for aluno especial é da classe que está no cadastro dele
                ElseIf mgTotal.Rows(i).Cells(7).Value.ToString = "True" Then
                    If _classe4 = asteristico Then
                        totalDeAlunos += 1
                    End If

                    'Se não for professor 
                    'Nem aluno especial
                    'É separado pela idade
                Else
                    If convdata >= ini And fim >= convdata And mgTotal.Rows(i).Cells(6).Value.ToString <> "True" Then
                        totalDeAlunos += 1
                    End If
                End If

            End If
            'End If
        Next
        Try
        Catch ex As Exception
            MsgBox("Não existe informações no banco de dados", MsgBoxStyle.Exclamation)
        End Try
    End Sub
    'Botão Consultar
    Private Sub MetroButton18_Click(sender As Object, e As EventArgs)
        consultasql()
    End Sub


    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        ' Displays an OpenFileDialog so the user can select a Cursor.
        Dim ofdImagem1 As New OpenFileDialog()
        ofdImagem1.Filter = "Imagem|*.jpg"
        ofdImagem1.Title = "Select a Cursor File"

        ' Show the Dialog.
        ' If the user clicked OK in the dialog and 
        ' a .CUR file was selected, open it.
        If ofdImagem1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ' Assign the cursor in the Stream to the Form's Cursor property.

            PictureBox1.Image = Image.FromFile(ofdImagem1.FileName)
            imagem1 = ofdImagem1.FileName
        End If
    End Sub
    Private Sub MetroButton27_Click(sender As Object, e As EventArgs) Handles MetroButton27.Click
        ' Displays an OpenFileDialog so the user can select a Cursor.
        Dim ofdImagem2 As New OpenFileDialog()
        ofdImagem2.Filter = "Imagem|*.jpg"
        ofdImagem2.Title = "Select a Cursor File"

        ' Show the Dialog.
        ' If the user clicked OK in the dialog and 
        ' a .CUR file was selected, open it.
        If ofdImagem2.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ' Assign the cursor in the Stream to the Form's Cursor property.

            PictureBox2.Image = Image.FromFile(ofdImagem2.FileName)
            imagem2 = ofdImagem2.FileName
        End If
    End Sub


    Private Sub MetroTextBox4_Click(sender As Object, e As EventArgs) Handles MetroTextBox4.Click
        If Not IsNumeric(MetroTextBox4.Text) And MetroTextBox4.Text <> "" Then
            MsgBox("Campo numérico!")
            'Else
            '    MetroTextBox4.Text = 50
        End If
    End Sub

    Private Sub MetroGrid1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles mgTotal.CellClick
        indice = e.RowIndex
        'ids = MetroGrid1.CurrentRow().Cells(0).Value.ToString
        txtIdPrincipal.Text = mgTotal.CurrentRow().Cells(0).Value.ToString             'ID
        TextBox9.Text = mgTotal.CurrentRow().Cells(1).Value.ToString             'Nome
        MaskedTextBox1.Text = mgTotal.CurrentRow().Cells(2).Value.ToString       'Data
        'Telefone
        MaskedTextBox2.Text = mgTotal.CurrentRow().Cells(3).Value.ToString          'Telefone


        MaskedTextBox2.Text = mgTotal.CurrentRow().Cells(3).Value.ToString()

        TextBox10.Text = mgTotal.CurrentRow().Cells(9).Value.ToString()          'E-mail
        If mgTotal.CurrentRow().Cells(4).Value.ToString() = "Homem" Then         'Sexo
            MetroComboBox1.SelectedIndex = 1
        ElseIf mgTotal.CurrentRow().Cells(4).Value.ToString() = "Mulher" Then    'sexo
            MetroComboBox1.SelectedIndex = 2
        Else
            MetroComboBox1.SelectedIndex = 0
        End If
        Dim novoTexto As String = mgTotal.CurrentRow().Cells(5).Value.ToString() 'Classe
        For y As Integer = 0 To MetroComboBox2.Items.Count - 1
            Dim novaClasse As String = MetroComboBox2.Items(y).replace(" (Inativo)", "").replace("*", "")
            If ((novoTexto = novaClasse) Or (novaClasse & "*" = novoTexto)) Then
                MetroComboBox2.SelectedIndex = y
                Exit For
            End If
        Next
        If mgTotal.CurrentRow().Cells(6).Value.ToString() = "True" Then          'Professor
            MetroCheckBox1.Checked = True
        Else
            MetroCheckBox1.Checked = False
        End If
        If mgTotal.CurrentRow().Cells(7).Value.ToString() = "True" Then          'Especial
            MetroCheckBox2.Checked = True
        Else
            MetroCheckBox2.Checked = False
        End If
        If mgTotal.CurrentRow().Cells(8).Value.ToString() = "True" Then          'Batizado
            MetroCheckBox3.Checked = True
        Else
            MetroCheckBox3.Checked = False
        End If
        If mgTotal.CurrentRow().Cells(10).Value = "True" Then                    'Inativo
            MetroCheckBox4.Checked = True
        Else
            MetroCheckBox4.Checked = False
        End If
        MetroLabel13.Text = mgTotal.CurrentRow().Cells(11).Value.ToString

        TestaIdade()

    End Sub

    Private Sub MetroTrackBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar1.Scroll
        MetroTextBox5.Text = MetroTrackBar1.Value
    End Sub

    Private Sub MetroTrackBar2_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar2.Scroll
        MetroTextBox6.Text = MetroTrackBar2.Value
    End Sub

    Private Sub MetroTrackBar3_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar3.Scroll
        MetroTextBox7.Text = MetroTrackBar3.Value
    End Sub

    Private Sub MetroTrackBar4_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar4.Scroll
        MetroTextBox8.Text = MetroTrackBar4.Value
    End Sub

    Private Sub MetroTrackBar5_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar5.Scroll
        MetroTextBox9.Text = MetroTrackBar5.Value
    End Sub

    Private Sub MetroTextBox5_Click(sender As Object, e As EventArgs) Handles MetroTextBox5.Click
        MetroTrackBar1.Value = MetroTextBox5.Text
    End Sub

    Private Sub MetroTextBox6_Click(sender As Object, e As EventArgs) Handles MetroTextBox6.Click
        MetroTrackBar2.Value = MetroTextBox6.Text
    End Sub

    Private Sub MetroTextBox7_Click(sender As Object, e As EventArgs) Handles MetroTextBox7.Click
        MetroTrackBar3.Value = MetroTextBox7.Text
    End Sub

    Private Sub MetroTextBox8_Click(sender As Object, e As EventArgs) Handles MetroTextBox8.Click
        MetroTrackBar4.Value = MetroTextBox8.Text
    End Sub

    Private Sub MetroTextBox9_Click(sender As Object, e As EventArgs) Handles MetroTextBox9.Click
        MetroTrackBar5.Value = MetroTextBox9.Text
    End Sub

    Private Sub MetroTrackBar6_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar6.Scroll
        MetroTextBox10.Text = MetroTrackBar6.Value
    End Sub

    Private Sub MetroTextBox10_Click(sender As Object, e As EventArgs) Handles MetroTextBox10.Click
        MetroTrackBar6.Value = MetroTextBox10.Text
    End Sub

    Private Sub MetroButton28_Click(sender As Object, e As EventArgs) Handles MetroButton28.Click
        MetroButton28.Enabled = False

        Dim convertedDate As Date
        Dim myCommand2 As MySqlCommand

        Try
            sqlconection.Open()
            _nome3 = TextBox9.Text
            _Nascimento3 = MaskedTextBox1.Text
            _telefone3 = MaskedTextBox2.Text
            _sexo3 = MetroComboBox3.SelectedIndex
            _classe3 = MetroComboBox4.SelectedItem
            _professor3 = MetroCheckBox1.CheckState
            _especial3 = MetroCheckBox2.CheckState
            _batismo3 = MetroCheckBox3.CheckState
            _email3 = TextBox10.Text
            _inativo3 = MetroCheckBox4.CheckState

            convertedDate = Convert.ToDateTime(_Nascimento3)

            _Nascimento3 = CDate(_Nascimento3).ToString("yyyy-MM-dd")

            Dim Sql3 As String = "INSERT INTO TOTAL (ALUNO, NASCIMENTO, TELEFONE, SEXO, CLASSE, PROFESSOR, ALUNOESPECIAL, BATISMO, EMAIL, INATIVO) VALUES ('" _
            & _nome3 & "', '" _
            & _Nascimento3 & "', '" _
            & _telefone3 & "', '" _
            & _sexo3 & "', '" _
            & _classe3 & "', '" _
            & _professor3 & "', '" _
            & _especial3 & "', '" _
            & _batismo3 & "', '" _
            & _email3 & "', '" _
            & _inativo3 & "');"


            'Dim cmd As New MySqlCommand
            myCommand2 = New MySqlCommand(Sql3, sqlconection)
            With myCommand2
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            MsgBox("Cadastrado com sucesso")
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
        sqlconection.Close()
    End Sub

    Private Sub MetroGrid3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles mgClasses.CellClick
        indice1 = e.RowIndex
        TextBox2.Text = mgClasses.CurrentRow().Cells(0).Value.ToString

        TextBox11.Text = mgClasses.CurrentRow().Cells(1).Value.ToString             'Nome

        mcbIdadeIni.SelectedIndex = mgClasses.CurrentRow().Cells(2).Value       'Data
        mcbIdadeFim.SelectedIndex = mgClasses.CurrentRow().Cells(3).Value          'data

        If mgClasses.CurrentRow().Cells(4).Value.ToString() = "True" Then          'Especial
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If
        If mgClasses.CurrentRow().Cells(5).Value.ToString() = "True" Then          'Batizado
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If

        Dim novoTexto As String = mgClasses.CurrentRow().Cells(6).Value.ToString() 'Categoria
        For y As Integer = 0 To MetroComboBox5.Items.Count - 1
            Dim novaClasse As String = MetroComboBox5.Items(y).replace(" (Inativo)", "").replace("*", "")
            If novoTexto = novaClasse Then
                MetroComboBox5.SelectedIndex = y
                Exit For
            End If
        Next

        'MetroComboBox5.Text = mgClasses.CurrentRow().Cells(6).Value.ToString()




        'If MaskedTextBox3.Text <> "" Then
        '    If MaskedTextBox3.Text Like "/" Then
        '    Else
        '        Dim dt As DateTime = CDate(MaskedTextBox3.Text).ToString("dd/MM/yyyy")
        '        dt = Convert.ToDateTime(dt)
        '        Dim ts As TimeSpan = DateTime.Today.Subtract(dt)
        '        Try
        '            MetroLabel21.Text = New DateTime(ts.Ticks).ToString("yy") - 1
        '            MetroLabel21.Refresh()
        '        Catch ex As Exception

        '        End Try
        '    End If
        'End If

        'If MaskedTextBox4.Text <> "" Then
        '    If MaskedTextBox4.Text Like "/" Then
        '    Else
        '        Dim dt As DateTime = CDate(MaskedTextBox4.Text).ToString("dd/MM/yyyy")
        '        dt = Convert.ToDateTime(dt)
        '        Dim ts As TimeSpan = DateTime.Today.Subtract(dt)
        '        Try
        '            MetroLabel20.Text = New DateTime(ts.Ticks).ToString("yy") - 1
        '            MetroLabel20.Refresh()
        '        Catch ex As Exception

        '        End Try
        '    End If
        'End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox2.Text = ""
        TextBox11.Text = ""
        'MaskedTextBox3.Text = ""
        'MaskedTextBox4.Text = ""
        CheckBox1.Checked = False
        CheckBox2.Checked = False

    End Sub

    Private Sub gridProf2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles mgProfessores.CellClick
        'Dim ids As Integer = gridProf2.CurrentRow().Cells(0).Value
        txtIdCLasse.Text = mgProfessores.CurrentRow().Cells(0).Value
        TextBox1.Text = mgProfessores.CurrentRow().Cells(1).Value.ToString            'Nome
        MaskedTextBox5.Text = mgProfessores.CurrentRow().Cells(2).Value       'Data
        MaskedTextBox6.Text = mgProfessores.CurrentRow().Cells(3).Value      'Telefone
        'MaskedTextBox2.Text = MetroGrid1.Rows(id).Cells(1).Value
        TextBox12.Text = mgProfessores.CurrentRow().Cells(9).Value.ToString()          'E-mail
        If mgProfessores.CurrentRow().Cells(4).Value.ToString() = "Homem" Then         'Sexo
            MetroComboBox3.SelectedIndex = 0
        ElseIf mgProfessores.CurrentRow().Cells(4).Value.ToString() = "Mulher" Then    'dsgdsg
            MetroComboBox3.SelectedIndex = 1
        Else
            MetroComboBox3.SelectedIndex = 2
        End If
        MetroComboBox4.SelectedItem = mgProfessores.CurrentRow().Cells(5).Value.ToString() 'Classe
        If mgProfessores.CurrentRow().Cells(6).Value.ToString() = "True" Then          'Professor
            MetroCheckBox5.Checked = True
        Else
            MetroCheckBox5.Checked = False
        End If
        If mgProfessores.CurrentRow().Cells(7).Value.ToString() = "True" Then          'Especial
            MetroCheckBox6.Checked = True
        Else
            MetroCheckBox6.Checked = False
        End If
        If mgProfessores.CurrentRow().Cells(8).Value.ToString() = "True" Then          'Batizado
            MetroCheckBox7.Checked = True
        Else
            MetroCheckBox7.Checked = False
        End If
        If mgProfessores.CurrentRow().Cells(10).Value = "True" Then                    'Inativo
            MetroCheckBox8.Checked = True
        Else
            MetroCheckBox8.Checked = False
        End If
    End Sub

    Private Sub MetroGrid5_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles mgAlunos.CellClick
        ids = mgAlunos.CurrentRow().Cells(0).Value.ToString

        'novo ----------------------------------------------------------------------------------------------------------------
        indice = e.RowIndex
        'ids = MetroGrid5.CurrentRow().Cells(0).Value.ToString
        txtIdCLasse.Text = mgAlunos.CurrentRow().Cells(0).Value.ToString             'ID
        TextBox1.Text = mgAlunos.CurrentRow().Cells(1).Value.ToString             'Nome
        MaskedTextBox5.Text = mgAlunos.CurrentRow().Cells(2).Value.ToString       'Data
        MaskedTextBox6.Text = mgAlunos.CurrentRow().Cells(3).Value.ToString       'Telefone
        TextBox12.Text = mgAlunos.CurrentRow().Cells(9).Value.ToString()          'E-mail
        If mgAlunos.CurrentRow().Cells(4).Value.ToString() = "Homem" Then         'Sexo
            MetroComboBox3.SelectedIndex = 1
        ElseIf mgAlunos.CurrentRow().Cells(4).Value.ToString() = "Mulher" Then    'sexo
            MetroComboBox3.SelectedIndex = 2
        Else
            MetroComboBox3.SelectedIndex = 0
        End If
        '----------------------------------------------------------------------------------------------------------------
        MetroComboBox4.SelectedItem = mgAlunos.CurrentRow().Cells(5).Value.ToString() 'Classe
        '----------------------------------------------------------------------------------------------------------------
        If mgAlunos.CurrentRow().Cells(6).Value.ToString() = "True" Then          'Professor
            MetroCheckBox5.Checked = True
        Else
            MetroCheckBox5.Checked = False
        End If
        '----------------------------------------------------------------------------------------------------------------
        If mgAlunos.CurrentRow().Cells(7).Value.ToString() = "True" Then          'Especial
            MetroCheckBox6.Checked = True
        Else
            MetroCheckBox6.Checked = False
        End If
        '----------------------------------------------------------------------------------------------------------------
        If mgAlunos.CurrentRow().Cells(8).Value.ToString() = "True" Then          'Batizado
            MetroCheckBox7.Checked = True
        Else
            MetroCheckBox7.Checked = False
        End If
        '----------------------------------------------------------------------------------------------------------------
        If mgAlunos.CurrentRow().Cells(10).Value = "True" Then                    'Inativo
            MetroCheckBox8.Checked = True
        Else
            MetroCheckBox8.Checked = False
        End If

        TestaIdade2()
    End Sub

    Private Sub MetroButton29_Click(sender As Object, e As EventArgs) Handles MetroButton29.Click

        Dim convertedDate As Date
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Try
            sqlconection.Open()
            stridCLasse = txtIdCLasse.Text
            _nome2 = TextBox1.Text.Replace("*", "")
            _Nascimento2 = MaskedTextBox5.Text
            _telefone2 = MaskedTextBox6.Text
            _sexo2 = MetroComboBox3.SelectedItem
            _classe2 = MetroComboBox4.SelectedItem
            _professor2 = MetroCheckBox5.CheckState
            _especial2 = MetroCheckBox6.CheckState
            _batismo2 = MetroCheckBox7.CheckState
            _email2 = TextBox12.Text
            _inativo2 = MetroCheckBox8.CheckState

            convertedDate = Convert.ToDateTime(_Nascimento2)

            _Nascimento2 = CDate(_Nascimento2).ToString("yyyy-MM-dd")

            SQL2 = "update total set ALUNO = '" &
                _nome2 & "', NASCIMENTO = '" &
                _Nascimento2 & "', TELEFONE = '" &
                _telefone2 & "', SEXO = '" &
                _sexo2 & "', CLASSE = '" &
                _classe2 & "', PROFESSOR = '" &
                _professor2 & "', ALUNOESPECIAL = '" &
                _especial2 & "', BATISMO = '" &
                _batismo2 & "', EMAIL = '" &
                _email2 & "', INATIVO = '" &
                _inativo2 & "' where contador = " & stridCLasse


            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            If _inativo2 = "True" Then
                MsgBox(_nome2 & ", agora está inativo!")
            Else
                MsgBox(_nome2 & ", alterado com sucesso!")
            End If

        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
        sqlconection.Close()
        CargaBancoLog()
        PintaLinas()
        CarregaClasses()
    End Sub



    Private Sub MetroButton4_Click_1(sender As Object, e As EventArgs) Handles MetroButton4.Click
        consultasql()
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        selecao = True
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Try
            AltLog()

            stridPrincipal = txtIdPrincipal.Text.Trim
            _nome2 = TextBox9.Text
            _Nascimento2 = MaskedTextBox1.Text
            _telefone2 = MaskedTextBox2.Text
            _sexo2 = MetroComboBox1.SelectedItem
            _classe2 = MetroComboBox2.SelectedItem
            _professor2 = MetroCheckBox1.CheckState
            _especial2 = MetroCheckBox2.CheckState
            _batismo2 = MetroCheckBox3.CheckState
            _email2 = TextBox10.Text
            _inativo2 = MetroCheckBox4.CheckState

            Dim convertedDate As Date
            convertedDate = Convert.ToDateTime(_Nascimento2)
            _Nascimento2 = CDate(_Nascimento2).ToString("yyyy-MM-dd")
            _classe2 = _classe2.Replace("*", "")
            'Sugestão de classe certa
            Dim dataAtualPessoa As String = MaskedTextBox1.Text
            dataAtualPessoa = dataAtualPessoa.Remove(0, 6)
            dataAtualPessoa = CInt(anoAtual - dataAtualPessoa)
            Dim data_1 As String
            Dim data_2 As String
            Dim _sugestao As String = ""
            For x As Integer = 0 To mgClasses.Rows.Count - 1
                data_1 = CInt(mgClasses.Rows(x).Cells(2).Value)
                data_2 = CInt(mgClasses.Rows(x).Cells(3).Value)
                If data_1 <= dataAtualPessoa And dataAtualPessoa <= data_2 Then
                    MetroLabel13.Text = mgClasses.Rows(x).Cells(1).Value.ToString
                    MetroLabel13.Refresh()
                    _sugestao = mgClasses.Rows(x).Cells(1).Value.ToString
                    Exit For
                End If
            Next
            If (_especial2 = "False") Or (_professor2 = "False") Then
                _classe2 = _sugestao
            End If
            'Fim da sugestão
            sqlconection.Open()
            'Adicionar _sugestao na QUERY
            SQL2 = "update total set ALUNO = '" &
                    _nome2 & "', NASCIMENTO = '" &
                    _Nascimento2 & "', TELEFONE = '" &
                    _telefone2 & "', SEXO = '" &
                    _sexo2 & "', CLASSE = '" &
                    _classe2 & "', PROFESSOR = '" &
                    _professor2 & "', ALUNOESPECIAL = '" &
                    _especial2 & "', BATISMO = '" &
                    _batismo2 & "', EMAIL = '" &
                    _email2 & "', INATIVO = '" &
                    _inativo2 & "', CLASSESUBESTAO = '" &
                    _sugestao & "' where contador = " & stridPrincipal

            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            If _inativo2 = "True" Then
                MsgBox(_nome2 & ", agora está inativo!")
            Else
                MsgBox(_nome2 & ", foi alterado com sucesso!")
            End If
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try

        sqlconection.Close()
        CargaBancoAlunos()
        CargaBancoLog()
        LimpaPrincipal()
        PintaLinas()
        TextBox9.Focus()

        mgTotal.CurrentCell = mgTotal.Rows(indice).Cells(0)
    End Sub

    Private Sub MetroButton7_Click(sender As Object, e As EventArgs) Handles MetroButton7.Click
        TextBox9.Text = ""
        MaskedTextBox1.Text = ""
        _telefone3 = MaskedTextBox2.Text = ""
        MetroComboBox3.SelectedIndex = 0
        MetroComboBox4.SelectedIndex = 0
        MetroCheckBox1.Checked = False
        MetroCheckBox2.Checked = False
        MetroCheckBox3.Checked = False
        TextBox10.Text = ""
        MetroCheckBox4.Checked = False
        'Habilita botão de alteração
        MetroButton28.Enabled = True
        'Apaga id para insersão usando incremento do banco de dados
        id = 0
    End Sub

    Private Sub MetroButton8_Click(sender As Object, e As EventArgs) Handles MetroButton8.Click
        LimpaPrincipal()

    End Sub

    Private Sub LimpaPrincipal()
        txtIdPrincipal.Text = ""
        TextBox9.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
        MetroComboBox3.SelectedIndex = 0
        MetroComboBox4.SelectedIndex = 0
        MetroCheckBox1.Checked = False
        MetroCheckBox2.Checked = False
        MetroCheckBox3.Checked = False
        TextBox10.Text = ""
        MetroCheckBox4.Checked = False
        'Habilita botão de alteração
        MetroButton28.Enabled = True
        'Apaga id para insersão usando incremento do banco de dados
        id = 0
        TextBox9.Focus()
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

        If buscaCampo.Checked = True Then
            Dim texto As String = Nothing
            If TextBox9.Text <> String.Empty Then
                'percorre cada linha do DataGridView
                'For i As Integer = 0 To MetroGrid1.Rows.Count - 1
                For Each linha As DataGridViewRow In mgTotal.Rows
                    For Each celula As DataGridViewCell In mgTotal.Rows(linha.Index).Cells
                        If celula.ColumnIndex = 1 Then
                            texto = celula.Value.ToString
                            'se o texto informado estiver contido na célula então seleciona toda linha
                            If texto.Contains(TextBox9.Text) Then
                                'seleciona a linha
                                mgTotal.CurrentCell = celula
                                Exit Sub
                            End If

                        End If
                    Next
                Next
                'se a coluna for a coluna 1 (Nome) então verifica o criterio
            End If
        End If
    End Sub

    Private Sub MetroButton10_Click(sender As Object, e As EventArgs) Handles MetroButton10.Click
        CarregaDuplicados()

    End Sub
    Private Sub CarregaDuplicados()
        If lvClasses.Items.Count - 1 > 0 Then
            lvClasses.Items.Clear()
        End If

        For iLinha As Integer = 0 To mgTotal.RowCount - 1
            For jLinha As Integer = 0 To mgTotal.RowCount - 1
                If mgTotal.Rows(jLinha).Cells(0).Value <> mgTotal.Rows(iLinha).Cells(0).Value Then
                    If mgTotal.Rows(jLinha).Cells(1).Value.ToString = mgTotal.Rows(iLinha).Cells(1).Value.ToString Then
                        lvClasses.Items.Add(mgTotal.Rows(jLinha).Cells(1).Value.ToString & " (" & mgTotal.Rows(jLinha).Cells(0).Value.ToString & ")")
                    End If
                End If
            Next
        Next
        lvClasses.Sorting = SortOrder.Ascending
    End Sub

    Private Sub lvClasses_ItemSelectionChanged(sender As Object, e As ListViewItemSelectionChangedEventArgs) Handles lvClasses.ItemSelectionChanged
        Dim selecao As String = lvClasses.Items(e.ItemIndex).Text
        selecao = String.Join(Nothing, System.Text.RegularExpressions.Regex.Split(selecao, "[^\d]"))
        For i As Integer = 0 To mgTotal.Rows.Count - 1
            If selecao = mgTotal.Rows(i).Cells(0).Value Then
                mgTotal.CurrentCell = mgTotal.Rows(i).Cells(0)
                Exit For
            End If
        Next
    End Sub

    Private Sub MetroButton9_Click(sender As Object, e As EventArgs) Handles MetroButton9.Click
        If txtIdPrincipal.Text = "" Then
            MsgBox("Código principal deve ser informado.")
            txtIdPrincipal.Focus()
            Return
        End If
        If MsgBox("Deseja apagar o(a) aluno(a): " + mgTotal.CurrentRow().Cells(1).Value.ToString + "?", MsgBoxStyle.YesNo, "Deletando Aluno:") = MsgBoxResult.No Then
            Return
        Else
            Dim convertedDate As Date
            Dim myCommand As New MySqlCommand
            Dim SQL2 As String
            Try
                DelLog()
                sqlconection.Open()

                convertedDate = Convert.ToDateTime(_Nascimento2)

                _Nascimento2 = CDate(_Nascimento2).ToString("yyyy-MM-dd")

                SQL2 = "delete from total where contador = " & txtIdPrincipal.Text


                'Dim cmd As New MySqlCommand
                myCommand = New MySqlCommand(SQL2, sqlconection)
                With myCommand
                    '.CommandText = SQL
                    .CommandType = CommandType.Text
                    '.Connection = sqlconection
                    .ExecuteNonQuery()
                End With
                MsgBox("Aluno deletado com sucesso!")
            Catch ex As Exception
                MsgBox("Erro : " & ex.Message)
            End Try
            sqlconection.Close()
            CargaBancoAlunos()
            CargaBancoLog()
            'MetroGrid1.CurrentCell = MetroGrid1.Rows(MetroGrid1.Rows.Count).Cells(0)
            LimpaPrincipal()
            CarregaDuplicados()
        End If
        PintaLinas()
    End Sub
    Private Sub TestaIdade()
        If MaskedTextBox1.Text <> "" Then
            'If MetroLabel1.Text <> 0 Then
            If MaskedTextBox1.Text Like "/" Then
            Else
                Dim dt As DateTime = CDate(MaskedTextBox1.Text).ToString("dd/MM/yyyy")
                dt = Convert.ToDateTime(dt)
                Dim ts As TimeSpan = DateTime.Today.Subtract(dt)
                MetroLabel1.Text = New DateTime(ts.Ticks).ToString("yy") - 1
                MetroLabel1.Refresh()

            End If
        End If
        'End If
    End Sub
    Private Sub TestaIdade2()
        If MaskedTextBox5.Text Like "/" Then
        Else
            Dim dt As DateTime = CDate(MaskedTextBox5.Text).ToString("dd/MM/yyyy")
            dt = Convert.ToDateTime(dt)
            Dim ts As TimeSpan = DateTime.Today.Subtract(dt)
            MetroLabel2.Text = New DateTime(ts.Ticks).ToString("yy")
        End If
    End Sub
    Private Sub MaskedTextBox1_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles MaskedTextBox1.MaskInputRejected


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If TextBox2.Text = "" Then
            MsgBox("O campo ID deve conter um número referente a alteração do registro.")
            Return
        End If
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim _cId As String = ""
        Dim _cNome As String
        Dim _dInicio As String
        Dim _dFim As String
        Dim _cEspecial As Boolean = False
        Dim _cInativo As Boolean = False
        Dim _cCategoria As String
        Try
            Dim dtInicio As Integer = mcbIdadeIni.Text
            Dim dtFim As Integer = mcbIdadeFim.Text
            'Dim anoAtual As Integer = CInt(Format(dataAtual, "yyyy"))
            'dtInicio = anoAtual - dtInicio
            'dtFim = anoAtual - dtFim
            _dInicio = dtInicio
            _dFim = dtFim
            sqlconection.Open()
            _cId = TextBox2.Text
            _cNome = TextBox11.Text
            _cEspecial = CheckBox1.CheckState
            _cInativo = CheckBox2.CheckState
            _cCategoria = MetroComboBox5.SelectedItem
            'Fornece o id da categoria
            For i As Integer = 0 To mgCategoria.Rows.Count - 1
                If _cCategoria = mgCategoria.Rows(i).Cells(1).Value Then
                    _cCategoria = mgCategoria.Rows(i).Cells(0).Value
                    Exit For
                End If
            Next

            'Dim convertedDate As Date
            'convertedDate = Convert.ToDateTime(_dInicio)
            '_dInicio = CDate(_dInicio).ToString("yyyy-MM-dd")


            'Dim convertedDate2 As Date
            'convertedDate2 = Convert.ToDateTime(_dFim)
            '_dFim = CDate(_dFim).ToString("yyyy-MM-dd")

            'If _cCategoria = "Crianças" Then
            '    _cCategoria = 1
            'ElseIf _cCategoria = "Jovens e Adultos" Then
            '    _cCategoria = 2
            'Else
            '    _cCategoria = 3
            'End If

            SQL2 = "update classes set Classe = '" & _cNome &
                "', dataini = '" & _dInicio &
                "', datafim = '" & _dFim &
                "', especial = '" & _cEspecial &
                "', inativo = '" & _cInativo &
                "', idcategoria = '" & _cCategoria &
                "' where contador = " & _cId


            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            MsgBox("Cadastrado alterado com sucesso")
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
        sqlconection.Close()

        CargaBancoClasses()
        CargaBancoAlunos()
        AbasClasses()
        CargaBancoClassesResumo()
        PintaLinasClasses()

        mgClasses.CurrentCell = mgClasses.Rows(indice1).Cells(0)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        CargaBancoLog()
    End Sub

    Private Sub MetroTrackBar12_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar12.Scroll
        MetroTextBox13.Text = MetroTrackBar12.Value
    End Sub

    Private Sub MetroTrackBar11_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar11.Scroll
        MetroTextBox12.Text = MetroTrackBar11.Value
    End Sub

    Private Sub MetroTrackBar10_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar10.Scroll
        MetroTextBox11.Text = MetroTrackBar10.Value
    End Sub

    Private Sub MetroTrackBar9_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar9.Scroll
        MetroTextBox3.Text = MetroTrackBar9.Value
    End Sub

    Private Sub MetroTrackBar8_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar8.Scroll
        MetroTextBox2.Text = MetroTrackBar8.Value
    End Sub

    Private Sub MetroTrackBar7_Scroll(sender As Object, e As ScrollEventArgs) Handles MetroTrackBar7.Scroll
        MetroTextBox1.Text = MetroTrackBar7.Value
    End Sub

    Private Sub MetroTextBox13_Click(sender As Object, e As EventArgs) Handles MetroTextBox13.Click
        MetroTrackBar12.Value = MetroTextBox13.Text
    End Sub

    Private Sub MetroTextBox12_Click(sender As Object, e As EventArgs) Handles MetroTextBox12.Click
        MetroTrackBar11.Value = MetroTextBox12.Text
    End Sub

    Private Sub MetroTextBox11_Click(sender As Object, e As EventArgs) Handles MetroTextBox11.Click
        MetroTrackBar10.Value = MetroTextBox11.Text
    End Sub

    Private Sub MetroTextBox3_Click(sender As Object, e As EventArgs) Handles MetroTextBox3.Click
        MetroTrackBar9.Value = MetroTextBox3.Text
    End Sub

    Private Sub MetroTextBox2_Click(sender As Object, e As EventArgs) Handles MetroTextBox2.Click
        MetroTrackBar8.Value = MetroTextBox2.Text
    End Sub

    Private Sub MetroTextBox1_Click(sender As Object, e As EventArgs) Handles MetroTextBox1.Click
        MetroTrackBar8.Value = MetroTextBox1.Text
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim _cNome As String

        Dim _cEspecial As Boolean = False
        Dim _cInativo As Boolean = False
        Dim _cCategoria As String
        Try
            Dim dtInicio As Integer = mcbIdadeIni.Text
            Dim dtFim As Integer = mcbIdadeFim.Text
            'Dim anoAtual As Integer = CInt(Format(dataAtual, "yyyy"))
            'dtInicio = anoAtual - dtInicio
            'dtFim = anoAtual - dtFim
            _dInicio = dtInicio
            _dFim = dtFim

            sqlconection.Open()

            _cNome = TextBox11.Text

            _cEspecial = CheckBox1.CheckState
            _cInativo = CheckBox2.CheckState
            _cCategoria = MetroComboBox5.SelectedItem
            'Dim convertedDate As Date
            'convertedDate = Convert.ToDateTime(_dInicio)
            '_dInicio = CDate(_dInicio).ToString("yyyy-MM-dd")
            'Verifica duplicidade
            For f As Integer = 0 To mgClasses.Rows.Count - 1
                If TextBox11.Text = mgClasses.Rows(f).Cells(1).Value.ToString Then
                    MsgBox("Já existe o nome de classe: " & TextBox11.Text)
                    TextBox11.Focus()
                    sqlconection.Close()
                    Return
                End If
            Next

            For i As Integer = 0 To mgCategoria.Rows.Count - 1
                If _cCategoria = mgCategoria.Rows(i).Cells(1).Value Then
                    _cCategoria = mgCategoria.Rows(i).Cells(0).Value
                    Exit For
                End If
            Next
            'Dim convertedDate2 As Date
            'convertedDate2 = Convert.ToDateTime(_dFim)
            '_dFim = CDate(_dFim).ToString("yyyy-MM-dd")


            SQL2 = "INSERT INTO CLASSES (Classe, dataini, datafim, especial, inativo, idcategoria) values ('" &
                _cNome & "', '" & _dInicio & "', '" & _dFim & "', '" & _cEspecial & "', '" & _cInativo & "', '" & _cCategoria & "');"


            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            MsgBox("Cadastrado alterado com sucesso")
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try

        sqlconection.Close()
        'CargaBancoClasses()
        'CargaBancoAlunos()
        'CargaBancoClassesResumo()
        'PintaLinasClasses()
        CarregaTudo()


        'MetroGrid3.CurrentCell = MetroGrid3.Rows(indice1).Cells(0)
    End Sub

    Private Sub MetroToggle1_CheckedChanged(sender As Object, e As EventArgs) Handles MetroToggle1.CheckedChanged
        If MetroToggle1.Checked = True Then
            SplitContainer1.Orientation = Orientation.Horizontal
        Else
            SplitContainer1.Orientation = Orientation.Vertical
        End If
    End Sub

    Private Sub txtIdPrincipal_TextChanged(sender As Object, e As EventArgs) Handles txtIdPrincipal.TextChanged
        'Localiza seleção pelo ID do aluno
        If CheckBox3.Checked = True Then
            Dim texto As String = Nothing
            If txtIdPrincipal.Text <> String.Empty Then
                'percorre cada linha do DataGridView
                'For i As Integer = 0 To MetroGrid1.Rows.Count - 1
                For Each linha As DataGridViewRow In mgTotal.Rows
                    For Each celula As DataGridViewCell In mgTotal.Rows(linha.Index).Cells
                        If celula.ColumnIndex = 0 Then
                            texto = celula.Value.ToString
                            'se o texto informado estiver contido na célula então seleciona toda linha
                            If texto.Contains(txtIdPrincipal.Text) Then
                                'seleciona a linha
                                mgTotal.CurrentCell = celula
                                Exit Sub
                            End If

                        End If
                    Next
                Next
                'se a coluna for a coluna 1 (Nome) então verifica o criterio
            End If
        End If

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        'Localiza seleção pelo ID do aluno
        Dim texto As String = Nothing

        If txtIdPrincipal.Text <> String.Empty Then
            'percorre cada linha do DataGridView
            'For i As Integer = 0 To MetroGrid1.Rows.Count - 1
            For Each linha As DataGridViewRow In mgTotal.Rows
                For Each celula As DataGridViewCell In mgTotal.Rows(linha.Index).Cells
                    If celula.ColumnIndex = 9 Then
                        texto = celula.Value.ToString
                        'se o texto informado estiver contido na célula então seleciona toda linha
                        If texto.Contains(txtIdPrincipal.Text) Then
                            'seleciona a linha
                            mgTotal.CurrentCell = celula
                            Exit Sub
                        End If

                    End If
                Next
            Next
            'se a coluna for a coluna 1 (Nome) então verifica o criterio
        End If
    End Sub

    Private Sub MetroButton11_Click(sender As Object, e As EventArgs)
        'If ConsultaExistenciaTabela(retornoTabela:=True) Then
        '    If sqlconection.State = ConnectionState.Open Then
        '        sqlconection.Close()
        '    End If
        '    MsgBox("Tabela existe")

        'Else
        '    If sqlconection.State = ConnectionState.Open Then
        '        sqlconection.Close()
        '    End If
        '    MsgBox("Tabela não existe")
        'End If
        CargaBancoResumo()

    End Sub
    Private Function ConsultaExistenciaTabela(retornoTabela As String) As Boolean
        Dim myDataAd As MySqlDataAdapter
        Dim tabela As New DataTable

        sqlconection.Open()
        Dim Sql3 As String = "SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' ORDER BY TABLE_TYPE"

        myDataAd = New MySqlDataAdapter(Sql3, sqlconection)

        myDataAd.Fill(tabela)

        Try
            For Each dr As DataRow In tabela.Rows
                If dr("TABLE_NAME").ToString = "log" Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            MessageBox.Show("ERRO " & ex.Message, "Verifica tabela")
            Return False
        End Try
        sqlconection.Close()
    End Function

    Private Sub MetroButton12_Click(sender As Object, e As EventArgs)

        CargaBancoClassesResumo()

    End Sub

    Public Sub CargaBancoClassesResumo()
        CheckedListBox1.Items.Clear()
        CheckedListBox1.Items.Add("Escolha a Classe")
        For i As Integer = 0 To mgClasses.Rows.Count - 1
            'TENTATIVA DE POR CORES POR LINHAS
            'If ((i Mod 2) = 0) Then
            '    lbResumo.ForeColor = Color.BlueViolet
            'End If
            If (mgClasses.Rows(i).Cells(5).Value.ToString() <> "False") Then
                CheckedListBox1.Items.Add(mgClasses.Rows(i).Cells(1).Value.ToString() + "(X)")
            Else
                CheckedListBox1.Items.Add(mgClasses.Rows(i).Cells(1).Value.ToString())
                feitasNoResumo += 1
            End If
        Next
        GroupBox4.Text = GroupBox4.Text & " - faltam (" & feitasNoResumo & ") classes."

    End Sub

    Private Sub MetroButton13_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub AcertaResumo()
        Dim sql As String

        'Dim convertedDate As Date
        Dim myCommand As New MySqlCommand
        For i As Integer = 0 To mgClasses.Rows.Count - 1
            Try
                sqlconection.Open()


                Dim classeR As String = mgClasses.Rows(i).Cells(1).Value.ToString()
                sql = "ALTER TABLE `resumo` ADD `" + classeR + "` VARCHAR(20);"

                'Dim cmd As New MySqlCommand
                myCommand = New MySqlCommand(sql, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                    .ExecuteNonQuery()
                End With


            Catch ex As Exception
                sqlconection.Close()
                'MsgBox("Erro : " & ex.Message)
            End Try
        Next
        sqlconection.Close()
        CargaBancoResumo()

    End Sub

    Private Sub MetroButton14_Click(sender As Object, e As EventArgs) Handles MetroButton14.Click
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim _rData As Date
        _rData = Convert.ToDateTime(dtRelatorio.Text)
        _rData = _rData.ToString("yyyy-MM-dd")


        If CheckedListBox1.SelectedIndex = 0 Then
            Return
        End If
        If ((MetroTextBox21.Text = "0") Or (MetroTextBox21.Text = "") Or (MetroTextBox20.Text = "0") Or (MetroTextBox20.Text = "")) _
            And MetroToggle3.Checked = False Then
            MsgBox("Não permitido valor zerado no resumo das classes." & vbCrLf &
                   "Para ativar vá em Configurações.")
            Return
        End If

        Try
            sqlconection.Open()
            MetroTextBox20.Text = totalDeAlunos
            Dim _rOfertas As String
            For i As Integer = 0 To mgResumo.Rows.Count - 1
                Dim datag As Date = mgResumo.Rows(i).Cells(1).Value.ToString
                If (_rCLasse = mgResumo.Rows(i).Cells(2).Value.ToString) And (_rData = datag) Then
                    For m As Integer = 0 To mgClasses.Rows.Count - 1
                        If mgResumo.Rows(i).Cells(2).Value.ToString = mgClasses.Rows(m).Cells(1).Value.ToString Then
                            _rIdClasse = mgClasses.Rows(m).Cells(0).Value.ToString
                        End If
                    Next
                    Dim result As Integer = MessageBox.Show("A classe " & mgResumo.Rows(i).Cells(2).Value.ToString & " já foi registrada, deseja atualizar?", "Atualização de Classe", MessageBoxButtons.YesNo)
                    If result = DialogResult.No Then
                        Return
                    Else
                        If IsNumeric(MetroTextBox23.Text) Then
                            _rOfertas = MetroTextBox23.Text
                            _rOfertas = _rOfertas.Replace(",", ".")
                        Else
                            MsgBox("Somente Números!")
                            Return
                        End If
                        _rVisitantes = MetroTextBox26.Text
                        If MetroTextBox26.Text = "" Then
                            _rVisitantes = 0
                        End If

                        SQL2 = "UPDATE RESUMOS SET totalalunos = '" & totalDeAlunos &
                            "',  presentes = '" & _rPresentes &
                            "', ausentes = '" & _rAusentes &
                            "', visitantes = '" & _rVisitantes &
                            "', ofertas = '" & _rOfertas &
                            "' WHERE id_classes = '" & _rIdClasse &
                            "' AND data = '" & _rData & "';"

                        myCommand = New MySqlCommand(SQL2, sqlconection)
                        With myCommand
                            '.CommandText = SQL
                            .CommandType = CommandType.Text
                            '.Connection = sqlconection
                            .ExecuteNonQuery()
                        End With
                        'MsgBox("Atualizações salvos com sucesso")
                        CheckedListBox1.SetItemCheckState(indices + 1, CheckState.Checked)
                        MetroTextBox21.Text = "0"
                        MetroTextBox26.Text = "0"
                        MetroTextBox23.Text = "0,00"
                        sqlconection.Close()
                        CargaBancoResumo()
                        TotalOfertas()
                        Return
                    End If
                End If
            Next
            'AltClasseLog()




            If IsNumeric(MetroTextBox23.Text) Then
                _rOfertas = MetroTextBox23.Text
                _rOfertas = _rOfertas.Replace(",", ".")
            Else
                MsgBox("Somente Números!")
                Return
            End If
            _rVisitantes = MetroTextBox26.Text
            If MetroTextBox26.Text = "" Then
                _rVisitantes = 0
            End If

            SQL2 = "INSERT INTO resumos (id_classes, totalalunos, presentes, ausentes, visitantes, ofertas, data) VALUES (" & _rCLasseId & ", " & totalDeAlunos & ", " & _rPresentes & ", " & _rAusentes & ", " & _rVisitantes & ", " & _rOfertas & ", '" & _rData & "')"

            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            feitasMenos = feitasNoResumo - mgResumo.Rows.Count
            GroupBox4.Text = "Relatório da CLasses Dominical - faltam (" & feitasMenos & ") classes."
            MsgBox("Dados salvos com sucesso")
            CheckedListBox1.SetItemCheckState(indices + 1, CheckState.Checked)

        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
        MetroTextBox21.Text = "0"
        MetroTextBox26.Text = "0"
        MetroTextBox23.Text = "0,00"
        sqlconection.Close()
        CargaBancoResumo()
        TotalOfertas()
    End Sub

    Private Sub MetroButton15_Click(sender As Object, e As EventArgs) Handles MetroButton15.Click
        indices = 0
        'define o objeto para visualizar a impressao
        Dim objPrintPreview As New PrintPreviewDialog
        Dim pd As Printing.PrintDocument = New Printing.PrintDocument()
        For indices = 0 To ComboBox1.Items.Count - 1
            Dim asteristico As String = "*"
            Dim inativo As String = "Inativo"
            If (ComboBox1.Items(indices) Like asteristico) Or (ComboBox1.Items(indices) Like inativo) Then
                '    MsgBox(ComboBox1.Items(indices))
                'Else
                MsgBox("Corrento " + ComboBox1.Items(indices))
                CarregaClasses()
                Impressao()
                Try
                    'define o formulário como maximizado e com Zoom
                    With objPrintPreview
                        .WindowState = FormWindowState.Maximized
                        .PrintPreviewControl.Zoom = 0.65
                        .Text = "Relacao de Alunos"
                        .ShowDialog()
                    End With
                    'start = True
                    'iii = 0
                Catch ex As Exception
                    MessageBox.Show(ex.ToString())
                End Try
            End If
        Next
    End Sub

    Private Sub MetroTextBox21_Click(sender As Object, e As EventArgs) Handles MetroTextBox21.Click
        If MetroTextBox21.Text <> "0" Then
            Return
        Else
            If MetroTextBox21.Text <> "" Then
                _rPresentes = MetroTextBox21.Text
                _rAusentes = (totalDeAlunos - _rPresentes)
                MetroTextBox22.Text = _rAusentes
            Else
                Return
            End If
        End If
    End Sub
    Private Sub MetroTextBox23_Click(sender As Object, e As EventArgs) Handles MetroTextBox23.Click
        If (MetroTextBox23.Text = "0,00") Or (MetroTextBox23.Text = "") Then
            MetroTextBox23.Text = ""
        Else
            _rPresentes = MetroTextBox21.Text
            _rAusentes = (totalDeAlunos - _rPresentes)
            MetroTextBox22.Text = _rAusentes

        End If
    End Sub

    Private Sub MetroTextBox21_Leave(sender As Object, e As EventArgs) Handles MetroTextBox21.Leave
        If IsNumeric(MetroTextBox21.Text) Then
            If MetroTextBox21.Text <> "" Then
                Dim rTotal As Integer = MetroTextBox20.Text
                Dim rPresente As Integer = MetroTextBox21.Text
                If ((rPresente < 0) Or (rTotal < rPresente)) Then
                    MsgBox("Valor menor que 0 ou Maior que o total da classe!")
                    MetroTextBox21.Text = ""
                    MetroTextBox21.Focus()
                    Return
                End If
                _rPresentes = MetroTextBox21.Text
                _rAusentes = (totalDeAlunos - _rPresentes)
                MetroTextBox22.Text = _rAusentes
            Else
                Return
            End If
        Else
            MetroTextBox21.Text = "0"
            Return
        End If
    End Sub

    Private Sub TotalOfertas()
        Dim soma As Double = "0"
        Dim reserva As Double = MetroTextBox24.Text
        For i As Integer = 0 To mgResumo.Rows.Count - 1
            soma += mgResumo.Rows(i).Cells(7).Value
        Next
        MetroTextBox25.Text = soma - reserva
    End Sub

    Private Sub MetroButton11_Click_1(sender As Object, e As EventArgs)

    End Sub
    Private Sub BackupMySql()
        Dim localDir As String = Application.StartupPath
        Dim strData As String = Date.Now.ToShortDateString
        Dim fileName As String = strData.Replace("/", "-") & "_MeuArquivo.sql"
        Dim saveDile As String = localDir & fileName
        Dim DBServer As String = "mysql host"
        Dim DBServerPort As String = "mysql port"
        Dim DataBase As String = "satabase name"
        Dim DBUser As String = "root"
        Dim DBPass As String = ""

    End Sub

    Private Sub MaskedTextBox1_Leave(sender As Object, e As EventArgs) Handles MaskedTextBox1.Leave
        If MaskedTextBox1.Text <> "" Then
            Dim _sugestao As String = ""
            Dim dataAtualPessoa As String = MaskedTextBox1.Text
            dataAtualPessoa = dataAtualPessoa.Remove(0, 6)
            dataAtualPessoa = CInt(anoAtual - dataAtualPessoa)
            MetroLabel1.Text = dataAtualPessoa.ToString
            Dim data_1 As Integer
            Dim data_2 As Integer
            For x As Integer = 0 To mgClasses.Rows.Count - 1
                'dataAtualPessoa = CDate(dataAtualPessoa.ToString).ToString("yyyy-MM-dd")
                data_1 = CInt(mgClasses.Rows(x).Cells(2).Value)
                data_2 = CInt(mgClasses.Rows(x).Cells(3).Value)
                If data_1 <= dataAtualPessoa And dataAtualPessoa <= data_2 And mgClasses.Rows(x).Cells(5).Value.ToString = "False" Then
                    'MetroLabel13.Text = mgClasses.CurrentRow().Cells(1).Value.ToString
                    MetroLabel13.Text = mgClasses.Rows(x).Cells(1).Value.ToString
                    MetroLabel13.Refresh()
                    If (MetroCheckBox1.Checked = False) And (MetroCheckBox2.Checked = False) Then
                        MetroComboBox2.Text = mgClasses.Rows(x).Cells(1).Value.ToString
                    End If
                    Exit For
                End If
            Next
        End If
        If MetroCheckBox1.Checked = False And MetroCheckBox2.Checked = False Then
            MetroComboBox2.Enabled = False
        Else
            MetroComboBox2.Enabled = True
        End If


    End Sub

    Private Sub MetroButton16_Click(sender As Object, e As EventArgs) Handles MetroButton16.Click
        maxDominical = MetroTextBox27.Text
        Dim objPrintPreview As New PrintPreviewDialog
        Dim pd As Printing.PrintDocument = New Printing.PrintDocument()
        GeraRelatorioGeral()
    End Sub

    Private Sub GeraRelatorioGeral()

        Dim objPrintPreview As New PrintPreviewDialog


        'Tamanho da fonte selecinada em configurações
        inttamanhofontnormal = cbTamanhoFonte.SelectedItem
        'Título do Relatório
        RelatorioTitulo = "Relatório Geral - "

        'Define os objetos printdocument e os eventos associados
        Dim pd As Printing.PrintDocument = New Printing.PrintDocument()

        'IMPORTANTE - definimos 2 eventos para tratar a impressão : PringPage e BeginPrint.
        AddHandler pd.PrintPage, New Printing.PrintPageEventHandler(AddressOf Me.RelatGeral)
        AddHandler pd.BeginPrint, New Printing.PrintEventHandler(AddressOf Me.Begin_Print)
        Try
            'define o formulário como maximizado e com Zoom
            With objPrintPreview
                .WindowState = FormWindowState.Maximized
                .Document = pd
                .PrintPreviewControl.Zoom = 0.65
                .Text = "Relacao de Alunos"
                .ShowDialog()
            End With
            'start = True
            'iii = 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Private Sub RelatGeral(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        _Ttfertas = 0
        _ttAusentes = 0
        _ttPresentes = 0
        _ttVisitantes = 0
        _ttAlunos = 0
        Dim z As String = "?,??"
        Dim za As String = "?,?"
        Dim zz As String = "??,??"
        Dim zza As String = "??,?"
        Dim zzz As String = "???,??"
        Dim zzza As String = "???,?"
        Dim zzzz As String = "????,??"
        Dim zzzza As String = "????,?"
        'Variaveis das linhas
        n = 1
        Dim LinhasPorPagina As Single = 0
        Dim PosicaoDaLinha As Single = 160
        Dim PosicaoDaLinha2 As Single = 0
        Dim LinhaAtual As Integer = 0

        'Variaveis das margens
        Dim MargemEsquerda As Single = e.MarginBounds.Left - 10
        Dim MargemSuperior As Single = e.MarginBounds.Top + 60
        Dim MargemSuperior2 As Single = e.MarginBounds.Top + 60
        Dim MargemDireita As Single = e.MarginBounds.Right + 80
        Dim MargemInferior As Single = e.MarginBounds.Bottom + 30
        Dim CanetaDaImpressora As Pen = New Pen(Color.Black, 1)
        'Dim codigo As Integer
        Dim esquerda As Integer = 35
        Dim direita As Integer = 795
        'Variaveis das fontes
        Dim FonteNegrito As Font
        Dim FonteTitulo As Font
        Dim FonteSubTitulo As Font
        Dim FonteRodape As Font
        Dim FonteNormal As Font
        Dim FonteNormalProf As Font
        Dim FonteNormaltel As Font
        Dim FonteNormaltel2 As Font
        'Dim totalPaginas As Integer

        'define efeitos em fontes
        FonteNegrito = New Font("Arial", 12, FontStyle.Bold)
        FonteTitulo = New Font("Century Gothic", 20, FontStyle.Bold)
        FonteSubTitulo = New Font("Century Gothic", 12, FontStyle.Bold)
        FonteRodape = New Font("Arial", 8)
        FonteNormal = New Font("Arial", 12)
        FonteNormalProf = New Font("Arial", inttamanhofontnormal, FontStyle.Bold)
        FonteNormaltel = New Font("Arial", 7, FontStyle.Bold)
        FonteNormaltel2 = New Font("Arial", 7)

        'define valores para linha atual e para linha da impressao
        LinhaAtual = 0
        'Cabecalho
        'e.Graphics.DrawLine(CanetaDaImpressora, 10, 10, MargemDireita, 10)

        'Imagem
        Try
            e.Graphics.DrawImage(Image.FromFile(imagem1), 20, 20)
            e.Graphics.DrawImage(Image.FromFile(imagem2), e.MarginBounds.Right - 140, 20)
            'e.Graphics.DrawString(RelatorioTitulo & System.DateTime.Today, FonteSubTitulo, Brushes.Black, MargemEsquerda + 250, 120, New StringFormat())
        Catch ex As Exception
        End Try
        'nome da Classe
        'e.Graphics.DrawString(nomeClasse, FonteTitulo, Brushes.Black, distancia7, 100, New StringFormat())
        'Linha 2
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia7, 130, MargemEsquerda + distancia6, 130)
        'campos a serem impressos: Codigo e Nome
        e.Graphics.DrawString("Boletim Geral", FonteNegrito, Brushes.Black, MetroTrackBar12.Value + 5, 130, New StringFormat())

        LinhasPorPagina = CInt(e.MarginBounds.Height / FonteNormalProf.GetHeight(e.Graphics) - 9) + 6

        '================================================================================================================
        '               Inicia da escrita na folha
        '================================================================================================================
        'While ((LinhaAtual < LinhasPorPagina) AndAlso (iii <= MetroGrid2.Rows.Count - 1))
        Dim CanetaDaImpressora3 As Brush = Brushes.Coral
        Dim nomeClasses As String = ""
        Dim ativo As Boolean = False

        For j As Integer = 0 To mgCategoria.Rows.Count - 1
            'Se campo inativo da Categoria for diferente de Verdadeiro
            If (mgCategoria.Rows(j).Cells(2).Value.ToString <> True) Or (MetroToggle4.Checked = True) Then


                'Imprime os departamentos
                Dim central As Integer = (mgCategoria.Rows(j).Cells(1).Value.ToString().Length * 10) / 2


                e.Graphics.DrawString(mgCategoria.Rows(j).Cells(1).Value.ToString(), FonteNegrito, Brushes.Black, (MargemEsquerda + 350) - central, PosicaoDaLinha - 10, New StringFormat())

                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

                e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
                e.Graphics.DrawString("Classes", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + 180, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + 280, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + 500, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

                'Analiza a Grid do Resumo (mETROfRID2 = GRID DO RESUMO)
                For i As Integer = 0 To mgResumo.Rows.Count - 1

                    'Busca Categoria de acordo com o nome da classe
                    For x As Integer = 0 To mgClasses.Rows.Count - 1
                        If mgResumo.Rows(i).Cells(2).Value.ToString = mgClasses.Rows(x).Cells(1).Value.ToString() And
                            mgClasses.Rows(x).Cells(6).Value.ToString() = mgCategoria.Rows(j).Cells(1).Value.ToString() Then
                            nomeClasses = mgClasses.Rows(x).Cells(1).Value.ToString()
                            ativo = False
                            Exit For
                        Else
                            ativo = True
                        End If
                    Next

                    If ativo = False Then



                        If mgResumo.Rows(i).Cells(2).Value.ToString = nomeClasses Then
                            If nomeClasses = "Geração Jr." Then
                                'MsgBox(nomeClasses)
                            End If
                            'obtem os valores da grid
                            Try
                                nome = mgResumo.Rows(i).Cells(2).Value.ToString
                                data = mgResumo.Rows(i).Cells(3).Value.ToString
                                tel = mgResumo.Rows(i).Cells(4).Value.ToString
                                prof = mgResumo.Rows(i).Cells(5).Value.ToString
                                ofertas = mgResumo.Rows(i).Cells(7).Value.ToString
                                ofertass = mgResumo.Rows(i).Cells(6).Value.ToString
                            Catch ex As Exception
                            End Try
                            'inicia a impressao
                            LinhaAtual += 1
                            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                            Try
                                e.Graphics.DrawString(nome.ToString(), FonteNormal, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
                                e.Graphics.DrawString(data.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
                                e.Graphics.DrawString(tel.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
                                e.Graphics.DrawString(prof.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
                                e.Graphics.DrawString(ofertass.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())


                                If ofertas.ToString() Like za Then
                                    e.Graphics.DrawString("    " & ofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                ElseIf ofertas.ToString() Like zza Then
                                    e.Graphics.DrawString("  " & ofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                ElseIf ofertas.ToString() Like zzza Then
                                    e.Graphics.DrawString(ofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                Else
                                    If ofertas.Length = 1 Then
                                        e.Graphics.DrawString("    " & ofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                    ElseIf ofertas.Length = 2 Then
                                        e.Graphics.DrawString("  " & ofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                    ElseIf ofertas.Length = 3 Then
                                        e.Graphics.DrawString(ofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                    ElseIf ofertas.ToString() Like z Then
                                        e.Graphics.DrawString("    " & ofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                    ElseIf ofertas.ToString() Like zz Then
                                        e.Graphics.DrawString("  " & ofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                    ElseIf ofertas.ToString() Like zzz Then
                                        e.Graphics.DrawString(ofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                                    End If
                                End If

                                'Linhas das classes
                                e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
                            Catch ex As Exception
                                MsgBox("sfsdf", ex.Message)
                            End Try
                            'Insere linha dos alunos (DIas)
                            e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2)
                            'Insere coluna para alunos
                            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            'Colunas dos alunos (inicio)
                            e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, distancia7, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, distancia8, PosicaoDaLinha2, distancia8, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, distancia9, PosicaoDaLinha2, distancia9, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, distancia10, PosicaoDaLinha2, distancia10, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, distancia11, PosicaoDaLinha2, distancia11, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            e.Graphics.DrawLine(CanetaDaImpressora, distancia12, PosicaoDaLinha2, distancia12, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                            'Número das linhas (Sem numeros)
                            'n += 1
                            _Talunos += data
                            _Tpresentes += tel
                            _Tausentes += prof
                            If ofertass <> "" Then
                                _Tvisitantes += ofertass
                            End If

                            _Tofertas += ofertas
                        End If
                    End If
                Next

                nomeClasses = ""
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

                'Totais
                e.Graphics.DrawString("Total", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Talunos, FonteNegrito, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Tpresentes, FonteNegrito, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Tausentes, FonteNegrito, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Tvisitantes, FonteNegrito, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())

                'e.Graphics.DrawString(_Tofertas, FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())

                If _Tofertas.ToString() Like za Then
                    e.Graphics.DrawString("    " & _Tofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                ElseIf _Tofertas.ToString() Like zza Then
                    e.Graphics.DrawString("  " & _Tofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                ElseIf _Tofertas.ToString() Like zzza Then
                    e.Graphics.DrawString(_Tofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                Else
                    If _Tofertas.ToString.Length = 1 Then
                        e.Graphics.DrawString("    " & _Tofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString.Length = 2 Then
                        e.Graphics.DrawString("  " & _Tofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString.Length = 3 Then
                        e.Graphics.DrawString(_Tofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString() Like z Then
                        e.Graphics.DrawString("    " & _Tofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString() Like zz Then
                        e.Graphics.DrawString("  " & _Tofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString() Like zzz Then
                        e.Graphics.DrawString(_Tofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    End If
                End If
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)



                e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                'e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

                'Guardando Totais
                _Ttfertas += _Tofertas
                _ttAusentes += _Tausentes
                _ttPresentes += _Tpresentes
                _ttVisitantes += _Tvisitantes
                _ttAlunos += _Talunos


                'Zeranto contadores
                _Talunos = 0
                _Tpresentes = 0
                _Tausentes = 0
                _Tvisitantes = 0
                _Tofertas = 0.00
            End If
        Next
        'MsgBox("total em Ofertas: " & _Ttfertas & vbCrLf &
        '       "Total de Presentes: " & _ttPresentes & vbCrLf &
        '       "Total de Alunos: " & _ttAlunos & vbCrLf &
        '       "Total de Ausentes: " & _ttAusentes & vbCrLf &
        '       "Total de Visitantes: " & _ttVisitantes)
        'LinhaAtual += 1
        'PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Total", FonteNegrito, Brushes.Black, (MargemEsquerda + 350) - 40, PosicaoDaLinha - 10, New StringFormat())

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
        e.Graphics.DrawString("Total de Classes", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + 180, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + 280, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + 500, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        '-----------------------------------------------------------------------------------------------------------------------------------------------
        'Impressão do total
        '-----------------------------------------------------------------------------------------------------------------------------------------------
        e.Graphics.DrawString(dtRelatorio.Text, FonteNormal, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttAlunos, FonteNormal, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttPresentes, FonteNormal, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttAusentes, FonteNormal, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttVisitantes, FonteNormal, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())
        'e.Graphics.DrawString(_Ttfertas, FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        If _Ttfertas.ToString() Like za Then
            e.Graphics.DrawString("    " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        ElseIf _Ttfertas.ToString() Like zza Then
            e.Graphics.DrawString("  " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        ElseIf _Ttfertas.ToString() Like zzza Then
            e.Graphics.DrawString(_Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        Else
            If _Ttfertas.ToString.Length = 1 Then
                e.Graphics.DrawString("    " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString.Length = 2 Then
                e.Graphics.DrawString("  " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString.Length = 3 Then
                e.Graphics.DrawString(_Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString() Like z Then
                e.Graphics.DrawString("    " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString() Like zz Then
                e.Graphics.DrawString("  " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString() Like zzz Then
                e.Graphics.DrawString(_Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            End If
        End If

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)


        '-----------------------------------------------------------------------------------------------------------------------------------------------
        '               INICIA RESUMO DOS DOMINGOS ANTERIORES
        '-----------------------------------------------------------------------------------------------------------------------------------------------



        e.Graphics.DrawString("Totais nos Domingos anteriores", FonteNegrito, Brushes.Black, (MargemEsquerda + 250) - 40, PosicaoDaLinha - 10, New StringFormat())

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
        e.Graphics.DrawString("Total de Classes", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + 180, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + 280, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + 500, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        For i As Integer = 0 To mgRelHistorico.Rows.Count - 1
            If dtRelatorio.Text <> mgRelHistorico.Rows(i).Cells(1).Value.ToString Then
                If maxDominical = i Then
                    Exit For
                Else        'Guardando Totais
                    Dim _data1 As String = mgRelHistorico.Rows(i).Cells(1).Value.ToString
                    _ttAlunos = mgRelHistorico.Rows(i).Cells(2).Value
                    _ttPresentes = mgRelHistorico.Rows(i).Cells(3).Value
                    _ttAusentes = _mgRelHistorico.Rows(i).Cells(4).Value
                    _ttVisitantes = mgRelHistorico.Rows(i).Cells(6).Value
                    _Ttfertas = mgRelHistorico.Rows(i).Cells(7).Value

                    'Zeranto contadores
                    _Talunos = 0
                    _Tpresentes = 0
                    _Tausentes = 0
                    _Tvisitantes = 0
                    _Tofertas = 0.00

                    e.Graphics.DrawString(_data1, FonteNormal, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttAlunos, FonteNormal, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttPresentes, FonteNormal, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttAusentes, FonteNormal, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttVisitantes, FonteNormal, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_Ttfertas, FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    If _Ttfertas.ToString() Like za Then
                        e.Graphics.DrawString("    " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Ttfertas.ToString() Like zza Then
                        e.Graphics.DrawString("  " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Ttfertas.ToString() Like zzza Then
                        e.Graphics.DrawString(_Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    Else
                        If _Ttfertas.ToString.Length = 1 Then
                            e.Graphics.DrawString("    " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString.Length = 2 Then
                            e.Graphics.DrawString("  " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString.Length = 3 Then
                            e.Graphics.DrawString(_Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString() Like z Then
                            e.Graphics.DrawString("    " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString() Like zz Then
                            e.Graphics.DrawString("  " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString() Like zzz Then
                            e.Graphics.DrawString(_Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        End If
                    End If

                    LinhaAtual += 1
                    PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                    e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
                End If
            End If

        Next
        '-----------------------------------------------------------------------------------------------------------------------------------------------
        'fim
        '-----------------------------------------------------------------------------------------------------------------------------------------------

        'Imprime assinaturas
        LinhaAtual += 3
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawLine(CanetaDaImpressora, 490, PosicaoDaLinha, direita - 100, PosicaoDaLinha)
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda + 100, PosicaoDaLinha, 340, PosicaoDaLinha)
        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        e.Graphics.DrawString("Assinatura", FonteNegrito, Brushes.Black, esquerda + 160, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Assinatura", FonteNegrito, Brushes.Black, esquerda + 515, PosicaoDaLinha, New StringFormat())

        LinhaAtual += 3
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Itaboraí, " & System.DateTime.Now.ToString("dd") & " de " & System.DateTime.Now.ToString("MMMM") & " de " & System.DateTime.Now.ToString("yyyy"), FonteNegrito, Brushes.Black,
                              esquerda + 270, PosicaoDaLinha, New StringFormat()) ', FonteRodape, Brushes.Black, MargemEsquerda - 60, MargemInferior, New StringFormat())


        '================================================================================================================
        '               Finaliza carga na folha
        '================================================================================================================
        'Rodape
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia7, MargemInferior, MargemEsquerda + distancia6, MargemInferior)

        e.Graphics.DrawString("Página : " & paginaAtual, FonteRodape, Brushes.Black, MargemDireita - 70, MargemInferior, New StringFormat())

        novo1 = novo1 - LinhaAtual
        'verifica se continua imprimindo
        If (LinhaAtual >= LinhasPorPagina And novo1 > 0) Then
            'If (MetroGrid5.Rows.Count - 1 < LinhaAtual) Then
            e.HasMorePages = True
            paginaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
        Else
            'start = True
            'start2 = True
            'ativa3 = True
            'iii = 0
            e.HasMorePages = False
        End If
    End Sub

    Private Sub mgCategoria_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles mgCategoria.CellClick
        indice1 = e.RowIndex
        TextBox3.Text = mgCategoria.CurrentRow().Cells(0).Value.ToString

        TextBox4.Text = mgCategoria.CurrentRow().Cells(1).Value.ToString             'Nome
        If mgCategoria.CurrentRow().Cells(2).Value.ToString() = "True" Then          'Especial
            CheckBox5.Checked = True
        Else
            CheckBox5.Checked = False
        End If

    End Sub

    Private Sub ComboBox1_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles ComboBox1.SelectionChangeCommitted
        If ComboBox1.SelectedIndex = 0 Then
            MsgBox("Escolha uma Classe!", MsgBoxStyle.Information, "Seleção incorreta")
        Else
            indices = ComboBox1.SelectedIndex
            indices -= 1
            CarregaClasses()
        End If
    End Sub

    Private Sub dtRelatorio_ValueChanged(sender As Object, e As EventArgs) Handles dtRelatorio.ValueChanged
        CargaBancoResumo()
        TotalOfertas()
        ChecaCheckBox1()
    End Sub

    Private Sub MetroButton12_Click_1(sender As Object, e As EventArgs) Handles MetroButton12.Click
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim _rData As Date


        _rData = Convert.ToDateTime(dtRelatorio.Text)
        _rData = _rData.ToString("yyyy-MM-dd")
        Try

            _ttAlunos = 0
            _ttPresentes = 0
            _ttAusentes = 0
            _ttVisitantes = 0
            _Ttfertas = 0

            For i As Integer = 0 To mgResumo.Rows.Count - 1
                _ttAlunos += mgResumo.Rows(i).Cells(3).Value
                _ttPresentes += mgResumo.Rows(i).Cells(4).Value
                _ttAusentes += mgResumo.Rows(i).Cells(5).Value
                _ttVisitantes += mgResumo.Rows(i).Cells(6).Value
                _Ttfertas += mgResumo.Rows(i).Cells(7).Value
            Next
            Dim tttoal As Integer = _ttAlunos + _ttVisitantes
            '_rData = dtRelatorio.Text
            For x As Integer = 1 To mgRelHistorico.Rows.Count - 1
                'BUSCA SE EXISTE DATA IGUAL A DATA ATUAL
                If (_rData = mgRelHistorico.Rows(x).Cells(1).Value) Then
                    'SE EXISTIR VAI A PERGUNTA
                    Dim result As Integer = MessageBox.Show("já foi registrada, deseja atualizar?", "Atualização de Classe", MessageBoxButtons.YesNo)
                    'CASO NÃO CONCORDE APENSA SAI!
                    If result = DialogResult.No Then
                        Return
                        'SE CONCORDAR ATUALIZAR, O MESMO É ATUALIZADO DE ACORDO COM A DATA.
                    Else
                        sqlconection.Open()
                        SQL2 = "UPDATE resumosdominical SET TALUNOS = '" & _ttAlunos & "', TPRESENTES = '" & _ttPresentes & "', TAUSENTES = '" & _ttAusentes & "', TVISITANTES = '" & _ttVisitantes & "', TOTAL = '" & tttoal & "', TOFERTAS = '" & _Ttfertas & "' WHERE DATA = '" & _rData & "';"

                        'Dim cmd As New MySqlCommand
                        myCommand = New MySqlCommand(SQL2, sqlconection)
                        With myCommand
                            '.CommandText = SQL
                            .CommandType = CommandType.Text
                            '.Connection = sqlconection
                            .ExecuteNonQuery()
                        End With
                        'MsgBox("Dados alterados com sucesso")
                        sqlconection.Close()
                        _ttAlunos = 0
                        _ttPresentes = 0
                        _ttAusentes = 0
                        _ttVisitantes = 0
                        _Ttfertas = 0
                        CargaBancoResumoD()
                        Return
                    End If
                End If
            Next
            sqlconection.Open()
            SQL2 = "INSERT INTO resumosdominical (DATA, TALUNOS, TPRESENTES, TAUSENTES, TVISITANTES, TOTAL, TOFERTAS) VALUES ('" & _rData & "', '" & _ttAlunos & "', '" & _ttPresentes & "', '" & _ttAusentes & "', '" & _ttVisitantes & "', '" & tttoal & "', '" & _Ttfertas & "');"

            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            'MsgBox("Dados salvos com sucesso")
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
        sqlconection.Close()
        _ttAlunos = 0
        _ttPresentes = 0
        _ttAusentes = 0
        _ttVisitantes = 0
        _Ttfertas = 0
        CargaBancoResumoD()

    End Sub

    Private Sub MetroButton17_Click(sender As Object, e As EventArgs) Handles MetroButton17.Click
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        'Dim mtniver As Boolean
        Dim srtGrade As String = ""
        srtGrade = MetroTrackBar1.Value & "," &
                   MetroTrackBar2.Value & "," &
                   MetroTrackBar3.Value & "," &
                   MetroTrackBar4.Value & "," &
                   MetroTrackBar5.Value & "," &
                   MetroTrackBar6.Value & "," &
                   MetroTrackBar7.Value & "," &
                   MetroTrackBar8.Value & "," &
                   MetroTrackBar9.Value & "," &
                   MetroTrackBar10.Value & "," &
                   MetroTrackBar11.Value & "," &
                   MetroTrackBar12.Value

        'Dim _rOfertas As String

        '_rData = Convert.ToDateTime(dtRelatorio.Text)
        '_rData = _rData.ToString("yyyy-MM-dd")
        Try

            If imagem1 <> "" Then
                imagem1 = imagem1.Replace("\", "\\")
            End If
            If imagem2 <> "" Then
                imagem2 = imagem2.Replace("\", "\\")
            End If

            sqlconection.Open()

            'Salva tudo

            SQL2 = "INSERT INTO config (ID, IMAGEM1, IMAGEM2, CORLAYOUT, TAMANHOLETRA, TAMANHONOME, RIENTACAO, GRADE, MES, DIAS, ULTIMOBACKUP, DATACONF, SPLIT1, SPLIT2, SPLIT3, SPLIT4, FUNDORESERVA, NIVER, MAXDOMINICAL, DOMINGO1, DOMINGO2, DOMINGO3, DOMINGO4, DOMINGO5, VLRZERORESU, DEPARINATIVO) values (0, '" & imagem1.ToString & "', '" & imagem2.ToString & "', '" & cbCor.SelectedItem & "', '" & cbTamanhoFonte.SelectedItem & "', '" & MetroTextBox4.Text & "', '" & MetroToggle1.CheckState & "', '" & srtGrade & "', '" & bcMes.Text & "', '07;14;21;28;', '01/05/2017', '" & dataAtual & "', '" & SplitContainer4.SplitterDistance & "', '" & SplitContainer2.SplitterDistance & "', '" & SplitContainer3.SplitterDistance & "', '" & SplitContainer1.SplitterDistance & "', '" & MetroTextBox24.Text & "', '" & mtAniversariantes.CheckState & "', '" & MetroTextBox27.Text & "', '" & MetroTextBox14.Text & "', '" & MetroTextBox15.Text & "', '" & MetroTextBox16.Text & "', '" & MetroTextBox17.Text & "', '" & MetroTextBox18.Text & "', '" & MetroToggle3.CheckState & "', '" & MetroToggle4.CheckState & "');"

            'Dim cmd As New MySqlCommand
            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                '.CommandText = SQL
                .CommandType = CommandType.Text
                '.Connection = sqlconection
                .ExecuteNonQuery()
            End With
            MsgBox("Dados salvos com sucesso")
        Catch ex As Exception
            sqlconection.Close()
            MsgBox("Erro : " & ex.Message)
        End Try
        sqlconection.Close()

    End Sub

    Private Sub mgTabelas_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles mgTabelas.CellClick
        Dim conn As New MySqlConnection
        Dim myCommandC As New MySqlCommand
        Dim myAdapterC As New MySqlDataAdapter
        Dim myDataC As New DataTable
        Dim SQLC As String

        conn = New MySqlConnection
        'conn.ConnectionString = "Server=mysql.hostinger.com.br;Database=u918624441_banco;Uid=u918624441_root;Pwd=fx74com.;"
        'conn.ConnectionString = "server=mysql.hostinger.com.br;user id=u918624441_root;password=fx74com.;database=u918624441_banco"

        conn.ConnectionString = "server=localhost;user id=root;password=;database=ebd"

        SQLC = "show columns from " & mgTabelas.CurrentRow().Cells(0).Value.ToString

        Try
            conn.Open()
            Try
                myCommandC.Connection = conn
                myCommandC.CommandText = SQLC.Trim()
                myAdapterC.SelectCommand = myCommandC
                myAdapterC.Fill(myDataC)
                mgColunas.DataSource = myDataC
                conn.Close()
            Catch ex As Exception
                MsgBox("Erro")
            End Try
        Catch ex As Exception

        End Try
        If MetroToggle2.Checked = "True" Then
            RichTextBox1.Text = "SELECT FROM " & mgTabelas.CurrentRow().Cells(0).Value.ToString
        End If

    End Sub

    Private Sub mgColunas_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles mgColunas.CellClick

        Dim texto2 As String = mgColunas.CurrentRow().Cells(0).Value.ToString
        Dim aster As String = "*T F*"
        Dim texto As String = RichTextBox1.Text
        If texto Like aster Then
            texto = texto.Replace("T F", "T " & texto2 & " F")
            'ElseIf Not texto Like aster Then
            '    texto = texto.Replace("FROM", texto2 & " FROM ")
        ElseIf Not texto Like "," Then
            texto = texto.Replace("FROM", texto2 & ", FROM ")
        ElseIf texto Like "," Then
            texto = texto.Replace("FROM", texto2 & ", FROM ")
        End If

        RichTextBox1.Text = texto
    End Sub

    Private Sub MetroButton21_Click(sender As Object, e As EventArgs) Handles MetroButton21.Click
        Dim caminho1 As String
        Dim saveFileDialog1 As New SaveFileDialog
        saveFileDialog1.Filter = "Ficheiros sql (*.sql)|*sql"
        If saveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            caminho1 = saveFileDialog1.FileName

            Dim shellcommand As String
            Try
                shellcommand = "C:\wamp\bin\mysql\mysql5.7.14\bin\mysqldump  --opt --password= --user=root --databases ebd -r " & caminho1 & ".sql"
                Shell(shellcommand)
                MsgBox("Backup Realizado com Sucesso.", MsgBoxStyle.Information)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub MetroButton19_Click(sender As Object, e As EventArgs) Handles MetroButton19.Click
        Dim caminho1 As String
        Dim OpenFileDialog1 As New OpenFileDialog
        OpenFileDialog1.Filter = "Ficheiros sql (*.sql)|*sql"
        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            caminho1 = OpenFileDialog1.FileName

            Dim shellcommand As String
            Try
                shellcommand = "C:\wamp\bin\mysql\mysql5.7.14\bin\mysql -u root -p ebd < " & caminho1
                Shell(shellcommand).ToString()
                MsgBox("Backup Realizado com Sucesso.", MsgBoxStyle.Information)
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub MetroTextBox21_Enter(sender As Object, e As EventArgs) Handles MetroTextBox21.Enter
        If MetroTextBox21.Text = "0" Then
            MetroTextBox21.Text = ""
        End If
    End Sub

    Private Sub MetroButton11_Click_2(sender As Object, e As EventArgs) Handles MetroButton11.Click
        maxDominical = MetroTextBox27.Text
        Dim objPrintPreview2 As New PrintPreviewDialog
        Dim pda As Printing.PrintDocument = New Printing.PrintDocument()
        GeraRelatorioSimples()
    End Sub
    Private Sub GeraRelatorioSimples()

        Dim objPrintPreview2 As New PrintPreviewDialog


        'Tamanho da fonte selecinada em configurações
        inttamanhofontnormal = cbTamanhoFonte.SelectedItem
        'Título do Relatório
        RelatorioTitulo = "Relatório Simples - "

        'Define os objetos printdocument e os eventos associados
        Dim pd2 As Printing.PrintDocument = New Printing.PrintDocument()

        'IMPORTANTE - definimos 2 eventos para tratar a impressão : PringPage e BeginPrint.
        AddHandler pd2.PrintPage, New Printing.PrintPageEventHandler(AddressOf Me.RelatSimples)
        AddHandler pd2.BeginPrint, New Printing.PrintEventHandler(AddressOf Me.Begin_Print)
        Try
            'define o formulário como maximizado e com Zoom
            With objPrintPreview2
                .WindowState = FormWindowState.Maximized
                .Document = pd2
                .PrintPreviewControl.Zoom = 0.65
                .Text = "Relacao de Alunos"
                .ShowDialog()
            End With
            'start = True
            'iii = 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
    'Impressão do Relatório Simples
    Private Sub RelatSimples(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        'VARIAVEIS DE POSIÇÃO (HORIZONTAL)
        Dim rTotal As Integer = 50
        Dim rAluno As Integer = 130
        Dim rPresentes As Integer = 250
        Dim rAusentes As Integer = 390
        Dim rVisitantes As Integer = 510
        Dim rTotalAlunos As Integer = 160
        Dim rTotalOfertas As Integer = 610

        'VARIAVEIS PARA POSICAO DAS LETRAS
        Dim rLTotal As Integer = rTotal
        Dim rLAluno As Integer = rAluno - 10
        Dim rLPresentes As Integer = rPresentes - 25
        Dim rLAusentes As Integer = rAusentes - 30
        Dim rLVisitantes As Integer = rVisitantes - 30
        Dim rLTotalAlunos As Integer = rTotalAlunos
        Dim rLTotalOfertas As Integer = rTotalOfertas

        _Ttfertas = 0
        _ttAusentes = 0
        _ttPresentes = 0
        _ttVisitantes = 0
        _ttAlunos = 0
        Dim z As String = "?,??"
        Dim za As String = "?,?"
        Dim zz As String = "??,??"
        Dim zza As String = "??,?"
        Dim zzz As String = "???,??"
        Dim zzza As String = "???,?"
        'Variaveis das linhas
        n = 1
        Dim LinhasPorPagina As Single = 0
        Dim PosicaoDaLinha As Single = 160
        Dim PosicaoDaLinha2 As Single = 0
        Dim LinhaAtual As Integer = 0

        'Variaveis das margens
        Dim MargemEsquerda As Single = e.MarginBounds.Left - 10
        Dim MargemSuperior As Single = e.MarginBounds.Top + 60
        Dim MargemSuperior2 As Single = e.MarginBounds.Top + 60
        Dim MargemDireita As Single = e.MarginBounds.Right + 80
        Dim MargemInferior As Single = e.MarginBounds.Bottom + 30
        Dim CanetaDaImpressora As Pen = New Pen(Color.Black, 1)
        'Dim codigo As Integer
        Dim esquerda As Integer = 35
        Dim direita As Integer = 795
        'Variaveis das fontes
        Dim FonteNegrito As Font
        Dim FonteTitulo As Font
        Dim FonteSubTitulo As Font
        Dim FonteRodape As Font
        Dim FonteNormal As Font
        Dim FonteNormalProf As Font
        Dim FonteNormaltel As Font
        Dim FonteNormaltel2 As Font
        'Dim totalPaginas As Integer

        'define efeitos em fontes
        FonteNegrito = New Font("Arial", 12, FontStyle.Bold)
        FonteTitulo = New Font("Century Gothic", 20, FontStyle.Bold)
        FonteSubTitulo = New Font("Century Gothic", 15, FontStyle.Bold)
        FonteRodape = New Font("Arial", 8)
        FonteNormal = New Font("Arial", 12)
        FonteNormalProf = New Font("Arial", inttamanhofontnormal, FontStyle.Bold)
        FonteNormaltel = New Font("Arial", 7, FontStyle.Bold)
        FonteNormaltel2 = New Font("Arial", 7)

        'define valores para linha atual e para linha da impressao
        LinhaAtual = 0
        'Cabecalho
        'e.Graphics.DrawLine(CanetaDaImpressora, 10, 10, MargemDireita, 10)

        'Imagem
        Try
            e.Graphics.DrawImage(Image.FromFile(imagem1), 20, 20)
            e.Graphics.DrawImage(Image.FromFile(imagem2), e.MarginBounds.Right - 140, 20)
            'e.Graphics.DrawString(RelatorioTitulo & System.DateTime.Today, FonteSubTitulo, Brushes.Black, MargemEsquerda + 250, 120, New StringFormat())
        Catch ex As Exception
        End Try
        e.Graphics.DrawString("Boletim Simplificado", FonteNegrito, Brushes.Black, MetroTrackBar12.Value + 5, 130, New StringFormat())

        LinhasPorPagina = CInt(e.MarginBounds.Height / FonteNormalProf.GetHeight(e.Graphics) - 9) + 6

        '================================================================================================================
        '               Inicia da escrita na folha
        '================================================================================================================
        'While ((LinhaAtual < LinhasPorPagina) AndAlso (iii <= MetroGrid2.Rows.Count - 1))
        Dim CanetaDaImpressora3 As Brush = Brushes.Coral
        Dim nomeClasses As String = ""

        For j As Integer = 0 To mgCategoria.Rows.Count - 1
            'obtem +- o centro do nome do departamento
            Dim central As Integer = (mgCategoria.Rows(j).Cells(1).Value.ToString().Length * 10) / 2
            If (mgCategoria.Rows(j).Cells(2).Value <> True) Or (MetroToggle4.Checked = True) Then
                'e.Graphics.DrawRectangle(CanetaDaImpressora, esquerda, PosicaoDaLinha, 760, 18)
                e.Graphics.DrawString(mgCategoria.Rows(j).Cells(1).Value.ToString(), FonteNegrito, Brushes.Black, (MargemEsquerda + 350) - central, PosicaoDaLinha - 10, New StringFormat())

                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

                e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
                e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + rLAluno, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + rLPresentes, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + rLAusentes, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + rLVisitantes, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + rLTotalOfertas, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

                'Analiza a Grid do Resumo (METROGRID2 = GRID DO RESUMO)
                For i As Integer = 0 To mgResumo.Rows.Count - 1
                    'Busca Categoria de acordo com o nome da classe
                    For x As Integer = 0 To mgClasses.Rows.Count - 1
                        Dim nomeResumo As String = mgResumo.Rows(i).Cells(2).Value.ToString()
                        Dim nomeClasse As String = mgClasses.Rows(x).Cells(1).Value.ToString()
                        Dim categoriaClasse As String = mgClasses.Rows(x).Cells(6).Value.ToString()
                        Dim categoriaAtual As String = mgCategoria.Rows(j).Cells(1).Value.ToString()
                        If (nomeResumo = nomeClasse) And (categoriaClasse = categoriaAtual) Then
                            nomeClasses = nomeClasse
                            Exit For
                        End If
                    Next

                    If mgResumo.Rows(i).Cells(2).Value.ToString = nomeClasses Then
                        If nomeClasses = "Geração Jr." Then
                            'MsgBox(nomeClasses)
                        End If
                        'obtem os valores da grid
                        Try
                            'Data do registro
                            nome = mgResumo.Rows(i).Cells(2).Value.ToString
                            'Nomes das classes
                            data = mgResumo.Rows(i).Cells(3).Value.ToString
                            'Total de presentes
                            tel = mgResumo.Rows(i).Cells(4).Value.ToString
                            'Total de ausrntes
                            prof = mgResumo.Rows(i).Cells(5).Value.ToString
                            'Total de vizitantes
                            ofertass = mgResumo.Rows(i).Cells(6).Value.ToString
                            'Total de ofertas
                            ofertas = mgResumo.Rows(i).Cells(7).Value.ToString
                        Catch ex As Exception
                        End Try
                        'inicia a impressao
                        'LinhaAtual += 1
                        'PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                        Try

                        Catch ex As Exception
                            MsgBox("sfsdf", ex.Message)
                        End Try
                        'Insere linha dos alunos (DIas)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2)
                        'Insere coluna para alunos
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        'Colunas dos alunos (inicio)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, distancia7, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia8, PosicaoDaLinha2, distancia8, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia9, PosicaoDaLinha2, distancia9, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia10, PosicaoDaLinha2, distancia10, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia11, PosicaoDaLinha2, distancia11, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        e.Graphics.DrawLine(CanetaDaImpressora, distancia12, PosicaoDaLinha2, distancia12, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                        'Número das linhas (Sem numeros)
                        'n += 1
                        _Talunos += data
                        _Tpresentes += tel
                        _Tausentes += prof
                        If ofertass <> "" Then
                            _Tvisitantes += ofertass
                        End If

                        _Tofertas += ofertas
                    End If

                Next

                nomeClasses = ""
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                'Totais
                e.Graphics.DrawString("Total", FonteNegrito, Brushes.Black, MargemEsquerda - rLTotal, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Talunos, FonteNegrito, Brushes.Black, MargemEsquerda + rAluno, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Tpresentes, FonteNegrito, Brushes.Black, MargemEsquerda + rPresentes, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Tausentes, FonteNegrito, Brushes.Black, MargemEsquerda + rAusentes, PosicaoDaLinha, New StringFormat())
                e.Graphics.DrawString(_Tvisitantes, FonteNegrito, Brushes.Black, MargemEsquerda + rVisitantes, PosicaoDaLinha, New StringFormat())
                'e.Graphics.DrawString(_Tofertas, FonteNegrito, Brushes.Black, MargemEsquerda + rTotalOfertas, PosicaoDaLinha, New StringFormat())

                'AQUI ARRUMAMOS AS CASAS DECIMAIS
                If _Tofertas.ToString() Like za Then
                    e.Graphics.DrawString("    " & _Tofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                ElseIf _Tofertas.ToString() Like zza Then
                    e.Graphics.DrawString("  " & _Tofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                ElseIf _Tofertas.ToString() Like zzza Then
                    e.Graphics.DrawString(_Tofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                Else
                    If _Tofertas.ToString.Length = 1 Then
                        e.Graphics.DrawString("    " & _Tofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString.Length = 2 Then
                        e.Graphics.DrawString("  " & _Tofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString.Length = 3 Then
                        e.Graphics.DrawString(_Tofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString() Like z Then
                        e.Graphics.DrawString("    " & _Tofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString() Like zz Then
                        e.Graphics.DrawString("  " & _Tofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Tofertas.ToString() Like zzz Then
                        e.Graphics.DrawString(_Tofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    End If
                End If
                'FINALIZA ARRUMACAO DAS CASAS DECIMAIS

                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

                e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
                LinhaAtual += 1
                PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                'e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

                'Guardando Totais
                _Ttfertas += _Tofertas
                _ttAusentes += _Tausentes
                _ttPresentes += _Tpresentes
                _ttVisitantes += _Tvisitantes
                _ttAlunos += _Talunos


                'Zeranto contadores
                _Talunos = 0
                _Tpresentes = 0
                _Tausentes = 0
                _Tvisitantes = 0
                _Tofertas = 0.00
            End If
        Next
        'MsgBox("total em Ofertas: " & _Ttfertas & vbCrLf &
        '       "Total de Presentes: " & _ttPresentes & vbCrLf &
        '       "Total de Alunos: " & _ttAlunos & vbCrLf &
        '       "Total de Ausentes: " & _ttAusentes & vbCrLf &
        '       "Total de Visitantes: " & _ttVisitantes)
        'LinhaAtual += 1
        'PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Total", FonteNegrito, Brushes.Black, (MargemEsquerda + 350) - 40, PosicaoDaLinha - 10, New StringFormat())

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
        e.Graphics.DrawString("Total de Classes", FonteNegrito, Brushes.Black, MargemEsquerda - rLTotal, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + rLAluno, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + rLPresentes, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + rLAusentes, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + rLVisitantes, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + rLTotalOfertas, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 2
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        'Impressão do total

        e.Graphics.DrawString(dtRelatorio.Text, FonteNormal, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttAlunos, FonteNormal, Brushes.Black, MargemEsquerda + 160, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttPresentes, FonteNormal, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttAusentes, FonteNormal, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttVisitantes, FonteNormal, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())


        'AQUI ARRUMAMOS AS CASAS DECIMAIS
        If _Ttfertas.ToString() Like za Then
            e.Graphics.DrawString("    " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        ElseIf _Ttfertas.ToString() Like zza Then
            e.Graphics.DrawString("  " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        ElseIf _Ttfertas.ToString() Like zzza Then
            e.Graphics.DrawString(_Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        Else
            If _Ttfertas.ToString.Length = 1 Then
                e.Graphics.DrawString("    " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString.Length = 2 Then
                e.Graphics.DrawString("  " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString.Length = 3 Then
                e.Graphics.DrawString(_Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString() Like z Then
                e.Graphics.DrawString("    " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString() Like zz Then
                e.Graphics.DrawString("  " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            ElseIf _Ttfertas.ToString() Like zzz Then
                e.Graphics.DrawString(_Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            End If
        End If
        'FINALIZA ARRUMACAO DAS CASAS DECIMAIS

        LinhaAtual += 2
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Totais nos Domingos anteriores", FonteNegrito, Brushes.Black, (MargemEsquerda + 250) - 40, PosicaoDaLinha - 10, New StringFormat())

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
        e.Graphics.DrawString("Total de Classes", FonteNegrito, Brushes.Black, MargemEsquerda - rLTotal, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + rLAluno, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + rLPresentes, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + rLAusentes, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + rLVisitantes, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + rLTotalOfertas, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        For i As Integer = 0 To mgRelHistorico.Rows.Count - 1
            If dtRelatorio.Text = mgRelHistorico.Rows(i).Cells(1).Value.ToString Then
                maxDominical += 1
            Else
                If maxDominical = i Then
                    Exit For
                Else            'Guardando Totais
                    Dim _data1 As String = mgRelHistorico.Rows(i).Cells(1).Value.ToString
                    _ttAlunos = mgRelHistorico.Rows(i).Cells(6).Value
                    _ttPresentes = mgRelHistorico.Rows(i).Cells(3).Value
                    _ttAusentes = _mgRelHistorico.Rows(i).Cells(4).Value
                    _ttVisitantes = mgRelHistorico.Rows(i).Cells(5).Value
                    _Ttfertas = mgRelHistorico.Rows(i).Cells(7).Value

                    'Zeranto contadores
                    _Talunos = 0
                    _Tpresentes = 0
                    _Tausentes = 0
                    _Tvisitantes = 0
                    _Tofertas = 0.00

                    e.Graphics.DrawString(_data1, FonteNormal, Brushes.Black, MargemEsquerda - rTotal, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttAlunos, FonteNormal, Brushes.Black, MargemEsquerda + rAluno, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttPresentes, FonteNormal, Brushes.Black, MargemEsquerda + rPresentes, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttAusentes, FonteNormal, Brushes.Black, MargemEsquerda + rAusentes, PosicaoDaLinha, New StringFormat())
                    e.Graphics.DrawString(_ttVisitantes, FonteNormal, Brushes.Black, MargemEsquerda + rVisitantes, PosicaoDaLinha, New StringFormat())

                    'AQUI ARRUMAMOS AS CASAS DECIMAIS
                    If _Ttfertas.ToString() Like za Then
                        e.Graphics.DrawString("    " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Ttfertas.ToString() Like zza Then
                        e.Graphics.DrawString("  " & _Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    ElseIf _Ttfertas.ToString() Like zzza Then
                        e.Graphics.DrawString(_Ttfertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                    Else
                        If _Ttfertas.ToString.Length = 1 Then
                            e.Graphics.DrawString("    " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString.Length = 2 Then
                            e.Graphics.DrawString("  " & _Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString.Length = 3 Then
                            e.Graphics.DrawString(_Ttfertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString() Like z Then
                            e.Graphics.DrawString("    " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString() Like zz Then
                            e.Graphics.DrawString("  " & _Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf _Ttfertas.ToString() Like zzz Then
                            e.Graphics.DrawString(_Ttfertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        End If
                    End If
                    'FINALIZA ARRUMACAO DAS CASAS DECIMAIS

                    LinhaAtual += 1
                    PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                    e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
                End If
            End If
        Next


        'Imprime assinaturas
        LinhaAtual += 3
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawLine(CanetaDaImpressora, 490, PosicaoDaLinha, direita - 100, PosicaoDaLinha)
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda + 100, PosicaoDaLinha, 340, PosicaoDaLinha)
        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        e.Graphics.DrawString("Assinatura", FonteNegrito, Brushes.Black, esquerda + 160, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Assinatura", FonteNegrito, Brushes.Black, esquerda + 515, PosicaoDaLinha, New StringFormat())

        LinhaAtual += 3
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Itaboraí, " & System.DateTime.Now.ToString("dd") & " de " & System.DateTime.Now.ToString("MMMM") & " de " & System.DateTime.Now.ToString("yyyy"), FonteNegrito, Brushes.Black,
                              esquerda + 270, PosicaoDaLinha, New StringFormat()) ', FonteRodape, Brushes.Black, MargemEsquerda - 60, MargemInferior, New StringFormat())


        '================================================================================================================
        '               Finaliza carga na folha
        '================================================================================================================
        'Rodape
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia7, MargemInferior, MargemEsquerda + distancia6, MargemInferior)

        e.Graphics.DrawString("Página : " & paginaAtual, FonteRodape, Brushes.Black, MargemDireita - 70, MargemInferior, New StringFormat())

        novo1 = novo1 - LinhaAtual
        'verifica se continua imprimindo
        If (LinhaAtual >= LinhasPorPagina And novo1 > 0) Then
            'If (MetroGrid5.Rows.Count - 1 < LinhaAtual) Then
            e.HasMorePages = True
            paginaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
        Else
            'start = True
            'start2 = True
            'ativa3 = True
            'iii = 0
            e.HasMorePages = False
        End If
    End Sub

    Private Sub MetroButton13_Click_1(sender As Object, e As EventArgs) Handles MetroButton13.Click
        'Armazena data atual
        dtMesAtual = Date.Now
        'Obtem apena o Mês atual da data Atual
        intMesAtual = Month(dtMesAtual)
        For i As Integer = 0 To mgTotal.Rows.Count - 1
            'Pega data do aluno e converte para "Date"
            dtNiverAluno = Date.Parse(mgTotal.Rows(i).Cells(2).Value)
            'Converte "Data" do aluno para Inteiro(Somente o Mês)
            intNiverAluno = Month(dtNiverAluno)
            'Compara os meses.
            If intNiverAluno = intMesAtual Then
                MsgBox("Feliz aniversário " & mgTotal.Rows(i).Cells(1).Value.ToString)
            End If
        Next
        MsgBox("Nenhum niver?")
    End Sub

    Private Sub MetroButton23_Click(sender As Object, e As EventArgs) Handles MetroButton23.Click
        If ColorDialog1.ShowDialog() = DialogResult.OK Then
            MetroButton23.BackColor = ColorDialog1.Color
        End If
    End Sub

    Private Sub MetroButton6_Click(sender As Object, e As EventArgs) Handles MetroButton6.Click
        Dim objPrintPreview3 As New PrintPreviewDialog
        Dim pd As Printing.PrintDocument = New Printing.PrintDocument()
        GeralTotal()
    End Sub
    Private Sub GeralTotal()

        Dim objPrintPreview3 As New PrintPreviewDialog


        'Tamanho da fonte selecinada em configurações
        inttamanhofontnormal = cbTamanhoFonte.SelectedItem
        'Título do Relatório
        RelatorioTitulo = "Relatório Geral - "

        'Define os objetos printdocument e os eventos associados
        Dim pd As Printing.PrintDocument = New Printing.PrintDocument()

        'IMPORTANTE - definimos 2 eventos para tratar a impressão : PringPage e BeginPrint.
        AddHandler pd.PrintPage, New Printing.PrintPageEventHandler(AddressOf Me.RelatGeralTotal3)
        AddHandler pd.BeginPrint, New Printing.PrintEventHandler(AddressOf Me.Begin_Print)
        Try
            'define o formulário como maximizado e com Zoom
            With objPrintPreview3
                .WindowState = FormWindowState.Maximized
                .Document = pd
                .PrintPreviewControl.Zoom = 0.65
                .Text = "Relacao de Alunos"
                .ShowDialog()
            End With
            'start = True
            'iii = 0
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub
    Private Sub RelatGeralTotal3(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        _Ttfertas = 0
        _ttAusentes = 0
        _ttPresentes = 0
        _ttVisitantes = 0
        _ttAlunos = 0
        Dim z As String = "?,??"
        Dim za As String = "?,?"
        Dim zz As String = "??,??"
        Dim zza As String = "??,?"
        Dim zzz As String = "???,??"
        Dim zzza As String = "???,?"
        'Variaveis das linhas
        n = 1
        Dim LinhasPorPagina As Single = 0
        Dim PosicaoDaLinha As Single = 160
        Dim PosicaoDaLinha2 As Single = 0
        Dim LinhaAtual As Integer = 0

        'Variaveis das margens
        Dim MargemEsquerda As Single = e.MarginBounds.Left - 10
        Dim MargemSuperior As Single = e.MarginBounds.Top + 60
        Dim MargemSuperior2 As Single = e.MarginBounds.Top + 60
        Dim MargemDireita As Single = e.MarginBounds.Right + 80
        Dim MargemInferior As Single = e.MarginBounds.Bottom + 30
        Dim CanetaDaImpressora As Pen = New Pen(Color.Black, 1)
        'Dim codigo As Integer
        Dim esquerda As Integer = 35
        Dim direita As Integer = 795
        'Variaveis das fontes
        Dim FonteNegrito As Font
        Dim FonteTitulo As Font
        Dim FonteSubTitulo As Font
        Dim FonteRodape As Font
        Dim FonteNormal As Font
        Dim FonteNormalProf As Font
        Dim FonteNormaltel As Font
        Dim FonteNormaltel2 As Font
        'Dim totalPaginas As Integer

        'define efeitos em fontes
        FonteNegrito = New Font("Arial", 12, FontStyle.Bold)
        FonteTitulo = New Font("Century Gothic", 20, FontStyle.Bold)
        FonteSubTitulo = New Font("Century Gothic", 12, FontStyle.Bold)
        FonteRodape = New Font("Arial", 8)
        FonteNormal = New Font("Arial", 12)
        FonteNormalProf = New Font("Arial", inttamanhofontnormal, FontStyle.Bold)
        FonteNormaltel = New Font("Arial", 7, FontStyle.Bold)
        FonteNormaltel2 = New Font("Arial", 7)

        'define valores para linha atual e para linha da impressao
        LinhaAtual = 0
        'Cabecalho
        'e.Graphics.DrawLine(CanetaDaImpressora, 10, 10, MargemDireita, 10)

        'Imagem
        Try
            e.Graphics.DrawImage(Image.FromFile(imagem1), 20, 20)
            e.Graphics.DrawImage(Image.FromFile(imagem2), e.MarginBounds.Right - 140, 20)
            'e.Graphics.DrawString(RelatorioTitulo & System.DateTime.Today, FonteSubTitulo, Brushes.Black, MargemEsquerda + 250, 120, New StringFormat())
        Catch ex As Exception
        End Try
        'nome da Classe
        'e.Graphics.DrawString(nomeClasse, FonteTitulo, Brushes.Black, distancia7, 100, New StringFormat())
        'Linha 2
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia7, 130, MargemEsquerda + distancia6, 130)
        'campos a serem impressos: Codigo e Nome
        e.Graphics.DrawString("Boletim Geral", FonteNegrito, Brushes.Black, MetroTrackBar12.Value + 5, 130, New StringFormat())
        'Busca Mes em configurações
        'e.Graphics.DrawString(MetroTextBox19.Text.Trim, FonteNegrito, Brushes.Black, MargemDireita - 122, 110, New StringFormat())
        'Busca dias de domingo em configurações
        'e.Graphics.DrawString(MetroTextBox14.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia1 + 10, 137, New StringFormat())
        'e.Graphics.DrawString(MetroTextBox15.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia2 + 10, 137, New StringFormat())
        'e.Graphics.DrawString(MetroTextBox16.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia3 + 10, 137, New StringFormat())
        'e.Graphics.DrawString(MetroTextBox17.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia4 + 10, 137, New StringFormat())
        'e.Graphics.DrawString(MetroTextBox18.Text.Trim, FonteNegrito, Brushes.Black, MargemEsquerda + distancia5 + 10, 137, New StringFormat())

        'Culunas do indice
        'e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, 130, MargemEsquerda + distancia1, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, 130, MargemEsquerda + distancia2, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, 130, MargemEsquerda + distancia3, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, 130, MargemEsquerda + distancia4, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, 130, MargemEsquerda + distancia5, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, 130, MargemEsquerda + distancia6, 160)

        'e.Graphics.DrawLine(CanetaDaImpressora, distancia7, 130, distancia7, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia8, 130, distancia8, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia9, 130, distancia9, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia10, 130, distancia10, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia11, 130, distancia11, 160)
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia12, 130, distancia12, 160)
        'Linha 3
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia7, 1600, MargemEsquerda + distancia6, 1600)

        LinhasPorPagina = CInt(e.MarginBounds.Height / FonteNormalProf.GetHeight(e.Graphics) - 9) + 6

        '================================================================================================================
        '               Inicia da escrita na folha
        '================================================================================================================
        'While ((LinhaAtual < LinhasPorPagina) AndAlso (iii <= MetroGrid2.Rows.Count - 1))
        Dim CanetaDaImpressora3 As Brush = Brushes.Ivory
        Dim nomeClasses As String = ""

        For j As Integer = 0 To mgCategoria.Rows.Count - 1

            'Imprime primeira linha acima dos departamentos.
            'e.Graphics.DrawLine(CanetaDaImpressora, distancia1, PosicaoDaLinha, MargemEsquerda + distancia6, PosicaoDaLinha)
            'e.Graphics.DrawLine(CanetaDaImpressora, distancia1, PosicaoDaLinha, distancia1, PosicaoDaLinha + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            'e.Graphics.DrawLine(CanetaDaImpressora, distancia10, PosicaoDaLinha, distancia10, PosicaoDaLinha + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
            'Imprime os departamentos

            Dim central As Integer = (mgCategoria.Rows(j).Cells(1).Value.ToString().Length * 10) / 2
            'e.Graphics.DrawRectangle(CanetaDaImpressora, esquerda, PosicaoDaLinha, 760, 18)
            e.Graphics.DrawString(mgCategoria.Rows(j).Cells(1).Value.ToString(), FonteNegrito, Brushes.Black, (MargemEsquerda + 100) - central, PosicaoDaLinha - 10, New StringFormat())

            LinhaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

            e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
            e.Graphics.DrawString("Classes", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + 180, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + 280, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + 500, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

            'Analiza a Grid do Resumo (mETROfRID2 = GRID DO RESUMO)
            For i As Integer = 0 To mgResumo.Rows.Count - 1

                'Busca Categoria de acordo com o nome da classe
                For x As Integer = 0 To mgClasses.Rows.Count - 1
                    If mgResumo.Rows(i).Cells(2).Value.ToString = mgClasses.Rows(x).Cells(1).Value.ToString() And
                        mgClasses.Rows(x).Cells(6).Value.ToString() = mgCategoria.Rows(j).Cells(1).Value.ToString() Then
                        nomeClasses = mgClasses.Rows(x).Cells(1).Value.ToString()
                        Exit For
                    End If
                Next

                If mgResumo.Rows(i).Cells(2).Value.ToString = nomeClasses Then
                    If nomeClasses = "Geração Jr." Then
                        'MsgBox(nomeClasses)
                    End If
                    'obtem os valores da grid
                    Try
                        nome = mgResumo.Rows(i).Cells(2).Value.ToString
                        data = mgResumo.Rows(i).Cells(3).Value.ToString
                        tel = mgResumo.Rows(i).Cells(4).Value.ToString
                        prof = mgResumo.Rows(i).Cells(5).Value.ToString
                        ofertas = mgResumo.Rows(i).Cells(7).Value.ToString
                        ofertass = mgResumo.Rows(i).Cells(6).Value.ToString
                    Catch ex As Exception
                    End Try
                    'inicia a impressao
                    LinhaAtual += 1
                    PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
                    Try
                        ' e.Graphics.DrawString("Nascimento", FonteNegrito, Brushes.Black, MargemEsquerda + 300, 137, New StringFormat())
                        ' e.Graphics.DrawString("Telefone", FonteNegrito, Brushes.Black, MargemEsquerda + 400, 137, New StringFormat())

                        'e.Graphics.DrawString(n.ToString(), FonteNormal, Brushes.Black, MetroTrackBar12.Value + 5, PosicaoDaLinha, New StringFormat())
                        e.Graphics.DrawString(nome.ToString(), FonteNormal, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
                        e.Graphics.DrawString(data.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
                        e.Graphics.DrawString(tel.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
                        e.Graphics.DrawString(prof.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
                        e.Graphics.DrawString(ofertass.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())


                        If ofertas.ToString() Like za Then
                            e.Graphics.DrawString("    " & ofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf ofertas.ToString() Like zza Then
                            e.Graphics.DrawString("  " & ofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        ElseIf ofertas.ToString() Like zzza Then
                            e.Graphics.DrawString(ofertas.ToString() & "0", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                        Else
                            If ofertas.Length = 1 Then
                                e.Graphics.DrawString("    " & ofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                            ElseIf ofertas.Length = 2 Then
                                e.Graphics.DrawString("  " & ofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                            ElseIf ofertas.Length = 3 Then
                                e.Graphics.DrawString(ofertas.ToString() & ",00", FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                            ElseIf ofertas.ToString() Like z Then
                                e.Graphics.DrawString("    " & ofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                            ElseIf ofertas.ToString() Like zz Then
                                e.Graphics.DrawString("  " & ofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                            ElseIf ofertas.ToString() Like zzz Then
                                e.Graphics.DrawString(ofertas.ToString(), FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
                            End If
                        End If

                        'Linhas das classes
                        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
                    Catch ex As Exception
                        MsgBox("sfsdf", ex.Message)
                    End Try
                    'Insere linha dos alunos (DIas)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2)
                    'Insere coluna para alunos
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia1, PosicaoDaLinha2, MargemEsquerda + distancia1, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia2, PosicaoDaLinha2, MargemEsquerda + distancia2, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia3, PosicaoDaLinha2, MargemEsquerda + distancia3, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia4, PosicaoDaLinha2, MargemEsquerda + distancia4, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia5, PosicaoDaLinha2, MargemEsquerda + distancia5, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, MargemEsquerda + distancia6, PosicaoDaLinha2, MargemEsquerda + distancia6, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    'Colunas dos alunos (inicio)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia7, PosicaoDaLinha2, distancia7, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia8, PosicaoDaLinha2, distancia8, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia9, PosicaoDaLinha2, distancia9, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia10, PosicaoDaLinha2, distancia10, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia11, PosicaoDaLinha2, distancia11, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    e.Graphics.DrawLine(CanetaDaImpressora, distancia12, PosicaoDaLinha2, distancia12, PosicaoDaLinha2 + (FonteNormalProf.GetHeight(e.Graphics)) - 1)
                    'Número das linhas (Sem numeros)
                    'n += 1
                    _Talunos += data
                    _Tpresentes += tel
                    _Tausentes += prof
                    If ofertass <> "" Then
                        _Tvisitantes += ofertass
                    End If

                    _Tofertas += ofertas
                End If

            Next

            nomeClasses = ""
            LinhaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

            'Totais
            e.Graphics.DrawString("Total", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_Talunos, FonteNegrito, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_Tpresentes, FonteNegrito, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_Tausentes, FonteNegrito, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_Tvisitantes, FonteNegrito, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())

            e.Graphics.DrawString(_Tofertas, FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
            LinhaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)



            e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
            LinhaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            'e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

            'Guardando Totais
            _Ttfertas += _Tofertas
            _ttAusentes += _Tausentes
            _ttPresentes += _Tpresentes
            _ttVisitantes += _Tvisitantes
            _ttAlunos += _Talunos


            'Zeranto contadores
            _Talunos = 0
            _Tpresentes = 0
            _Tausentes = 0
            _Tvisitantes = 0
            _Tofertas = 0.00

        Next
        'MsgBox("total em Ofertas: " & _Ttfertas & vbCrLf &
        '       "Total de Presentes: " & _ttPresentes & vbCrLf &
        '       "Total de Alunos: " & _ttAlunos & vbCrLf &
        '       "Total de Ausentes: " & _ttAusentes & vbCrLf &
        '       "Total de Visitantes: " & _ttVisitantes)
        'LinhaAtual += 1
        'PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Total", FonteNegrito, Brushes.Black, (MargemEsquerda + 350) - 40, PosicaoDaLinha - 10, New StringFormat())

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
        e.Graphics.DrawString("Total de Classes", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + 180, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + 280, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + 500, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        'Impressão do total

        e.Graphics.DrawString(dtRelatorio.Text, FonteNormal, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttAlunos, FonteNormal, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttPresentes, FonteNormal, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttAusentes, FonteNormal, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_ttVisitantes, FonteNormal, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString(_Ttfertas, FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Totais nos Domingos anteriores", FonteNegrito, Brushes.Black, (MargemEsquerda + 250) - 40, PosicaoDaLinha - 10, New StringFormat())

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.FillRectangle(CanetaDaImpressora3, esquerda, PosicaoDaLinha, 760, 18)
        e.Graphics.DrawString("Total de Classes", FonteNegrito, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Alunos", FonteNegrito, Brushes.Black, MargemEsquerda + 180, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Presentes", FonteNegrito, Brushes.Black, MargemEsquerda + 280, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ausentes", FonteNegrito, Brushes.Black, MargemEsquerda + 400, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Vizitantes", FonteNegrito, Brushes.Black, MargemEsquerda + 500, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Ofetras", FonteNegrito, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        For i As Integer = 0 To mgRelHistorico.Rows.Count - 1
            'Guardando Totais
            Dim _data1 As String = mgRelHistorico.Rows(i).Cells(1).Value.ToString
            _ttAlunos = mgRelHistorico.Rows(i).Cells(2).Value
            _ttPresentes = mgRelHistorico.Rows(i).Cells(3).Value
            _ttAusentes = _mgRelHistorico.Rows(i).Cells(4).Value
            _ttVisitantes = mgRelHistorico.Rows(i).Cells(5).Value
            _Ttfertas = mgRelHistorico.Rows(i).Cells(6).Value

            'Zeranto contadores
            _Talunos = 0
            _Tpresentes = 0
            _Tausentes = 0
            _Tvisitantes = 0
            _Tofertas = 0.00

            e.Graphics.DrawString(_data1, FonteNormal, Brushes.Black, MargemEsquerda - 50, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_ttAlunos, FonteNormal, Brushes.Black, MargemEsquerda + 200, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_ttPresentes, FonteNormal, Brushes.Black, MargemEsquerda + 310, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_ttAusentes, FonteNormal, Brushes.Black, MargemEsquerda + 430, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_ttVisitantes, FonteNormal, Brushes.Black, MargemEsquerda + 530, PosicaoDaLinha, New StringFormat())
            e.Graphics.DrawString(_Ttfertas, FonteNormal, Brushes.Black, MargemEsquerda + 610, PosicaoDaLinha, New StringFormat())

            LinhaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)

        Next


        'Imprime assinaturas
        LinhaAtual += 3
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawLine(CanetaDaImpressora, 490, PosicaoDaLinha, direita - 100, PosicaoDaLinha)
        e.Graphics.DrawLine(CanetaDaImpressora, esquerda + 100, PosicaoDaLinha, 340, PosicaoDaLinha)
        LinhaAtual += 1
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
        e.Graphics.DrawString("Assinatura", FonteNegrito, Brushes.Black, esquerda + 160, PosicaoDaLinha, New StringFormat())
        e.Graphics.DrawString("Assinatura", FonteNegrito, Brushes.Black, esquerda + 515, PosicaoDaLinha, New StringFormat())

        LinhaAtual += 3
        PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)

        e.Graphics.DrawString("Itaboraí, " & System.DateTime.Now.ToString("dd") & " de " & System.DateTime.Now.ToString("MMMM") & " de " & System.DateTime.Now.ToString("yyyy"), FonteNegrito, Brushes.Black,
                              esquerda + 270, PosicaoDaLinha, New StringFormat()) ', FonteRodape, Brushes.Black, MargemEsquerda - 60, MargemInferior, New StringFormat())


        '================================================================================================================
        '               Finaliza carga na folha
        '================================================================================================================
        'Rodape
        'e.Graphics.DrawLine(CanetaDaImpressora, distancia7, MargemInferior, MargemEsquerda + distancia6, MargemInferior)

        e.Graphics.DrawString("Página : " & paginaAtual, FonteRodape, Brushes.Black, MargemDireita - 70, MargemInferior, New StringFormat())

        novo1 = novo1 - LinhaAtual
        'verifica se continua imprimindo
        If (LinhaAtual >= LinhasPorPagina And novo1 > 0) Then
            'If (MetroGrid5.Rows.Count - 1 < LinhaAtual) Then
            e.HasMorePages = True
            paginaAtual += 1
            PosicaoDaLinha = MargemSuperior + (LinhaAtual * FonteNormal.GetHeight)
            e.Graphics.DrawLine(CanetaDaImpressora, esquerda, PosicaoDaLinha, direita, PosicaoDaLinha)
        Else
            'start = True
            'start2 = True
            'ativa3 = True
            'iii = 0
            e.HasMorePages = False
        End If
    End Sub


    Private Sub CheckedListBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedValueChanged
        If CheckedListBox1.SelectedIndex <= 0 Or CheckedListBox1.SelectedItem Like "*(X)" Then
            MetroTextBox20.Text = "0"
            Return
        Else
            _rCLasse = CheckedListBox1.SelectedItem()
            For i As Integer = 0 To mgClasses.Rows.Count - 1
                If _rCLasse = mgClasses.Rows(i).Cells(1).Value.ToString Then
                    _rCLasseId = mgClasses.Rows(i).Cells(0).Value
                End If
            Next
            '_rCLasseId = CheckedListBox1.SelectedIndex
            indices = CheckedListBox1.SelectedIndex
            indices -= 1

            CarregaClassesResumo()
            MetroTextBox20.Text = totalDeAlunos.ToString

        End If
    End Sub

    Private Sub MetroTextBox26_Click(sender As Object, e As EventArgs) Handles MetroTextBox26.Click
        If MetroTextBox26.Text = "0" Then
            MetroTextBox26.Text = ""
        Else
            Return
        End If
    End Sub
    'Insere novo departamento
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim inativo As String = CheckBox5.CheckState
        If CheckBox5.CheckState = "1" Then
            inativo = "True"
        Else
            inativo = "False"
        End If

        If TextBox4.Text = "" Then
            MsgBox("Preencha o campo!")
            Return
        End If

        For x As Integer = 0 To mgCategoria.Rows.Count - 1
            If TextBox4.Text = mgCategoria.Rows(x).Cells(1).Value.ToString Then
                MsgBox(TextBox4.Text & ", este departamento já existe!")
                Return
                Exit For
            End If
        Next
        Try
            sqlconection.Open()
            SQL2 = "INSERT INTO CATEGORIA (ID, NOME, INATIVO) values (NULL, '" &
                TextBox4.Text & "', '" &
                inativo & "');"


            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With
            MsgBox("Cadastrado alterado com sucesso")
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try

        sqlconection.Close()
        CarregaTudo()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim myCommand As New MySqlCommand
        Dim SQL2 As String
        Dim inativo As String
        If CheckBox5.CheckState = "1" Then
            inativo = "True"
        Else
            inativo = "False"
        End If
        Dim catclasses As String = ""
        If inativo = "True" Then
            For i As Integer = 0 To mgClasses.Rows.Count - 1
                If TextBox4.Text = mgClasses.Rows(i).Cells(6).Value.ToString Then
                    If catclasses = "" Then
                        catclasses = mgClasses.Rows(i).Cells(1).Value.ToString()
                    Else
                        catclasses = catclasses & ", " & mgClasses.Rows(i).Cells(1).Value.ToString()
                    End If
                End If
            Next
            If catclasses <> "" Then
                MsgBox("Este departamento esta vinculado a(s) segunte(s) classe(s):" & vbCrLf & catclasses & ".", MsgBoxStyle.OkOnly, "Informação")
            End If
        End If
        Try

            sqlconection.Open()
            SQL2 = ("Update CATEGORIA SET NOME = '" & TextBox4.Text & "', INATIVO = '" & inativo & "' where id = '" & TextBox3.Text & "';")

            myCommand = New MySqlCommand(SQL2, sqlconection)
            With myCommand
                .CommandType = CommandType.Text
                .ExecuteNonQuery()
            End With
            MsgBox("Cadastrado alterado com sucesso")
        Catch ex As Exception
            MsgBox("Erro : " & ex.Message)
        End Try
        sqlconection.Close()
        CarregaTudo()
    End Sub

    Private Sub MetroTextBox26_Leave(sender As Object, e As EventArgs) Handles MetroTextBox26.Leave
        If IsNumeric(MetroTextBox26.Text) Then
            If (MetroTextBox26.Text < 0) Then
                MsgBox("Valor menor que 0!", vbOKOnly, "Informação")
                MetroTextBox26.Text = "0"
            Else
                Return
            End If
        Else
            MetroTextBox26.Text = "0"
        End If
    End Sub

    Private Sub MetroTextBox23_Leave(sender As Object, e As EventArgs) Handles MetroTextBox23.Leave
        If IsNumeric(MetroTextBox23.Text) Then
            If (MetroTextBox23.Text < 0) Then
                MsgBox("Valor menor que 0!", vbOKOnly, "Informação")
                MetroTextBox23.Text = "0,00"
            Else
                Return
            End If
        Else
            MetroTextBox23.Text = "0,00"
        End If
    End Sub

    Private Sub MetroTextBox24_Leave(sender As Object, e As EventArgs) Handles MetroTextBox24.Leave
        If IsNumeric(MetroTextBox24.Text) Then
            If (MetroTextBox24.Text < 0) Then
                MsgBox("Valor menor que 0!", vbOKOnly, "Informação")
                MetroTextBox24.Text = "0,00"
            Else
                TotalOfertas()
            End If
        Else
            MetroTextBox24.Text = "0,00"
        End If

    End Sub

    'Private Sub MaskedTextBox4_Leave(sender As Object, e As EventArgs)
    '    If MaskedTextBox4.Text <> "" Then
    '        If MaskedTextBox4.Text Like "/" Then
    '        Else
    '            Try
    '                Dim dt As DateTime = CDate(MaskedTextBox4.Text).ToString("dd/MM/yyyy")
    '                dt = Convert.ToDateTime(dt)
    '                Dim ts As TimeSpan = DateTime.Today.Subtract(dt)

    '                MetroLabel20.Text = New DateTime(ts.Ticks).ToString("yy") - 1
    '                MetroLabel20.Refresh()
    '            Catch ex As Exception

    '            End Try
    '        End If
    '    End If
    'End Sub

    ''Private Sub MaskedTextBox3_Leave(sender As Object, e As EventArgs)
    ''    If MaskedTextBox3.Text <> "" Then
    ''        If MaskedTextBox3.Text Like "/" Then
    ''        Else
    ''            Try
    ''                Dim dt As DateTime = CDate(MaskedTextBox3.Text).ToString("dd/MM/yyyy")
    ''                dt = Convert.ToDateTime(dt)
    ''                Dim ts As TimeSpan = DateTime.Today.Subtract(dt)

    ''                MetroLabel21.Text = New DateTime(ts.Ticks).ToString("yy") - 1
    ''                MetroLabel21.Refresh()
    ''            Catch ex As Exception

    ''            End Try
    ''        End If
    ''    End If
    ''End Sub



    Private Sub MetroCheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles MetroCheckBox1.CheckedChanged
        If MetroCheckBox1.Checked = "True" Then
            MetroComboBox2.Enabled = True
        Else
            MetroComboBox2.Enabled = False

        End If
    End Sub

    Private Sub MetroCheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles MetroCheckBox2.CheckedChanged
        If MetroCheckBox2.Checked = "True" Then
            MetroComboBox2.Enabled = True
        Else
            MetroComboBox2.Enabled = False

        End If
    End Sub

    Private Sub MaskedTextBox1_Enter(sender As Object, e As EventArgs) Handles MaskedTextBox1.Enter
        MetroComboBox2.Enabled = True
    End Sub

    Private Sub MetroButton22_Click(sender As Object, e As EventArgs) Handles MetroButton22.Click
        Dim myCommand As New MySqlCommand
        Dim myAdapter As New MySqlDataAdapter
        Dim myData As New DataTable
        Dim SQL As String

        SQL = "CREATE DATATABLE TESTE;"
        Try
            Dim connectionString As String = String.Format("Data Source=(LocalDB)\v11.0;Initial Catalog=master;Integrated Security=True")
            Using connection As New SqlConnection(connectionString)

                connection.Open()
                Dim cmd As SqlCommand = connection.CreateCommand()


                'DetachDatabase(dbName)

                cmd.CommandText = String.Format("CREATE DATABASE {0} ON (NAME = N'{0}', FILENAME = '{1}')", "teste", "arquivo.mdb")
                cmd.ExecuteNonQuery()
            End Using
            Try

                myCommand = New MySqlCommand(SQL, sqlconection)
                With myCommand
                    .CommandType = CommandType.Text
                End With
                With myAdapter
                    .SelectCommand = myCommand
                    .Fill(myData)
                End With
                mgClasses.DataSource = myData
                sqlconection.Close()
                'MsgBox("Dados dos alunos atualizados com sucesso!", MsgBoxStyle.Information, "Atualização")
            Catch myerro As MySqlException
                MsgBox("Erro de leitura no banco de dados : " & myerro.Message)
            End Try
        Catch myerro As MySqlException
            MessageBox.Show("Erro ao conectar com o Banco de dados : " & myerro.Message)
        Finally
            sqlconection.Dispose()
        End Try
    End Sub

    Private Sub MetroTextBox4_Leave(sender As Object, e As EventArgs) Handles MetroTextBox4.Leave
        TextBox9.MaxLength = MetroTextBox4.Text
    End Sub

    Private Sub MetroButton24_Click(sender As Object, e As EventArgs) Handles MetroButton24.Click
        lblInfo.Visible = False
        RestoreMysql()
        lblInfo.Visible = False
    End Sub
    Sub RestoreMysql()
        LerArquivoSQL(AbrirArquivoSQL(True))
    End Sub
    Function AbrirArquivoSQL(ByVal tipo As Boolean) As String
        Try
            If tipo = True Then
                Dim AbrirComo As OpenFileDialog = New OpenFileDialog()
                Dim caminho As DialogResult

                AbrirComo.Title = "Abrir como"
                AbrirComo.FileName = "Nome Arquivo"
                AbrirComo.Filter = "Arquivos Textos (*.sql)|*.sql"
                caminho = AbrirComo.ShowDialog
                AbrirArquivoSQL = AbrirComo.FileName

                If AbrirArquivoSQL = Nothing Then
                    MessageBox.Show("Arquivo Inválido", "Abrir Arquivo", MessageBoxButtons.OK)
                Else
                    Return AbrirArquivoSQL
                End If
            Else
                Dim AbrirComo As SaveFileDialog = New SaveFileDialog()
                Dim caminho As DialogResult

                Dim strDate As String = Date.Now.ToShortDateString   'Prepend file with date for dated backups
                Dim fileName As String = lstBancos.FocusedItem.Text & "_" & strDate.Replace("/", "-") & "_" & Now.ToShortTimeString.Replace(":", "_").ToString & ".sql"

                AbrirComo.Title = "Salvar como"
                AbrirComo.FileName = fileName '& ".sql"
                AbrirComo.Filter = "Arquivos Textos (*.sql)|*.sql"
                caminho = AbrirComo.ShowDialog
                AbrirArquivoSQL = AbrirComo.FileName

                If AbrirArquivoSQL = Nothing Then
                    MessageBox.Show("Arquivo Inválido", "Salvar Como", MessageBoxButtons.OK)
                Else
                    Return AbrirArquivoSQL
                End If
            End If
        Catch ex As Exception

        End Try
        Return Nothing
    End Function
    Friend WithEvents myScript As MySql.Data.MySqlClient.MySqlScript
    Private MyConString As String
    Sub LerArquivoSQL(ByVal arqv As String)
        Dim fluxoTexto As IO.StreamReader
        Dim linhaTexto As String
        Dim Linha As String = Nothing
        Dim Linhas() As String = Nothing
        If IO.File.Exists(arqv) Then
            fluxoTexto = New IO.StreamReader(arqv)
            linhaTexto = fluxoTexto.ReadToEnd
            While linhaTexto <> Nothing
                Linha &= linhaTexto & vbCrLf
                linhaTexto = fluxoTexto.ReadLine
                Linhas = Split(Linha, "/*<>*/")
            End While
            Dim p As Integer
            Dim Con As MySqlConnection = New MySqlConnection(MyConString)
            Dim ShowCreateCommand As MySqlCommand = sqlconection.CreateCommand
            lblInfo.Visible = True
            lblInfo.Text = "Processo de Restauração Iniciado !!!"
            lblProcesso.Text = "Aguarde !!! Iniciado Processo de Restauração"
            Application.DoEvents()
            Cursor.Current = Cursors.WaitCursor
            Dim script As String = Linha

            Try
                myScript = New MySqlScript(script)
                myScript.Connection = sqlconection 'Con
                myScript.Execute()
                lblProcesso.Text = "Restauração Concluída !!!"
                lblInfo.Text = lblProcesso.Text
                Application.DoEvents()
            Catch ex As MySqlException
                MsgBox(ex.Message.ToString)
            End Try

            Cursor.Current = Cursors.Default

            For p = 1 To Linhas.Length - 1
                If 1 = 2 Then
                    myScript = New MySqlScript(script)
                    myScript.Connection = Con
                    myScript.Execute()
                Else
                    ShowCreateCommand.CommandText = Linhas(p)
                    ShowCreateCommand.ExecuteNonQuery()
                End If
            Next
            fluxoTexto.Close()
        Else
            MessageBox.Show("Arquivo não existe")
        End If
    End Sub

    Private Sub MetroButton25_Click(sender As Object, e As EventArgs) Handles MetroButton25.Click
        Dim teste As Form2 = New Form2

        teste.Show()


    End Sub
End Class
