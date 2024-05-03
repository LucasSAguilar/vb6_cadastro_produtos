VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13425
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   13425
   Begin VB.CommandButton btn_enviar 
      Caption         =   "Enviar"
      Height          =   615
      Left            =   8160
      TabIndex        =   6
      Top             =   5880
      Width           =   4575
   End
   Begin VB.ComboBox cbb_disponivel 
      Height          =   315
      Left            =   8160
      TabIndex        =   5
      Top             =   4800
      Width           =   4575
   End
   Begin VB.TextBox ipt_valor 
      Height          =   495
      Left            =   8160
      TabIndex        =   4
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox ipt_nome 
      Height          =   495
      Left            =   8160
      TabIndex        =   3
      Top             =   2400
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid listaProdutos 
      Height          =   3495
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      ForeColor       =   -2147483647
      BackColorBkg    =   4210752
      MergeCells      =   4
      FormatString    =   ""
   End
   Begin VB.CommandButton btn_findData 
      Caption         =   "Buscar dados"
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   5880
      Width           =   4095
   End
   Begin VB.Label Disponível 
      Caption         =   "Disponível"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Valor"
      Height          =   255
      Left            =   8160
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      Height          =   255
      Left            =   8160
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   6360
      X2              =   6360
      Y1              =   2040
      Y2              =   6480
   End
   Begin VB.Label titulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Produtos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5775
      TabIndex        =   0
      Top             =   840
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function configurarConexao() As ADODB.Connection
    Dim conn As ADODB.Connection

    ' String para se conectar no banco de dados
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Lucas\Programacao\MRC\db.mdb"

    Set configurarConexao = conn

End Function

Function ColetarTodosDados(conn As ADODB.Connection) As ADODB.recordset
    Dim querySQL As String
    Dim recordset As ADODB.recordset
    
    querySQL = "SELECT ID, nome, valor, disponivel FROM produtos"
    
    Set recordset = New ADODB.recordset
    recordset.Open querySQL, conn
    
            
    
    Set ColetarTodosDados = recordset

End Function

Sub EnviarDados(conexao As ADODB.Connection, nome As String, valor As Currency, disponivel As String)
    Dim sqlQuery As String
    Dim comando As ADODB.Command
    sqlQuery = "INSERT INTO produtos (nome, valor, disponivel) VALUES (?,?,?)"
    
    Set comando = New ADODB.Command
    comando.ActiveConnection = conexao
    comando.CommandText = sqlQuery
    
    comando.Parameters.Append comando.CreateParameter("nome_data", adVarChar, adParamInput, 255, nome)
    comando.Parameters.Append comando.CreateParameter("valor_data", adCurrency, adParamInput, , valor)
    comando.Parameters.Append comando.CreateParameter("disponivel_data", adVarChar, adParamInput, 3, disponivel)
    
    comando.Execute
    

End Sub

Private Sub btn_enviar_Click()
Dim conexao As ADODB.Connection
Dim nome As String
Dim valor As Currency
Dim disponivel As String

Set conexao = configurarConexao()

nome = ipt_nome.Text
valor = ipt_valor.Text
disponivel = cbb_disponivel.Text

conexao.Open
EnviarDados conexao, nome, valor, disponivel
conexao.Close
ipt_nome.Text = ""
ipt_valor.Text = ""
cbb_disponivel = ""

End Sub

Private Sub btn_findData_Click()
Dim conexao As ADODB.Connection
Dim dados As ADODB.recordset

Set conexao = configurarConexao()
conexao.Open

Set dados = ColetarTodosDados(conexao)
InicializarTabela dados
conexao.Close

End Sub
Sub InicializarTabela(dados As ADODB.recordset)
    Dim totalRegistros As Integer
    
    Do While Not dados.EOF
        totalRegistros = totalRegistros + 1
        dados.MoveNext
    Loop
    
    
    For Col = 0 To dados.Fields.Count - 1
        listaProdutos.TextMatrix(0, Col) = dados.Fields(Col).Name
    Next Col

    listaProdutos.Rows = totalRegistros + 1
    InserirDadosTabela dados, totalRegistros - 1
   
End Sub

Sub InserirDadosTabela(dados As ADODB.recordset, totalRegistros As Integer)
Dim linha As Integer
Dim coluna As Integer

    linha = 1
    
    dados.MoveFirst
    Do While Not dados.EOF
            For coluna = 0 To dados.Fields.Count - 1
                listaProdutos.TextMatrix(linha, coluna) = dados.Fields(coluna).Value
            Next coluna
        linha = linha + 1
        dados.MoveNext
    Loop
End Sub

Private Sub Form_Load()

    cbb_disponivel.AddItem ("Sim")
    cbb_disponivel.AddItem ("Não")

End Sub
