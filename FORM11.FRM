VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introducir cliente en factura"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form11.frx":030A
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "Form11.frx":031E
      TabIndex        =   5
      Top             =   1680
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ver  listado >>"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      MaxLength       =   9
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Introduce DNI del cliente"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ver As Boolean 'Variable para desplegar
'o replegar listado de clientes
Private Sub Command1_Click()
    If ver = False Then
        Form11.Height = 5775
        ver = True
        Form11.Command1.Caption = "Esconder listado <<"
    Else
        If ver = True Then
            Form11.Height = 1935
            ver = False
            Form11.Command1.Caption = "Ver listado >>"
        End If
    End If
End Sub

Private Sub Command2_Click()
    Dim consulta As String
    If Form11.Text1.Text = "" Then
        MsgBox "No hay ningun DNI en el cuadrito.Por favor introduzca un DNI para poder introducir en factura un cliente", vbExclamation, "Sin datos"
        Unload Form11
        Load Form11
        Form11.Show vbModal
    Else
        consulta = "select * from clientes where dni=""" + Form11.Text1.Text + """"
        Form11.Data1.RecordSource = consulta
        Form11.Data1.Refresh
        If Form11.Data1.Recordset.RecordCount < 1 Then
            MsgBox "El cliente introducido no existe,por favor compruebe el número de DNI", vbExclamation, "Error de datos"
            Unload Form11
            Load Form11
            Form11.Show vbModal
        Else
            identificacion = Form11.Text1.Text
            Unload Me
            Load Form12
            Form12.Show vbModal
        End If
    End If
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ver = False 'La inicializamos a false
    'Por que por defecto no estara desplegada
    'la lista de clientes
End Sub
