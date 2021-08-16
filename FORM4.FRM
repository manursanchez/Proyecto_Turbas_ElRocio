VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reparar Base de Datos"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Clientes"
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Facturas"
      Top             =   120
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comenzar la reparación"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Progreso de la reparación:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim contador As Integer
    Command1.Enabled = False
    Command2.Enabled = False
    If Not Data1.Recordset.BOF Then
        Label1.Caption = "Comenzando reparación de auxiliar de facturas..."
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
        Loop
    End If
    If Not Data2.Recordset.BOF Then
        Label1.Caption = "Comenzando reparación de auxiliar de clientes..."
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.Delete
            Data2.Recordset.MoveNext
        Loop
    End If
    Label1.Caption = "Reparacion terminada pulse salir para continuar..."
    Command1.Enabled = True
    Command2.Enabled = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
