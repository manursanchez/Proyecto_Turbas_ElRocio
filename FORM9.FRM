VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva factura"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Facturas"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Facturas"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      MaxLength       =   6
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Introduce número de factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim consulta As String
    If Form9.Text1.Text = "" Then
        MsgBox "Imposible crear factura.Tiene que introducir algun valor", vbExclamation, "Error de datos"
    Else
        consulta = "select * from facturas where nfac=""" + Form9.Text1.Text + """"
        Form9.Data1.RecordSource = consulta
        Form9.Data1.Refresh
        If Form9.Data1.Recordset.RecordCount < 1 Then
            numfactura = Form9.Text1.Text
            Unload Form9
            Load Form11
            Form11.Show vbModal
        Else
            MsgBox "El número de factura ya existe,eso significa que ya existe una factura con ese numero y no puede haber dos facturas con numeros iguales", vbExclamation, "Error de datos"
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


