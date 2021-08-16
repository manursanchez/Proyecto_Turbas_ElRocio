VERSION 5.00
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminación masiva de facturas"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "Eliminacion de facturas.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Facturas"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Comenzar la eliminación"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Estado...."
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Data1.Recordset.BOF Then
        MsgBox "Imposible eliminar por que no hay facturas", vbInformation, "Información de sistema"
    Else
        Label2.Caption = "Eliminando..."
        Data1.Recordset.MoveFirst
        Do While Not Data1.Recordset.EOF
            Data1.Recordset.Delete
            Data1.Recordset.MoveNext
        Loop
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.Delete
            Data2.Recordset.MoveNext
        Loop
        Label2.Caption = "Eliminacion terminada"
        MsgBox "Facturas eliminadas satisfactoriamente", vbInformation, "Informacón de sistema"
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
