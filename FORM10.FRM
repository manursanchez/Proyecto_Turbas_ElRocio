VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar elementos de factura"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Totales"
      Height          =   1215
      Left            =   3480
      TabIndex        =   10
      Top             =   480
      Width           =   2655
      Begin VB.Label Label21 
         Caption         =   "€+IVA"
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
         Left            =   1440
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "€uros"
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
         Left            =   1440
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TERMINADO"
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ELIMINAR"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Utiliza las flechas para moverte a través de los componentes de la factura"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Facturas"
      Top             =   2520
      Width           =   5985
   End
   Begin VB.Label Label14 
      BackColor       =   &H00E0E0E0&
      DataField       =   "peuros"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      DataField       =   "peurosiva"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      DataField       =   "concepto"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      DataField       =   "cantidad"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "€+IVA"
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
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Precio en €"
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
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cantidad"
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Concepto"
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
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub Command1_Click()
'    Dim opc As Integer
'    opc = MsgBox("¿Eliminar este elemento?", vbYesNo)
'    If opc = vbYes Then
'        If Data1.Recordset.BOF Then
'            Form10.Label15.Caption = 0
'            Form10.Label17.Caption = 0
'            MsgBox "Ya no hay ningun elemento en factura para eliminar, añada elementos a factura", vbExclamation, "Error de datos"
'        Else
'            Form10.Label15.Caption = Format(Form10.Label15.Caption - Form10.Label14.Caption, "######.##")
'            Form10.Label17.Caption = Format(Form10.Label17.Caption - Form10.Label12.Caption, "######.##")
'            Form10.Data1.Recordset.Delete
'            Form10.Data1.Refresh
'        End If
'    End If
'End Sub

'Private Sub Command2_Click()
    'Actualizo los labels y los texts con los nuevos valores
' Form12.euros.Text = Form10.Label15.Caption
'    Form12.eurosiva.Text = Form10.Label17.Caption
'    'Actualizo las variables de las operaciones matematicas
 '   teuros = Val(Form10.Label15.Caption)
 '   teurosiva = Val(Form10.Label17.Caption)
 '   Form12.Data3.Refresh
 '   Unload Me
'End Sub

'Private Sub Form_Load()
 '   Form10.Label15.Caption = Form12.euros.Text
 '   Form10.Label17.Caption = Form12.eurosiva.Text
'End Sub
Private Sub Command1_Click()

End Sub
