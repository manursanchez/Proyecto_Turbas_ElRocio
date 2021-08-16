VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Factura recuperada"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "Form14.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form14.frx":030A
      Height          =   1815
      Left            =   240
      OleObjectBlob   =   "Form14.frx":031E
      TabIndex        =   23
      Top             =   2640
      Width           =   9735
   End
   Begin VB.TextBox Text1 
      DataField       =   "OBSENFAC"
      DataSource      =   "Data1"
      Height          =   1095
      Left            =   6960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Data Data6 
      Caption         =   "Data6 FACTURAS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Facturas"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Imprimir factura"
      Height          =   855
      Left            =   7440
      Picture         =   "Form14.frx":0CF5
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar factura"
      Height          =   855
      Left            =   3960
      Picture         =   "Form14.frx":1137
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   2415
   End
   Begin VB.Data Data2 
      Caption         =   "DATA2 TABLA TOTALES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Caption         =   "Totales"
      Height          =   1695
      Left            =   240
      TabIndex        =   16
      Top             =   4560
      Width           =   6615
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   720
         TabIndex        =   28
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4680
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4680
         TabIndex        =   26
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4680
         TabIndex        =   25
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4680
         TabIndex        =   21
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   720
         TabIndex        =   20
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label24 
         Caption         =   "Suma total------------------------------------------------------->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label23 
         Caption         =   "De IVA------------------------------------------>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   33
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "De Recargo de Equivalencia----------->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Total-------------------------------------------------------------->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label18 
         Height          =   255
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1 AUXILIAR DE CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Clientes"
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   855
      Left            =   240
      Picture         =   "Form14.frx":1579
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del cliente"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   9735
      Begin VB.Data Data3 
         Caption         =   "Data3 AUXILIAR DE FACTURAS"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Aux Facturas"
         Top             =   0
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FAX"
         DataField       =   "fax"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   8160
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MOVIL"
         DataField       =   "movil"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   6600
         TabIndex        =   10
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TLF"
         DataField       =   "tlf"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5040
         TabIndex        =   9
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "PROVINCIA"
         DataField       =   "provincia"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LOCALIDAD"
         DataField       =   "localidad"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "CP"
         DataField       =   "cp"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DIRECCION"
         DataField       =   "direccion"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   9375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "APELLIDOS"
         DataField       =   "apellidos"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   9375
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOMBRE"
         DataField       =   "nombre"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   7935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DNI"
         DataField       =   "dni"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label17 
      Caption         =   "Observaciones de factura:"
      Height          =   255
      Left            =   6960
      TabIndex        =   24
      Top             =   4440
      Width           =   2535
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FECHA"
      DataField       =   "fecha"
      DataSource      =   "Data1"
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
      Left            =   7920
      TabIndex        =   15
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label15 
      BackColor       =   &H0000C000&
      Caption         =   " FECHA DE FACTURA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NFAC"
      DataField       =   "numfac"
      DataSource      =   "Data1"
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
      Left            =   4080
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackColor       =   &H0000C000&
      Caption         =   " NUMERO DE FACTURA: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   12
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim opc As Integer
    Dim consulta As String
    Dim consultat As String
    opc = MsgBox("¿Esta seguro de querer eliminar o abonar la factura actual?", vbOKCancel, "Eliminar factura")
    If opc = vbOK Then
        
        'Eliminacion de datos de tabla facturas
        consulta = "select * from facturas where nfac=""" + Label14.Caption + """"
        Data6.RecordSource = consulta
        Data6.Refresh
        If Data6.Recordset.RecordCount = 1 Then
            Data6.Recordset.MoveFirst
            Do While Not Data6.Recordset.EOF
                Data6.Recordset.Delete
                Data6.Recordset.MoveNext
            Loop
        Else
            MsgBox "¡NO ENCUENTRO LA FACTURA PARA PODER BORRARLA!", vbCritical, "Error en la busqueda de la factura"
            MsgBox "Consulte con el programador acerca de este error", vbInformation
        End If
        'Borramos de la tabla principal de totales
        consultat = "select * from totales where nfac=""" + Label14.Caption + """"
        Data2.RecordSource = consultat
        Data2.Refresh
        If Data2.Recordset.RecordCount = 1 Then
            Data2.Recordset.Delete
            Data2.Refresh
        Else
            MsgBox "NO ENCUENTRO EL REGISTRO A BORRAR EN TABLA TOTALES,CONSULTA AL PROGRAMADOR ACERCA DE ESTE ERROR", vbCritical, "Error en la eliminación"
        End If
        Unload Form14
    Else
        MsgBox "Operación de eliminacion cancelada por el usuario", vbInformation, "Factura no eliminada"
    End If
End Sub



Private Sub Command4_Click()
    DataEnvironment2.Connection1.Open
    DataEnvironment2.Command1
    DataReport5.Show vbModal
    DataEnvironment2.Connection1.Close
End Sub

Private Sub Form_Load()
    'Pasamos los totales a los labels que tenemos habilitados en el formulario
    
    
    Form14.Text2.Text = ivas
    Form14.Text3.Text = totalrecargo
    Form14.Text4.Text = totalivas
    Form14.Text5.Text = sumatotal
    Form14.Text6.Text = ptotal
    Form14.Text7.Text = recargo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Eliminacion de datos de las tablas auxiliares
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
            Data3.Recordset.Delete
            Data3.Recordset.MoveNext
        Loop
        Data1.Recordset.MoveFirst
        Data1.Recordset.Delete
        
End Sub


