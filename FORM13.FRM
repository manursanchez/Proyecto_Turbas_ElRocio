VERSION 5.00
Begin VB.Form Form13 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recuperar factura"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "Form13.frx":0000
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data4 
      Caption         =   "Data4 TOTALES"
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
      RecordSource    =   "Totales"
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Data Data3 
      Caption         =   "DATA 3 AUXILIAR DE FACTURAS"
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
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Data Data2 
      Caption         =   "DATA 2 AUXILIAR DE CLIENTES"
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
      RecordSource    =   "Aux Clientes"
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Data Data1 
      Caption         =   "DATA 1 FACTURAS"
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
      RecordSource    =   "Facturas"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCELAR"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ACEPTAR"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Inoduce el número de factura:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Buscamos la factura para ver si existe o no
    Dim consulta As String
    Dim consultat As String
    consulta = "select * from facturas where nfac=""" + Form13.Text1.Text + """"
    Form13.Data1.RecordSource = consulta
    Form13.Data1.Refresh
    If Form13.Data1.Recordset.RecordCount = 1 Then
        'Vamos a guardar el cliente en auxiliar de clientes
        
        Form13.Data2.Recordset.AddNew
        Form13.Data2.Recordset.Fields("numfac") = Form13.Data1.Recordset("nfac")
        Form13.Data2.Recordset.Fields("fecha") = Form13.Data1.Recordset("fecha")
        Form13.Data2.Recordset.Fields("dni") = Form13.Data1.Recordset("dni")
        Form13.Data2.Recordset.Fields("nombre") = Form13.Data1.Recordset("nombre")
        Form13.Data2.Recordset.Fields("apellidos") = Form13.Data1.Recordset("apellidos")
        Form13.Data2.Recordset.Fields("direccion") = Form13.Data1.Recordset("direccion")
        Form13.Data2.Recordset.Fields("cp") = Form13.Data1.Recordset("cp")
        Form13.Data2.Recordset.Fields("localidad") = Form13.Data1.Recordset("localidad")
        Form13.Data2.Recordset.Fields("provincia") = Form13.Data1.Recordset("provincia")
        Form13.Data2.Recordset.Fields("tlf") = Form13.Data1.Recordset("tlf")
        Form13.Data2.Recordset.Fields("movil") = Form13.Data1.Recordset("movil")
        Form13.Data2.Recordset.Fields("fax") = Form13.Data1.Recordset("fax")
        Form13.Data2.Recordset.Fields("observaciones") = Form13.Data1.Recordset("observaciones")
        Form13.Data2.Recordset.Fields("OBSENFAC") = Form13.Data1.Recordset.Fields("OBSENFAC")
        Form13.Data2.UpdateRecord
        Form13.Data2.Refresh
        
        'Vamos a añadir todos los elementos de la factura a Auxiliar de Facturas
        
        Form13.Data1.Recordset.MoveFirst
        Form13.Data1.Refresh
        Do While Not Form13.Data1.Recordset.EOF
           If Form13.Data1.Recordset.Fields("nfac") = Form13.Text1.Text Then
                Form13.Data3.Recordset.AddNew
                Form13.Data3.Recordset.Fields("facnum") = Form13.Data1.Recordset.Fields("nfac")
                Form13.Data3.Recordset.Fields("concepto") = Form13.Data1.Recordset.Fields("concepto")
                Form13.Data3.Recordset.Fields("cantidad") = Form13.Data1.Recordset.Fields("cantidad")
                Form13.Data3.Recordset.Fields("punidad") = Form13.Data1.Recordset.Fields("punidad")
                Form13.Data3.Recordset.Fields("ptotal") = Form13.Data1.Recordset.Fields("ptotal")
                Form13.Data3.Recordset.Fields("rq") = Form13.Data1.Recordset.Fields("rq")
                Form13.Data3.Recordset.Fields("totalrq") = Form13.Data1.Recordset.Fields("totalrq")
                Form13.Data3.Recordset.Fields("iva") = Form13.Data1.Recordset.Fields("iva")
                Form13.Data3.Recordset.Fields("sumatotalregistro") = Form13.Data1.Recordset.Fields("sumatotalregistro")
                Form13.Data3.Recordset.Fields("totaliva") = Form13.Data1.Recordset.Fields("totaliva")
                Form13.Data3.UpdateRecord
            End If
            Form13.Data1.Recordset.MoveNext
         Loop
         consultat = "select * from totales where nfac=""" + Form13.Text1.Text + """"
         Form13.Data4.RecordSource = consultat
         Form13.Data4.Refresh
         If Data4.Recordset.RecordCount = 1 Then
            recargo = Form13.Data4.Recordset.Fields("rq")
            totalrecargo = Form13.Data4.Recordset.Fields("totalrq")
            totalivas = Form13.Data4.Recordset.Fields("totaliva")
            ivas = Form13.Data4.Recordset.Fields("iva")
            sumatotal = Form13.Data4.Recordset.Fields("sumatotal")
            ptotal = Form13.Data4.Recordset.Fields("ptotal")
            Unload Form13
            Load Form14
            Form14.Show vbModal
         Else
            MsgBox "No encuentro los totales de la factura", vbCritical, "Error en la búsqueda"
         End If
    Else
        MsgBox "El número de factura introducido no existe.Introduzca un nuevo número de factura", vbInformation, "Factura inexistente"
    End If
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

