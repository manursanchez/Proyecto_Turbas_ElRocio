VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form12 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de factura"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11400
   Icon            =   "Form12.frx":0000
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   2175
      Left            =   360
      TabIndex        =   24
      Top             =   480
      Width           =   11055
      Begin VB.Data Data5 
         Caption         =   "DATA5 FACTURAS"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   7320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Facturas"
         Top             =   2040
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "DNI"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "NOMBRE"
         Height          =   255
         Left            =   1800
         TabIndex        =   45
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "APELLIDOS"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "DIRECCION"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "dni"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "nombre"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   2760
         TabIndex        =   41
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "apellidos"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1080
         TabIndex        =   40
         Top             =   720
         Width           =   7575
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "direccion"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1080
         TabIndex        =   39
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label Label11 
         Caption         =   "CP"
         Height          =   255
         Left            =   6720
         TabIndex        =   38
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "cp"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   7200
         TabIndex        =   37
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "LOCALIDAD"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "localidad"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1080
         TabIndex        =   35
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label15 
         Caption         =   "PROVINCIA"
         Height          =   255
         Left            =   3840
         TabIndex        =   34
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "provincia"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   4800
         TabIndex        =   33
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label17 
         Caption         =   "OBSERVACIONES"
         Height          =   255
         Left            =   8760
         TabIndex        =   32
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "observaciones"
         DataSource      =   "Data1"
         Height          =   855
         Left            =   8040
         TabIndex        =   31
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label19 
         Caption         =   "TELEFONO"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "tlf"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   1200
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label21 
         Caption         =   "MOVIL"
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "movil"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   3480
         TabIndex        =   27
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "FAX"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   1800
         Width           =   375
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "fax"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   5520
         TabIndex        =   25
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Observaciones en factura"
      Height          =   1095
      Left            =   6360
      Picture         =   "Form12.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   " Pulsa aqui para introducir observaciones en la factura "
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Data Data6 
      Caption         =   "TOTALES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Totales"
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame4 
      Caption         =   "Funciones"
      Height          =   4095
      Left            =   9240
      TabIndex        =   13
      Top             =   2880
      Width           =   2055
      Begin VB.CommandButton nuevafactura 
         Caption         =   "Nueva factura"
         Height          =   855
         Left            =   360
         Picture         =   "Form12.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   " Crear una nueva factura "
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton salir 
         Caption         =   "Salir"
         Height          =   855
         Left            =   360
         Picture         =   "Form12.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salir al panel principal del programa"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CommandButton imprimirinforme 
         Caption         =   "Imprimir factura"
         Height          =   855
         Left            =   360
         Picture         =   "Form12.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Sacar la factura por impresora"
         Top             =   1680
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Total"
      Height          =   1695
      Left            =   6360
      TabIndex        =   10
      Top             =   2760
      Width           =   2655
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   55
         Text            =   "0"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   54
         Text            =   "0"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   53
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox euros 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Text            =   "0"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "Suma total"
         Height          =   255
         Left            =   1320
         TabIndex        =   58
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Total IVA"
         Height          =   255
         Left            =   1320
         TabIndex        =   57
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label25 
         Caption         =   "Total Recargo "
         Height          =   255
         Left            =   1320
         TabIndex        =   56
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label42 
         Caption         =   "Total Factura"
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Data Data4 
      Caption         =   "DATA 4 AUX DE CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Clientes"
      Top             =   2040
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Data Data3 
      Caption         =   "DATA3 AUXILIAR DE FACTURAS"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Aux Facturas"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7800
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Bindings        =   "Form12.frx":1412
      Height          =   1935
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   3413
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame2 
      Caption         =   "Materiales para introducir en factura"
      Height          =   2895
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   5895
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   52
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   51
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   50
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   47
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton añadir 
         Caption         =   "Añadir a factura"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   " Pulsar para añadir elementos a la factura "
         Top             =   2280
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hacer Calculos"
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         MaxLength       =   100
         TabIndex        =   17
         Top             =   360
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         MaxLength       =   9
         TabIndex        =   4
         Text            =   "1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label32 
         Caption         =   "%"
         Height          =   255
         Left            =   3480
         TabIndex        =   49
         Top             =   1800
         Width           =   135
      End
      Begin VB.Label Label30 
         Caption         =   "%"
         Height          =   255
         Left            =   3480
         TabIndex        =   48
         Top             =   1440
         Width           =   135
      End
      Begin VB.Label Label29 
         Caption         =   "Recargo de equivalencia"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label38 
         Caption         =   "IVA a aplicar"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label34 
         Caption         =   "Precio en € por unidad"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "Introduce la cantidad de unidades"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
   End
   Begin VB.Data Data1 
      Caption         =   "DATA1 CLIENTES"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   960
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.Label Label40 
      Caption         =   "Fecha"
      Height          =   255
      Left            =   7080
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "NUMERO DE FACTURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean
Dim guardar As Boolean
Dim imprimir As Boolean
Private Sub añadir_Click()
    
    'Declaracion de variables para hacer los calculos de factura
    Dim resultadoprecio As Currency 'alamacena text9
    Dim resultadore As Double 'almacena text7
    Dim resultadoiva As Double 'alamcena text8
    Dim totalfprecio As Currency 'text euros alamcena el precio por la cantidad
    Dim totalfre As Double 'text10 almacena el total del re
    Dim totalfiva As Double 'text11 almacena el totaldel iva
    Dim sumatotaldefactura As Currency 'text12 Suma de todos los elementos de factura
    Dim supersuma As Double 'Variable que suma todos los totales
    Dim suma As Double
    'Fin de declaracion de variables para calculos en factura
    
    Command2.Enabled = True
    If Form12.Text1.Text = "" Or Not IsNumeric(Form12.Text1.Text) Or Text3.Text = "" Or Text4.Text = "" Or Not IsNumeric(Text4.Text) Then
        MsgBox "Algunos de los campos cantidad, concepto o precio en euros estan vacíos o tienen valores NO NUMERICOS .Por favor compruebelo antes de continuar y haga el calculo del IVA antes de continuar añadiendo elementos a factura", vbExclamation, "Error de datos"
    Else
    
        cant = 0 'Inicializamos cant
        'Añadimos el cliente a la tabla auxiliar de clientes
        If bandera = False Then 'Controlamos que solo lo introduzcamos una sola vez
            Form12.Data4.Recordset.AddNew
            Form12.Data4.Recordset.Fields("numfac") = Form12.Label2.Caption
            Form12.Data4.Recordset.Fields("fecha") = Form12.Text2.Text
            Form12.Data4.Recordset.Fields("dni") = Form12.Label7.Caption
            Form12.Data4.Recordset.Fields("nombre") = Form12.Label8.Caption
            Form12.Data4.Recordset.Fields("apellidos") = Form12.Label9.Caption
            Form12.Data4.Recordset.Fields("direccion") = Form12.Label10.Caption
            Form12.Data4.Recordset.Fields("cp") = Form12.Label12.Caption
            Form12.Data4.Recordset.Fields("localidad") = Form12.Label14.Caption
            Form12.Data4.Recordset.Fields("provincia") = Form12.Label16.Caption
            Form12.Data4.Recordset.Fields("tlf") = Form12.Label20.Caption
            Form12.Data4.Recordset.Fields("fax") = Form12.Label24.Caption
            Form12.Data4.Recordset.Fields("movil") = Form12.Label22.Caption
            Form12.Data4.Recordset.Fields("observaciones") = Form12.Label18.Caption
            
            Form12.Data4.Recordset.Update
            Form12.Data4.Refresh
            bandera = True
        End If
        
        'Vamos a añadir elementos a la tabla auxiliar de facturas
    
        Form12.Data3.Recordset.AddNew
        Form12.Data3.Recordset.Fields("facnum") = Form12.Label2.Caption
        Form12.Data3.Recordset.Fields("concepto") = Form12.Text3.Text
        Form12.Data3.Recordset.Fields("cantidad") = Form12.Text1.Text
        Form12.Data3.Recordset.Fields("punidad") = Val(Form12.Text4.Text)
        Form12.Data3.Recordset.Fields("ptotal") = Form12.Text9.Text
        Form12.Data3.Recordset.Fields("rq") = Val(Form12.Text5.Text)
        Form12.Data3.Recordset.Fields("totalrq") = Form12.Text7.Text
        Form12.Data3.Recordset.Fields("iva") = Val(Form12.Text6.Text)
        Form12.Data3.Recordset.Fields("totaliva") = Form12.Text8.Text
        
        'Vamos a hacer las operaciones matematicas para los totales
        'Las mostraremos primero por pantalla en sus correspondientes
        'textboxs y despues los meteremos en la tabla
        
        'Calculos de totales
        'Pasamos el contenido de los text a las variables
        resultadoprecio = Text9.Text
        resultadore = Text7.Text
        resultadoiva = Text8.Text
        suma = resultadoprecio + resultadore + resultadoiva
        Form12.Data3.Recordset.Fields("sumatotalregistro") = Round(suma, 2) 'Añadimos este a la base de datos
        
        totalfprecio = euros.Text
        totalfre = Text10.Text
        totalfiva = Text11.Text
        sumatotaldefactura = Text12.Text
        'Hacemos los calculos
        totalfprecio = totalfprecio + resultadoprecio
        totalfre = totalfre + resultadore
        totalfiva = totalfiva + resultadoiva
        supersuma = totalfprecio + totalfre + totalfiva
        
        sumatotaldefactura = supersuma
        'Y pasamos el valor de las variables a los text
        euros.Text = Round(totalfprecio, 2)
        Text10.Text = totalfre
        Text11.Text = totalfiva
        Text12.Text = Round(sumatotaldefactura, 2)
        
        'Introducimos las cantidades y los precios en la tabla auxiliar
        Form12.Data3.Recordset.Update
        Form12.Data3.Refresh
        Form12.Text1.Text = 1
    End If
    Text3.Text = ""
    Text4.Text = ""
    Text9.Text = ""
    Text8.Text = ""
    Text7.Text = ""
    añadir.Enabled = False
End Sub

Private Sub Command1_Click()
    'Variables para hacer los calculos
    Dim preciounidad As Currency 'alamcena text4
    Dim cant As Integer 'alamacena text1
    Dim recargo As Double 'almacena text5
    Dim iva As Double 'alamcena text6
    Dim resultadoprecio As Currency 'alamacena text9
    Dim resultadore As Double 'almacena text7
    Dim resultadoiva As Double 'alamcena text8
    Dim comodin As Double
    If Text4.Text = "" Or Not IsNumeric(Text4.Text) Or Text1.Text = "" Or Not IsNumeric(Text1.Text) Or Not IsNumeric(Text5.Text) Or Text5.Text = "" Or Not IsNumeric(Text6.Text) Or Text6.Text = "" Then
        MsgBox "Imposible hacer calculos ya que hay algún campo vacío o erroneo. Compruebe los campos precio en euros y cantidad antes de continuar", vbExclamation, "Error de datos"
    Else
        añadir.Enabled = True
        Text5.Enabled = False
        Text6.Enabled = False
        If Text4.Text = 0 Then
            Text9.Text = 0
            Text7.Text = 0
            Text8.Text = 0
        Else
            preciounidad = Val(Text4.Text)
            cant = Text1.Text
            recargo = Val(Text5.Text)
            iva = Val(Text6.Text)
            resultadoprecio = preciounidad * cant
            resultadore = resultadoprecio * recargo / 100
            resultadoiva = resultadoprecio * iva / 100
            Text9.Text = Round(resultadoprecio, 2)
            Text7.Text = resultadore
            Text8.Text = resultadoiva
        End If
    End If
End Sub

Private Sub Command2_Click()
    Load Form6
    Form6.Show vbModal
End Sub


Private Sub imprimirinforme_Click()
    If Not Form12.Data3.Recordset.BOF Then
        Form12.añadir.Enabled = False 'Desactivamos el boton añadir a factura
        'Form12.eliminar.Enabled = False 'Desactivamos el boton elimianr componente de factura
        Form12.Command2.Enabled = False 'Desactivamos el boton de observaciones en factura
        guardarfactura
        guardartotales
        guardar = True
        imprimir = True
        DataEnvironment2.Connection1.Open
        DataEnvironment2.Command1
        DataReport5.Show vbModal
        DataEnvironment2.Connection1.Close
    Else
        MsgBox "No puedo imprimir nada por que no se han introducido datos en la factura", vbInformation
    End If
End Sub

'Private Sub eliminar_Click()
 '   If Form12.Data3.Recordset.BOF Then
  '      MsgBox "Imposible hacer cambios ya que no hay datos en la factura", vbInformation
   ' Else
   '     Load Form10
   '     Form10.Show vbModal
   ' End If
    
'End Sub

Private Sub salir_Click()
    Dim opc As Integer
    If Not Form12.Data3.Recordset.BOF Then
        If guardar = False Then
            opc = MsgBox("¿Desea guardar la factura creada?", vbOKCancel, "Pulse ACEPTAR para guardar la factura,CANCELAR para no guardarla")
            If opc = vbOK Then
                guardar = True
                guardarfactura
                guardartotales
            End If
        End If
    End If
    Unload Me
End Sub

Private Sub nuevafactura_Click()
    Dim opc As Integer
    If Not Form12.Data3.Recordset.BOF Then
        If guardar = False Then
            opc = MsgBox("¿Desea guardar la factura?", vbOKCancel, "Elija ACEPTAR para guardar la factura o CANCEL para NO guardarla")
            If opc = vbOK Then
                guardarfactura
                guardartotales
            End If
        End If
    End If
    Unload Me
    Load Form9
    Form9.Show vbModal
End Sub


Private Sub Form_Load()
    Dim consulta As String
    Command2.Enabled = False
    guardar = False
    imprimir = False
    bandera = False
    tpesetas = 0
    teuros = 0
    tptasiva = 0
    teurosiva = 0
    
    Form12.Label2.Caption = numfactura
    'Las siguientes instrucciones son para buscar el cliente
    'y que lo muestre en el formulario de diseño de factura
    consulta = "select * from clientes where dni=""" + identificacion + """"
    Form12.Data1.RecordSource = consulta
    Form12.Data1.Refresh
    Form12.Text2.Text = Date
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Borramos registros de Auxiliar de clientes
    
    borrarauxclientes
    
    'Borramos registros de auxiliar de facturas
    borrarauxfacturas
    
End Sub
Private Sub guardarfactura() 'Guarda de facturas auxiliares a Facturas
     If imprimir = False Then
        
        If Form12.Data4.Recordset.BOF = True Then
            MsgBox "No hay factura para guardar ni para imprimir.Añada elementos a factura y guarde o imprima despues", vbInformation, "Error de datos"
        Else
            Form12.Data3.Recordset.MoveFirst
            Do While Not Form12.Data3.Recordset.EOF
                Form12.Data5.Recordset.AddNew
                Form12.Data5.Recordset.Fields("nfac") = Form12.Label2.Caption
                Form12.Data5.Recordset.Fields("fecha") = Form12.Text2.Text
                Form12.Data5.Recordset.Fields("dni") = Form12.Label7.Caption
                Form12.Data5.Recordset.Fields("nombre") = Form12.Label8.Caption
                Form12.Data5.Recordset.Fields("apellidos") = Form12.Label9.Caption
                Form12.Data5.Recordset.Fields("direccion") = Form12.Label10.Caption
                Form12.Data5.Recordset.Fields("cp") = Form12.Label12.Caption
                Form12.Data5.Recordset.Fields("localidad") = Form12.Label14.Caption
                Form12.Data5.Recordset.Fields("provincia") = Form12.Label16.Caption
                Form12.Data5.Recordset.Fields("tlf") = Form12.Label20.Caption
                Form12.Data5.Recordset.Fields("movil") = Form12.Label22.Caption
                Form12.Data5.Recordset.Fields("fax") = Form12.Label24.Caption
                Form12.Data5.Recordset.Fields("observaciones") = Form12.Label12.Caption
                Form12.Data5.Recordset.Fields("concepto") = Form12.Data3.Recordset.Fields("concepto")
                Form12.Data5.Recordset.Fields("cantidad") = Form12.Data3.Recordset.Fields("cantidad")
                Form12.Data5.Recordset.Fields("punidad") = Form12.Data3.Recordset.Fields("punidad")
                Form12.Data5.Recordset.Fields("ptotal") = Form12.Data3.Recordset.Fields("ptotal")
                Form12.Data5.Recordset.Fields("rq") = Form12.Data3.Recordset.Fields("rq")
                Form12.Data5.Recordset.Fields("totalrq") = Form12.Data3.Recordset.Fields("totalrq")
                Form12.Data5.Recordset.Fields("iva") = Form12.Data3.Recordset.Fields("iva")
                Form12.Data5.Recordset.Fields("totaliva") = Form12.Data3.Recordset.Fields("totaliva")
                Form12.Data5.Recordset.Fields("sumatotalregistro") = Form12.Data3.Recordset.Fields("sumatotalregistro")
                Form12.Data5.Recordset.Fields("OBSENFAC") = Form12.Data4.Recordset.Fields("OBSENFAC")
                Form12.Data5.UpdateRecord
                Form12.Data3.Recordset.MoveNext
            Loop
            MsgBox "Factura guardada sin problemas", vbInformation
            Form12.Data5.Refresh
        End If
    End If
End Sub
Private Sub borrarauxclientes()
    If Not Form12.Data4.Recordset.BOF Then
        Form12.Data4.Recordset.MoveFirst
        Form12.Data4.Recordset.Delete
        Form12.Data4.UpdateRecord
        Form12.Data4.Refresh
    End If
End Sub
Private Sub borrarauxfacturas()
    If Not Form12.Data3.Recordset.BOF Then
        Form12.Data3.Recordset.MoveFirst
        Do While Not Form12.Data3.Recordset.EOF
            Form12.Data3.Recordset.Delete
            Form12.Data3.Recordset.MoveNext
        Loop
        Form12.Data3.Refresh
    End If
End Sub
Private Sub guardartotales()
    If imprimir = False Then
        Form12.Data6.Recordset.AddNew
        Form12.Data6.Recordset.Fields("nfac") = Form12.Label2.Caption
        Form12.Data6.Recordset.Fields("ptotal") = Form12.euros.Text
        Form12.Data6.Recordset.Fields("rq") = Val(Form12.Text5.Text)
        Form12.Data6.Recordset.Fields("totalrq") = Form12.Text10.Text
        Form12.Data6.Recordset.Fields("iva") = Val(Form12.Text6.Text)
        Form12.Data6.Recordset.Fields("totaliva") = Form12.Text11.Text
        Form12.Data6.Recordset.Fields("sumatotal") = Form12.Text12.Text
        Form12.Data6.Recordset.Update
        Form12.Data6.Refresh
    End If
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Not IsDate(Form12.Text2.Text) Then
        MsgBox "No reconozco este formato de fecha. Utilice el formato dia/mes/año. Por ejemplo: 10/05/2002 sería 10 de mayo del año 2002", vbExclamation, "Error de tipos de datos"
        Cancel = True
    End If
End Sub

