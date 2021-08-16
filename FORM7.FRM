VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizacion de materiales"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Usa los cursores para desplazarte a través de los registros"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Taller\Taller Hmnos Maldonado Almagro.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Materiales"
      Top             =   2400
      Width           =   6135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   5160
      TabIndex        =   15
      ToolTipText     =   "Cierra esta ventana"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar"
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      ToolTipText     =   "Permite la modificacion y edicion de los campos"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   1800
      TabIndex        =   13
      ToolTipText     =   "Elimina el elemento que se ve en pantalla"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir listado"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Imprime un listado con todos los materiales de la Base de Datos"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      DataField       =   "peurosiva"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5160
      TabIndex        =   11
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      DataField       =   "pptasiva"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      DataField       =   "peuros"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "pptas"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "concepto"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      TabIndex        =   3
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      DataField       =   "codigo"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Precio en Ptas + IVA"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Precio en Euros + IVA"
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Precio en pesetas"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Precio en Euros"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Concepto o nombre:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Código material"
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bandera As Boolean
Private Sub Command1_Click()
    DataReport4.PrintReport
End Sub

Private Sub Command2_Click()
    Dim opc As Integer
    opc = MsgBox("¿Esta seguro de querer eliminar el elemento actual?", vbCritical + vbOKCancel)
    If opc = vbOK Then
        If Form7.Data1.Recordset.BOF = True Then
            MsgBox "¡¡¡No hay registros para borrar!!!", vbExclamation, "Error de datos"
        Else
            Form7.Data1.Recordset.Delete
            Form7.Data1.Refresh
            MsgBox "Material eliminado", vbInformation
        End If
    End If
End Sub

Private Sub Command3_Click()
    If bandera = False Then
        Form7.Text2.Enabled = True
        Form7.Text3.Enabled = True
        Form7.Command3.Caption = "Modificado"
        bandera = True
        Form7.Data1.Enabled = False
        Form7.Command4.Enabled = False
        Form7.Command2.Enabled = False
        Form7.Command1.Enabled = False
    Else
        If bandera = True Then
            Form7.Text2.Enabled = False
            Form7.Text3.Enabled = False
            Form7.Command3.Caption = "Modificar"
            bandera = False
            Form7.Data1.Enabled = True
            Form7.Command4.Enabled = True
            Form7.Command1.Enabled = True
            Form7.Command2.Enabled = True
        End If
    End If
        
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    bandera = False
    iva = 16
    valor1euro = 166.386
End Sub
Private Sub Text2_Validate(Cancel As Boolean)
    If Form7.Text2.Text = "" Then
        MsgBox "Debe introducir un valor en este campo", vbExclamation
        Cancel = True
    End If
End Sub
Private Sub Text3_Validate(Cancel As Boolean)
    If Form7.Text3.Text = "" Or Not IsNumeric(Form7.Text3.Text) Then
        MsgBox "Debe introducir un valor en este campo o el valor introducido no es numérico", vbExclamation
        Cancel = True
    Else
        ptas = Form7.Text3.Text
        € = ptas / valor1euro
        €iva = € * iva / 100
        €iva = €iva + €
        ptasiva = ptas * iva / 100
        ptasiva = ptas + ptasiva
        Form7.Text4.Text = Format(€, "###.##")
        Form7.Text5.Text = ptasiva
        Form7.Text6.Text = Format(€iva, "###.##")
    End If
End Sub


