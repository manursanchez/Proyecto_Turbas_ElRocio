VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recibos a clientes"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "Recibos.frx":0000
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Nuevo recibo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Calcula"
      Height          =   495
      Left            =   2280
      TabIndex        =   18
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Añade"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   "1"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   6
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   4
      Top             =   480
      Width           =   6015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   1335
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Recibos.frx":030A
      Height          =   1335
      Left            =   3480
      OleObjectBlob   =   "Recibos.frx":031E
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "recibo"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Turbas\Turbas y abonos el rocio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientesrecibos"
      Top             =   120
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1080
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Suma Total"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Precio"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Concepto"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "DNI/CIF"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim control As Boolean
Dim imprimido As Boolean
Private Sub Command1_Click()
    Dim sumatotal As Currency
    Dim total As Currency
    Command5.Enabled = True
    If Label9.Caption = 0 Then
        Label8.Caption = 0
    Else
        sumatotal = Val(Label8.Caption)
        total = Label9.Caption
        sumatotal = Round(sumatotal + total, 2)
        Label8.Caption = sumatotal
        If control = False Then
            Data1.Recordset.AddNew
            Data1.Recordset.Fields("dni") = Text1.Text
            Data1.Recordset.Fields("nombre") = Text2.Text
            Data1.Recordset.Update
            Data1.Refresh
            control = True
            Text1.Enabled = False
            Text2.Enabled = False
        End If
        Data2.Recordset.AddNew
        Data2.Recordset.Fields("nombre") = Text2.Text
        Data2.Recordset.Fields("concepto") = Text3.Text
        Data2.Recordset.Fields("cantidad") = Text5.Text
        Data2.Recordset.Fields("precio") = Val(Text4.Text)
        Data2.Recordset.Fields("total") = Label9.Caption
        Data2.Recordset.Update
        Data2.Refresh
        Command1.Enabled = False
        Text3.Text = ""
        Text4.Text = 0
        Text5.Text = 1
        Label9.Caption = 0
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If Data1.Recordset.BOF Then
        MsgBox "No se puede imprimir nada por que no se ha introducido ningun elemento en el recibo", vbInformation, "Error de datos"
    Else
        DataEnvironment2.Connection1.Open
        DataEnvironment2.Command3
        DataReport2.Show vbModal
        DataEnvironment2.Connection1.Close
        imprimido = True
    End If
End Sub

Private Sub Command4_Click()
    Dim precio As Currency
    Dim total As Currency
    Dim cant As Integer
    Dim sumatotal As Currency
    If Text3.Text = "" Or Text4.Text = "" Or Not IsNumeric(Text4.Text) Or Not IsNumeric(Text5.Text) Or Text5.Text = "" Then
        MsgBox "Imposible hacer calculos si no hay datos en los campos CONCEPTO, PRECIO ,CANTIDAD, o los datos introducidos son incorrectos", vbExclamation, "Error de datos"
    Else
        If Text4.Text = 0 Then
            Label8.Caption = 0
            Label9.Caption = 0
            Command1.Enabled = True
        Else
            If Not IsNumeric(Text4.Text) Then
                MsgBox "Imposible hacer calculos por que hay valores no válidos, compruebe el precio"
            Else
                precio = Val(Text4.Text)
                cant = Text5.Text
                total = precio * cant
                Label9.Caption = Round(total, 2)
                Command1.Enabled = True
            End If
        End If
    End If
End Sub
Private Sub Command5_Click()
    Dim opc As Integer
    If imprimido = False Then
        opc = MsgBox("No ha imprimido este recibo ¿desea imprimirlo?", vbYesNo, "Información de sistema")
        If opc = vbYes Then
            DataEnvironment2.Connection1.Open
            DataEnvironment2.Command3
            DataReport2.Show vbModal
            DataEnvironment2.Connection1.Close
        Else
            MsgBox "Se borrara el recibo actual y deberá hacerlo de nuevo para poder imprimirlo", vbInformation, "Informacion de sistema"
        End If
    End If
     Text1.Text = 0
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = 1
     Text1.Enabled = True
     Text2.Enabled = True
     Label9.Caption = ""
     Label8.Caption = ""
    If Not Form15.Data1.Recordset.BOF Then
         Form15.Data1.Recordset.MoveFirst
         Form15.Data1.Recordset.Delete
         Data1.UpdateRecord
         Form15.Data1.Refresh
     End If
     If Not Form15.Data2.Recordset.BOF Then
         Form15.Data2.Recordset.MoveFirst
         Do While Not Form15.Data2.Recordset.EOF
             Form15.Data2.Recordset.Delete
             Form15.Data2.Recordset.MoveNext
         Loop
         Form15.Data2.Refresh
     End If
     Command5.Enabled = False
     control = False
     imprimido = False
End Sub

Private Sub Form_Load()
    control = False
    imprimido = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Borramos lo que hay en las tablas
    If Not Data1.Recordset.BOF Then
        Data1.Recordset.MoveFirst
        Data1.Recordset.Delete
        Data1.UpdateRecord
        Data1.Refresh
    End If
    If Not Data2.Recordset.BOF Then
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            Data2.Recordset.Delete
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
    End If
End Sub
