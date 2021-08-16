VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo cliente"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10095
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4680
      MaxLength       =   9
      TabIndex        =   23
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Text13 
      Height          =   855
      Left            =   1800
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   6120
      MaxLength       =   10
      TabIndex        =   20
      Text            =   "0"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   19
      Text            =   "0"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   18
      Text            =   "0"
      Top             =   1920
      Width           =   1455
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   7320
      MaxLength       =   50
      TabIndex        =   17
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      DataField       =   "localidad"
      Height          =   285
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   16
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   15
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      MaxLength       =   100
      TabIndex        =   14
      Top             =   1200
      Width           =   8055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      MaxLength       =   100
      TabIndex        =   13
      Top             =   840
      Width           =   8055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      MaxLength       =   50
      TabIndex        =   12
      Top             =   480
      Width           =   6495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CERRAR"
      Height          =   855
      Left            =   8400
      Picture         =   "Form5.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AÑADIR "
      Height          =   855
      Left            =   6600
      Picture         =   "Form5.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2280
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Clientes"
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label2 
      Caption         =   "DNI/CIF/NIF"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "NOMBRE/RAZON SOCIAL"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label14 
      Caption         =   "APELLIDOS"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label13 
      Caption         =   "DIRECCION"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label12 
      Caption         =   "CP"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "LOCALIDAD"
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "PROVINCIA"
      Height          =   255
      Left            =   6240
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "OBSERVACIONES"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "TELEFONO"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "MOVIL"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "FAX"
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
            Dim consulta As String
            If Not IsNumeric(Form5.Text5.Text) Or Not IsNumeric(Form5.Text10.Text) Or Not IsNumeric(Form5.Text11.Text) Or Not IsNumeric(Form5.Text12.Text) Or Form5.Text1.Text = "" Then
                MsgBox "Algunos de los datos del cliente no estan correctos o hay campos que estan vacios,reviselos antes de continuar", vbInformation, "Error de datos introducidos"
            Else
                consulta = "select * from clientes where dni=""" + Form5.Text1.Text + """"
                Form5.Data1.RecordSource = consulta
                Form5.Data1.Refresh
                If Form5.Data1.Recordset.RecordCount < 1 Then
                    Form5.Data1.Recordset.AddNew
                    Form5.Data1.Recordset.Fields("dni") = UCase(Form5.Text1.Text)
                    Form5.Data1.Recordset.Fields("nombre") = UCase(Form5.Text2.Text)
                    Form5.Data1.Recordset.Fields("apellidos") = UCase(Form5.Text3.Text)
                    Form5.Data1.Recordset.Fields("direccion") = UCase(Form5.Text4.Text)
                    Form5.Data1.Recordset.Fields("cp") = Form5.Text5.Text
                    Form5.Data1.Recordset.Fields("localidad") = UCase(Form5.Text6.Text)
                    Form5.Data1.Recordset.Fields("provincia") = UCase(Form5.Text7.Text)
                    Form5.Data1.Recordset.Fields("tlf") = Val(Form5.Text10.Text)
                    Form5.Data1.Recordset.Fields("movil") = Val(Form5.Text11.Text)
                    Form5.Data1.Recordset.Fields("fax") = Val(Form5.Text12.Text)
                    Form5.Data1.Recordset.Fields("observaciones") = UCase(Form5.Text13.Text)
                    Form5.Data1.Recordset.Update
                    Form5.Data1.Refresh
                    Form5.Text1.Text = ""
                    Form5.Text2.Text = ""
                    Form5.Text3.Text = ""
                    Form5.Text4.Text = ""
                    Form5.Text5.Text = 0
                    Form5.Text6.Text = ""
                    Form5.Text7.Text = ""
                    
                    
                    Form5.Text10.Text = 0
                    Form5.Text11.Text = 0
                    Form5.Text12.Text = 0
                    Form5.Text13.Text = ""
                Else
                    MsgBox "El cliente que se va a dar de alta ya existe en la base de datos.Imposible introducirlo dos o mas veces", vbCritical, "Error de datos"
                End If
            End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub



