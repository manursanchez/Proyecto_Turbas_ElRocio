VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Turbas y abonos ""El Rocio"" SCA"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6780
   Icon            =   "Pagina principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   0
      Picture         =   "Pagina principal.frx":030A
      ScaleHeight     =   3195
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command1 
         Caption         =   "Clientes"
         Height          =   855
         Left            =   360
         Picture         =   "Pagina principal.frx":511D
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Dar de alta un nuevo cliente o ver los clientes en la base de datos "
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Facturas"
         Height          =   855
         Left            =   360
         Picture         =   "Pagina principal.frx":555F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " Crear o ver facturas "
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Recibos a clientes"
         Height          =   855
         Left            =   4200
         Picture         =   "Pagina principal.frx":59A1
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Hacer recibos "
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   855
         Left            =   4200
         Picture         =   "Pagina principal.frx":5DE3
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir del programa "
         Top             =   2160
         Width           =   2055
      End
   End
   Begin VB.Menu opciones 
      Caption         =   "Utilidades y Opciones"
      Begin VB.Menu acerca 
         Caption         =   "Acerca de este programa"
      End
      Begin VB.Menu repara 
         Caption         =   "Reparar base de datos"
      End
      Begin VB.Menu borrarfacturas 
         Caption         =   "Facturación nueva"
      End
      Begin VB.Menu guion 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir del programa"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub acerca_Click()
    Load Form18
    Form18.Show vbModal
End Sub

Private Sub borrarfacturas_Click()
    Load Form16
    Form16.Show vbModal
End Sub

Private Sub Command1_Click()
    Load Form2
    Form2.Show vbModal
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Load Form8
    Form8.Show vbModal
End Sub

Private Sub Command4_Click()
    Load Form15
    Form15.Show vbModal
End Sub

Private Sub repara_Click()
    Load Form4
    Form4.Show vbModal
End Sub

Private Sub salir_Click()
    Unload Me
End Sub
