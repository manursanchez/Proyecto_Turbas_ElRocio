VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Facturacion Hmnos Maldonado Almagro"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10320
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6015
      Left            =   240
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   5955
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   240
      Width           =   9735
      Begin VB.CommandButton Command4 
         Caption         =   "Presupuestos"
         Height          =   855
         Left            =   7320
         Picture         =   "Form1.frx":13A3A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Facturas"
         Height          =   855
         Left            =   4920
         Picture         =   "Form1.frx":13E7C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Materiales"
         Height          =   855
         Left            =   2520
         Picture         =   "Form1.frx":142BE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clientes"
         Height          =   855
         Left            =   120
         Picture         =   "Form1.frx":14700
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4800
         Width           =   2055
      End
   End
   Begin VB.Menu opciones 
      Caption         =   "Opciones de menu"
      Begin VB.Menu cambiar 
         Caption         =   "Cambiar imagen"
      End
      Begin VB.Menu datos1 
         Caption         =   "Datos del programa"
      End
      Begin VB.Menu datos2 
         Caption         =   "Datos del programador"
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
Private Sub cambiar_Click()
'    Picture1.Picture = LoadPicture("c:\rallyfocussainz.jpg")
End Sub

Private Sub Command1_Click()
    Load Form2
    Form2.Show vbModal
End Sub

Private Sub Command2_Click()
    Load Form4
    Form4.Show vbModal
End Sub

Private Sub Command3_Click()
    Load Form8
    Form8.Show vbModal
End Sub

Private Sub salir_Click()
    Unload Me
End Sub
