VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Elije opción"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.OptionButton Option2 
         Caption         =   "Recuperar una factura ya creada"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Crear una nueva factura"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Form8.Option1 Then
        Unload Form8
        Load Form9
        Form9.Show vbModal
    End If
    If Form8.Option2 Then
        Unload Form8
        Load Form13
        Form13.Show vbModal
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


