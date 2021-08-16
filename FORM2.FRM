VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccionar opcion"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "Ver clientes"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nuevo cliente"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Form2.Option1.Value Then
        Unload Form2
        Load Form5
        Form5.Show vbModal
    Else
        If Form2.Option2.Value Then
            Unload Form2
            Load Form3
            Form3.Show vbModal
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
