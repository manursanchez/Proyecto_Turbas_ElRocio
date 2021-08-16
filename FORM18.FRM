VERSION 5.00
Begin VB.Form Form18 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de este programa"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   Icon            =   "Form18.frx":0000
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "ACEPTAR"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Realizado por:"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Tlf: 607 544 555     Email: mrodriguezs@eresmas.com, chorbo@airtel.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Técnico especialista en Informática de Gestión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   1800
      Width           =   6015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Manuel Rodríguez Sánchez"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PROGRAMA DE FACTURAS PARA ""TURBAS Y ABONOS BIOLOGICOS EL ROCIO SCA"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub
