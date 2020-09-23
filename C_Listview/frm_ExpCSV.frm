VERSION 5.00
Begin VB.Form frm_ExpCSV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar A CSV"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSeparate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2205
      TabIndex        =   6
      Text            =   ";"
      Top             =   600
      Width           =   450
   End
   Begin VB.TextBox txtEnclose 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2205
      MaxLength       =   1
      TabIndex        =   5
      Top             =   180
      Width           =   465
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4155
      TabIndex        =   2
      Top             =   1440
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   2715
      TabIndex        =   1
      Top             =   1440
      Width           =   1245
   End
   Begin VB.Frame frame 
      Height          =   120
      Left            =   75
      TabIndex        =   0
      Top             =   1110
      Width           =   5295
   End
   Begin VB.Label Label 
      Caption         =   "Puede ser  ,  o  ;"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2850
      TabIndex        =   7
      Top             =   630
      Width           =   2280
   End
   Begin VB.Label Label 
      Caption         =   "Campo separado por :"
      Height          =   270
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   585
      Width           =   1680
   End
   Begin VB.Label Label 
      Caption         =   "Campo encerrado por :"
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   3
      Top             =   165
      Width           =   1830
   End
End
Attribute VB_Name = "frm_ExpCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    c_lv.FieldSeparatorCSV = Me.txtSeparate
    c_lv.FieldEncloseCSV = Me.txtEnclose
    If c_lv.Export_ListviewToCSV(frm_Main.ListView, App.Path & "\TEXTO.csv") = True Then
     OpenFile App.Path & "\TEXTO.csv", Me
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
