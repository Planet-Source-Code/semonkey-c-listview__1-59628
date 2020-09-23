VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_Main 
   Caption         =   "Listview"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pct 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1455
      ScaleWidth      =   11730
      TabIndex        =   2
      Top             =   6000
      Width           =   11730
      Begin VB.ComboBox Combo 
         Height          =   315
         ItemData        =   "frm_Main.frx":0000
         Left            =   9885
         List            =   "frm_Main.frx":000A
         TabIndex        =   14
         Text            =   "Green"
         Top             =   810
         Width           =   1785
      End
      Begin VB.TextBox txtMaximo 
         Height          =   300
         Left            =   10230
         TabIndex        =   11
         Text            =   "35"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdBar 
         Caption         =   "Disable ProgressBar"
         Height          =   360
         Left            =   9135
         TabIndex        =   10
         Top             =   45
         Width           =   2340
      End
      Begin VB.OptionButton opTbl 
         Caption         =   "Table Pedidos"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   8
         Top             =   75
         Width           =   3120
      End
      Begin VB.OptionButton opTbl 
         Caption         =   "Table detalle"
         Height          =   255
         Index           =   1
         Left            =   3300
         TabIndex        =   7
         Top             =   60
         Width           =   3120
      End
      Begin VB.OptionButton opTbl 
         Caption         =   "Table Productos"
         Height          =   255
         Index           =   2
         Left            =   6525
         TabIndex        =   6
         Top             =   60
         Width           =   3120
      End
      Begin VB.CommandButton cmdExportarTXT 
         Caption         =   "Export to Text"
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1485
      End
      Begin VB.CommandButton cmdExportarCSV 
         Caption         =   "Export to CSV"
         Height          =   360
         Left            =   1740
         TabIndex        =   4
         Top             =   540
         Width           =   1515
      End
      Begin VB.CommandButton cmdExportarHTML 
         Caption         =   "Export to Html"
         Height          =   360
         Left            =   3375
         TabIndex        =   3
         Top             =   555
         Width           =   1770
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   195
         Left            =   45
         TabIndex        =   9
         Top             =   1185
         Width           =   11640
         _ExtentX        =   20532
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Min             =   1e-4
      End
      Begin VB.Label label 
         Caption         =   "CSS style for exporting Listview to html"
         Height          =   225
         Index           =   1
         Left            =   6615
         TabIndex        =   13
         Top             =   870
         Width           =   3240
      End
      Begin VB.Label label 
         Caption         =   "Max character per Field"
         Height          =   270
         Index           =   0
         Left            =   8400
         TabIndex        =   12
         Top             =   510
         Width           =   1860
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   7530
      Width           =   12045
      _ExtentX        =   21246
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13018
            Text            =   "Estatus"
            TextSave        =   "Estatus"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "22-03-2005"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "2:03"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   10186
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    c_lv.Progresbar = ProgressBar1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ListView.Height = Me.ScaleHeight - ListView.Top - 2565
    ListView.Width = Me.ScaleWidth - ListView.Left - 120
    pct.Top = ListView.Height + 560
End Sub

Private Sub opTbl_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    c_lv.MaxCharForFields = Val(Me.txtMaximo)
    Select Case Index
    Case 0
        rs.Open "Select * from Pedidos", cnx
        c_lv.Load_ListView ListView, rs, False, True
        rs.Close
    Case 1
        rs.Open "Select * from Detalles", cnx
        c_lv.Load_ListView ListView, rs, False, True
        rs.Close
    Case 2
        rs.Open "Select * from Productos", cnx
        c_lv.Load_ListView ListView, rs, False, True
        rs.Close
    End Select
    StatusBar.Panels(1).Text = "Got [ " & ListView.ListItems.Count & " ] records from " & opTbl(Index).Caption
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExportarCSV_Click()
    frm_ExpCSV.Show vbModal
End Sub

Private Sub cmdExportarHTML_Click()
    Screen.MousePointer = vbHourglass
    If c_lv.Export_LisviewToHTML(ListView, App.Path & "\TEXTO.html", "Creado Con Clase C_listview de seponceh@yahoo.com", Combo.Text & ".css") = True Then
        OpenFile App.Path & "\TEXTO.html", Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdExportarTXT_Click()
    Screen.MousePointer = vbHourglass
    If c_lv.Export_ListviewToTXT(App.Path & "\TEXTO.txt", ListView, True) = True Then
        OpenFile App.Path & "\TEXTO.txt", Me
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBar_Click()
    If cmdBar.Caption = "Disable ProgressBar" Then
        c_lv.Progresbar = Nothing
        cmdBar.Caption = "Enable ProgressBar"
    Else
        c_lv.Progresbar = ProgressBar1
        cmdBar.Caption = "Disable ProgressBar"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set cnx = Nothing
    Set c_lv = Nothing
End Sub

