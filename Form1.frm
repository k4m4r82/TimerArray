VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Demo Timer Array"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1000
      Left            =   4800
      Top             =   1080
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5106
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim counter() As Integer

Private Sub cmdStart_Click()
    Dim Index As Integer
    
    Index = ListView1.SelectedItem.Index - 1
    counter(Index) = 1
    Timer1(Index).Enabled = True
End Sub

Private Sub Form_Load()
    Dim i       As Integer
    Dim Index   As Integer
    
    'inisialisasi listview
    With ListView1
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
       
        .ColumnHeaders.Add , , "Item", 1000
        .ColumnHeaders.Add , , "Counter", 1000
    End With
        
    ReDim counter(9)
    For i = 1 To 10
        ListView1.ListItems.Add , , "Item #" & i
        
        If Index > 0 Then Load Timer1(Index)
        Index = Index + 1
    Next i
End Sub

Private Sub Timer1_Timer(Index As Integer)
    ListView1.ListItems(Index + 1).SubItems(1) = counter(Index)
    
    counter(Index) = counter(Index) + 1
End Sub
