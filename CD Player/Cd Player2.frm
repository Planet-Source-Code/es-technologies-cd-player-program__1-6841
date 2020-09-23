VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Credits"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   ControlBox      =   0   'False
   Icon            =   "Cd Player2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3300
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   11040
      Top             =   7560
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "exit"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton cmdCredits 
      Height          =   195
      Left            =   11280
      TabIndex        =   2
      Top             =   7680
      Width           =   75
   End
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3970
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.TextBox txtScroll 
         Alignment       =   2  'Center
         BackColor       =   &H80000006&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   5415
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Cd Player2.frx":27A2
         Top             =   3240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 'Declare Variables
    
    Dim lLineCount As Long
    Dim lLineHeight As Long
    
        lLineCount = SendMessage(txtScroll.hwnd, EM_GETLINECOUNT, 0&, 0&)
        lLineHeight = TextHeight("TEST") 'Get the height of text in file
        txtScroll.Height = lLineHeight * lLineCount
        picScroll.Left = 0
        picScroll.Visible = True
        tmrScroll.Enabled = True
        Left = (Screen.Width - Width) \ 2       'will center form
        Top = (Screen.Height - Height) \ 2      'will center form
End Sub

Private Sub picScroll_GotFocus()
    cmdExit.SetFocus
End Sub

Private Sub tmrScroll_Timer()
  
    If txtScroll.Top + txtScroll.Height < picScroll.Top Then 'picScroll.Top
        txtScroll.Top = picScroll.Height
    Else
        txtScroll.Top = txtScroll.Top - 15
    End If
End Sub

Private Sub txtScroll_GotFocus()
    
    cmdExit.SetFocus
    
End Sub


