VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "NSC"
   ClientHeight    =   3630
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3885
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   1800
      Top             =   240
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      Height          =   3015
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   3615
      ExtentX         =   6376
      ExtentY         =   5318
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   3840
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line3 
      X1              =   3840
      X2              =   3840
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   0
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      Caption         =   "New Source Codes"
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
WebBrowser1.Navigate App.Path + "\NSC Ticker.html"
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Me.BorderStyle = 1
Me.WindowState = 1
End Sub

Private Sub Timer2_Timer()
If Me.WindowState = 0 Then Me.BorderStyle = 0
If Me.WindowState = 1 Then Me.BorderStyle = 1
End Sub
