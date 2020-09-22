VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Faster Strings with VB, NO API DECLARES"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Text Box"
      Height          =   285
      Left            =   3262
      TabIndex        =   7
      Top             =   1005
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2355
      Width           =   7830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add 20,000 Strings [BUFFERING]"
      Height          =   330
      Index           =   1
      Left            =   4260
      TabIndex        =   1
      Top             =   615
      Width           =   2670
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add 20,000 Strings [NORMAL]"
      Height          =   330
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   615
      Width           =   2670
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF80FF&
      Height          =   285
      Index           =   2
      Left            =   2700
      TabIndex        =   8
      Top             =   1935
      Width           =   2550
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Height          =   285
      Index           =   1
      Left            =   2707
      TabIndex        =   5
      Top             =   1650
      Width           =   2550
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Height          =   285
      Index           =   0
      Left            =   2707
      TabIndex        =   4
      Top             =   1365
      Width           =   2550
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Add a string to itself 20,000 times and see how long it takes! Start with NORMAL..."
      Height          =   270
      Left            =   735
      TabIndex        =   3
      Top             =   270
      Width           =   6345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "One of the biggest common mistakes in VB is when you need to join strings... Here are 2 ways to do it"
      Height          =   270
      Left            =   195
      TabIndex        =   2
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdClear_Click()
    Text1.Text = ""
End Sub
Private Sub Command1_Click(Index As Integer)
    
    Dim A     As String
    Dim B     As String
    Dim X     As Long
    Dim U     As Long
    Dim Time1 As Variant
    Dim Time2 As Variant
    
    Screen.MousePointer = vbHourglass
    Label3(0) = "Start: "
    Label3(1) = "End: "
    Label3(2) = "Time: "
    Text1.Text = ""
    A = "COMM/MIST"
    B = ""
    Text1 = ""
    DoEvents
        Select Case Index
            Case 0  ' USUAL METHOD
                Label3(0) = "Start: " & Time
                Time1 = Time
                For U = 1 To 20000
                    B = B & A
                    Me.Caption = "Faster Strings with VB, NO API  - " & Format$(Now(), "DDDD, MMMM d, yyyy h:nn:ss AMPM")
                'Next U  ' also, using Next instead of Next U saves > 1 sec!
                Next
            Label3(1) = "End: " & Time
            Time2 = Time
            Text1 = B
            Case 1 ' BUFFERING METHOD
                Label3(0) = "Start: " & Time
                Time1 = Time
                X = 1
                B = Space(Len(A) * 20000)
                For U = 1 To 20000
                    Mid(B, X, Len(A)) = A
                    X = X + Len(A)
                    Me.Caption = "Faster Strings with VB, NO API  - " & Format$(Now(), "DDDD, MMMM d, yyyy h:nn:ss AMPM")
                'Next U  ' also, using Next instead of Next U saves > 1 sec!
                Next
            Label3(1) = "End: " & Time
            Time2 = Time
            Text1 = B
        End Select
    X = DateDiff("s", Time1, Time2)
    Label3(2) = "Time: " & X & " sec"
    Screen.MousePointer = vbDefault
End Sub
Private Sub Form_Load()
    Top = (Screen.Height - Height) / 2
    Left = (Screen.Width - Width) / 2
End Sub
