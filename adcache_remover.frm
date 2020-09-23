VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "AdCache Remover"
   ClientHeight    =   2055
   ClientLeft      =   8475
   ClientTop       =   6090
   ClientWidth     =   3270
   Icon            =   "adcache_remover.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3270
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   1920
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "Standby"
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2880
      Top             =   1800
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   2955
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   3015
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Version: 1.0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   120
         Picture         =   "adcache_remover.frx":0442
         Stretch         =   -1  'True
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "AdCache Remover"
         BeginProperty Font 
            Name            =   "Orange LET"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.Label lblCaptionTime 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblStatusTime 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This program is freeware and is NOT COVERED BY ANY WARRANTY"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblTime 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdQuit_Click()
    End
End Sub

Private Sub txtStatus_Change()
If frmMain.WindowState = 1 Then
    Let Timer2.Enabled = True
Else
    Let Timer2.Enabled = False
    Let frmMain.Caption = "AdCache Remover"
End If
End Sub

Private Sub Form_Load()
Let lblTime.Caption = 0
Let Timer1.Enabled = True
Let Timer2.Enabled = False
End Sub

Private Sub Timer1_Timer()
Let lblTime.Caption = Str((lblTime.Caption) + 1)
Let lblStatusTime.Caption = Str((lblStatusTime.Caption) + 1)
On Error Resume Next
        x% = Len(Dir$("c:\Windows\System\AdCache\*.*")) 'If you have a different directory for the AdCache change it.
    If Err Or x% = 0 Then fileexists = False Else fileexists = True
If lblTime.Caption = 1 Then
    If fileexists = True Then
        Kill ("C:\Windows\System\AdCache\*.*") 'If you have a different directory for the AdCache change it.
        Let txtStatus.Text = "Deleting AdCache"
        Let lblTime.Caption = 0
    Else
        Let txtStatus.Text = "Standby"
        Let lblTime.Caption = 0
    End If
End If
If lblStatusTime.Caption = 1 Then
    Let txtStatus.Text = txtStatus.Text & "."
ElseIf lblStatusTime.Caption = 2 Then
    Let txtStatus.Text = txtStatus.Text & ".."
ElseIf lblStatusTime.Caption = 3 Then
    Let txtStatus.Text = txtStatus.Text & "..."
ElseIf lblStatusTime.Caption = 4 Then
    Let txtStatus.Text = txtStatus.Text & "...."
ElseIf lblStatusTime.Caption = 5 Then
    Let txtStatus.Text = txtStatus.Text & "....."
ElseIf lblStatusTime.Caption = 6 Then
    Let txtStatus.Text = txtStatus.Text & "......"
ElseIf lblStatusTime.Caption = 7 Then
    Let txtStatus.Text = txtStatus.Text & "......."
ElseIf lblStatusTime.Caption = 8 Then
    Let txtStatus.Text = txtStatus.Text & "........"
ElseIf lblStatusTime.Caption = 9 Then
    Let txtStatus.Text = txtStatus.Text & "........."
ElseIf lblStatusTime.Caption = 10 Then
    Let txtStatus.Text = txtStatus.Text & ".........."
ElseIf lblStatusTime.Caption = 11 Then
    Let txtStatus.Text = txtStatus.Text & "..........."
ElseIf lblStatusTime.Caption = 12 Then
    Let txtStatus.Text = txtStatus.Text & "............"
ElseIf lblStatusTime.Caption = 13 Then
    Let txtStatus.Text = txtStatus.Text & "............."
ElseIf lblStatusTime.Caption = 14 Then
    Let txtStatus.Text = txtStatus.Text & ".............."
ElseIf lblStatusTime.Caption = 15 Then
    Let txtStatus.Text = txtStatus.Text & "..............."
ElseIf lblStatusTime.Caption = 16 Then
    Let txtStatus.Text = txtStatus.Text & "................"
ElseIf lblStatusTime.Caption = 17 Then
    Let txtStatus.Text = txtStatus.Text & "................."
ElseIf lblStatusTime.Caption = 18 Then
    Let txtStatus.Text = txtStatus.Text & ".................."
ElseIf lblStatusTime.Caption = 19 Then
    Let txtStatus.Text = txtStatus.Text & "..................."
ElseIf lblStatusTime.Caption = 20 Then
    Let txtStatus.Text = txtStatus.Text & "...................."
ElseIf lblStatusTime.Caption = 21 Then
    Let txtStatus.Text = txtStatus.Text & ""
    Let lblStatusTime.Caption = 0
End If
End Sub

Private Sub Timer2_Timer()
Let lblCaptionTime.Caption = Str((lblCaptionTime.Caption) + 1)
If lblCaptionTime.Caption = 2 Then
    Let lblCaptionTime.Caption = 0
End If
If lblCaptionTime.Caption = 0 Then
    Let frmMain.Caption = "AdCache Remover"
ElseIf lblCaptionTime.Caption = 1 Then
    Let frmMain.Caption = txtStatus.Text
End If
End Sub
