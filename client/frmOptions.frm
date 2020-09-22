VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picButton 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   2475
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   360
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picCustomize 
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Frame Frame2 
         Caption         =   "Events"
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
         Begin VB.CommandButton cmdColor 
            Height          =   255
            Index           =   2
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Width           =   255
         End
         Begin VB.CommandButton cmdColor 
            Height          =   255
            Index           =   3
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Connect:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Disconnect:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Name Color"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1935
         Begin VB.CommandButton cmdColor 
            Height          =   255
            Index           =   1
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdColor 
            Height          =   255
            Index           =   0
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "All other clients:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Your name's color:"
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   1305
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdColor_Click(Index As Integer)
With CommonDialog

Select Case Index
    Case Is = 0
        'your color
        .ShowColor
        If .Color <> -1 Then
            cmdColor(0).BackColor = .Color
            'Options.Color.Mine = .Color
        End If
    
    Case Is = 1
        'others' color
        .ShowColor
        If .Color <> -1 Then
            cmdColor(1).BackColor = .Color
            'Options.Color.Others = .Color
        End If
    
    Case Is = 2
        'on connect
        .ShowColor
        If .Color <> -1 Then
            cmdColor(2).BackColor = .Color
        End If

    Case Is = 3
        .ShowColor
        If .Color <> -1 Then
            cmdColor(3).BackColor = .Color
        End If

End Select
End With
End Sub

Private Sub cmdOK_Click()
If picCustomize.Visible = True Then
    'set the colors
    Options.Color.Mine = cmdColor(0).BackColor
    Options.Color.Others = cmdColor(1).BackColor
    Options.Color.Connect = cmdColor(2).BackColor
    Options.Color.Disconnect = cmdColor(3).BackColor
    Call SaveSetting(AppName, "main", "mycolor", Options.Color.Mine)
    Call SaveSetting(AppName, "main", "othercolor", Options.Color.Others)
    Call SaveSetting(AppName, "main", "connectcolor", Options.Color.Connect)
    Call SaveSetting(AppName, "main", "disconnectcolor", Options.Color.Disconnect)
End If

'unload the form
Unload Me
End Sub

