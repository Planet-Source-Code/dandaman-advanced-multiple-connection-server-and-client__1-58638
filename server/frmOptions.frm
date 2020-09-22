VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picButton 
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picOptions 
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   2835
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin VB.Frame Frame1 
         Caption         =   "Packet Size:"
         Height          =   855
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2535
         Begin VB.CommandButton cmdChangePacket 
            Caption         =   "Change"
            Height          =   255
            Left            =   1560
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblPacketSize 
            AutoSize        =   -1  'True
            Caption         =   "4096"
            Height          =   195
            Left            =   840
            TabIndex        =   3
            Top             =   360
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Current:"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strPacketSize As String

Private Sub cmdChangePacket_Click()
strPacketSize = InputBox("Enter a new packet size." & vbCrLf & "Min: 146" & vbTab & vbTab & "Max: 4096", "Server Editor", iMaxSend)
If Trim(strPacketSize) = "" Then Exit Sub

If IsNumeric(strPacketSize) = False Then
    Msgbox2 "Error: Packet size must be numeric only.", vbExclamation, "Error"
    Exit Sub
End If

If (strPacketSize < 146) Or (strPacketSize > 4096) Then
    Msgbox2 "Error: Packet size incorrect.", vbExclamation, "Error"
Else
    lblPacketSize.Caption = strPacketSize
End If
End Sub

Private Sub cmdOK_Click()
If picOptions.Visible = True Then
    AddToLog "Packet size changed from: " & Options.MaxSend & " to: " & strPacketSize
    Options.MaxSend = strPacketSize
    SaveSetting Me.Caption, "main", "maxsend", Options.MaxSend
    Call frmMain.SendData("optmax" & Options.MaxSend)       'update all the users
End If

Unload Me 'unload the form
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

