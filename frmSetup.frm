VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bot Setup"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Idle Setup"
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   3735
      Begin VB.TextBox txtInterval 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   600
         Width           =   495
      End
      Begin VB.OptionButton OptOff 
         Caption         =   "Off"
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton OptOn 
         Caption         =   "On"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtIdle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Every"
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Minutes"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "[Vapor Bot]"
         Height          =   255
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login Settings"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Server"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Password"
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
WriteStuff "Login", "User", txtName.text
WriteStuff "Login", "Pass", txtPass.text
WriteStuff "Login", "Server", txtServer.text
WriteStuff "Idle", "Message", txtIdle.text
WriteStuff "Idle", "Interval", txtInterval.text
If OptOn = True Then
    WriteStuff "Idle", "Idle", "True"
Else
    WriteStuff "Idle", "Idle", "False"
End If
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtName.text = GetStuff("Login", "User")
txtPass.text = GetStuff("Login", "Pass")
txtServer.text = GetStuff("Login", "Server")
txtIdle.text = GetStuff("Idle", "Message")
txtInterval.text = GetStuff("Idle", "Interval")
If GetStuff("Idle", "Idle") = "True" Then
    OptOn.Value = True
    OptOff.Value = False
Else
    OptOn.Value = False
    OptOff.Value = True
End If
End Sub
