VERSION 5.00
Begin VB.Form frmSameLinks 
   Caption         =   "Same Links"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5565
   Icon            =   "frmSameLinks.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   5565
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.ListBox lbSameLinks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmSameLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    Unload Me
End Sub
