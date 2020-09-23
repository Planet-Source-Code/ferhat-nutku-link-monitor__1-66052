VERSION 5.00
Begin VB.UserControl LabeledTextBox 
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6855
   ScaleHeight     =   330
   ScaleMode       =   0  'User
   ScaleWidth      =   6850
   Begin VB.TextBox txtText 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   5715
   End
   Begin VB.Label lbLabel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1020
   End
End
Attribute VB_Name = "LabeledTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Written by Ferhat Nutku yeniferhat@yahoo.com
' Copyright (c) 2005. All Rights Reserved.
'
' This code may be used in compiled form in any way you desire. This
' file may be redistributed unmodified by any means providing it is
' not sold for profit without the authors written consent, and
' providing that this notice and the authors name and all copyright
' notices remains intact.
' This file and the accompanying source code
' may not be hosted on a website or bulletin board without the author's
' written permission.
'
' This file is provided "as is" with no expressed or implied warranty.
' The author accepts no liability for any damage/loss of business that this product may cause.
'
' Last Updated: Aug. 10, 2006
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public Property Get LabelText() As Variant
   LabelText = lbLabel.Caption
End Property

Public Property Let LabelText(ByVal vNewValue As Variant)
   lbLabel.Caption = vNewValue
End Property

Public Property Get Text() As Variant
   Text = txtText.Text
End Property

Public Property Let Text(ByVal vNewValue As Variant)
   txtText.Text = vNewValue
End Property

Public Property Let ReadOnly(ByVal vNewValue As Variant)
   If (CBool(vNewValue)) Then
      txtText.Locked = True
   Else
      txtText.Locked = False
   End If
End Property


'''''''''''''
'''METHODS'''
'''''''''''''

Private Sub UserControl_Initialize()
   Me.LabelText = lbLabel.Caption
   Me.Text = txtText.Text
End Sub

