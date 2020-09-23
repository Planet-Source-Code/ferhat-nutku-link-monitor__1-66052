VERSION 5.00
Begin VB.UserControl LabeledComboBox 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   ScaleHeight     =   360
   ScaleWidth      =   3585
   Begin VB.ComboBox cmbComboBox 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1019
   End
End
Attribute VB_Name = "LabeledComboBox"
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

Dim goControl As New CControl


Public Property Get LabelText() As Variant
   LabelText = goLabelText
End Property

Public Property Let LabelText(ByVal vNewValue As Variant)
   lblLabel.Caption = vNewValue
End Property

Private Sub UserControl_Initialize()
   Me.LabelText = lblLabel.Caption
End Sub

Public Property Get SeletedValue() As Variant
   SeletedValue = cmbComboBox.List(cmbComboBox.ListIndex)
End Property

Public Property Let SeletedValue(ByVal vNewValue As Variant)
   Dim size As Integer
   size = cmbComboBox.ListCount - 1
   
   For i = 0 To size
      If (cmbComboBox.List(i) = vNewValue) Then
         cmbComboBox.ListIndex = i
      End If
   Next
End Property

Public Property Get SelectedIndex() As Variant
   SelectedIndex = cmbComboBox.ListIndex
End Property

Public Property Let SelectedIndex(ByVal vNewValue As Variant)
   cmbComboBox.ListIndex = vNewValue
End Property


'''''''''''''
'''METHODS'''
'''''''''''''

Public Sub FillComboBoxFromComboBox(ByVal pSource As ComboBox)
   Call goControl.CopyComboBox(pSource, cmbComboBox)
End Sub
