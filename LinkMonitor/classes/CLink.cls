VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLink"
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


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CLink
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub LinkTextBox(ByRef pTextBox As TextBox, _
                       plinkTopic As String, _
                       plinkMode As LinkModeConstants, _
                       plinkItem As Variant)
   'Links a textbox with a program
   
   'EXAMPLE:
   'Call goCLink.LinkTextBox(txtMonitor, "iexplore|WWW_GetWindowInfo", vbLinkManual, &HFFFFFFFF)
   
   'If could not find link item sustain listen
   On Local Error GoTo listen:
   
   pTextBox.LinkTopic = plinkTopic
   pTextBox.LinkTimeout = 10
   pTextBox.LinkMode = plinkMode
   pTextBox.LinkItem = plinkItem
   pTextBox.LinkRequest
   
listen:
   On Error GoTo 0
End Sub
