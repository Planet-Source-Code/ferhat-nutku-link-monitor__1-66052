VERSION 5.00
Begin VB.Form frmCategories 
   Caption         =   "Select Categories"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   3195
   Icon            =   "frmCategories.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnSelectAll 
      Caption         =   "Select All"
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
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   2400
      Width           =   735
   End
   Begin VB.ListBox lbCategories 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Global Variables
'Global Variables
Private bListBoxItemSelected As Boolean
Public selectedCategories As String

Private Sub btnOK_Click()
    'Check if any category is selected
    Dim removeLength As Integer
    
    If lbCategories.SelCount = 0 Then
        Call MsgBox("Please select at least one category from the list." & strAddress, vbOKOnly + vbExclamation, MApplication.ProductNameVersion + " - Category Selection Warning")
    Else
        For i = 0 To lbCategories.ListCount - 1
            If lbCategories.Selected(i) = True Then
                selectedCategories = selectedCategories & "'" & lbCategories.List(i) & "'" & " or Category="
            End If
        Next
        
        removeLength = Len(selectedCategories) - 13
        selectedCategories = Mid(selectedCategories, 1, removeLength)
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim sSql As String
    bListBoxItemSelected = False
    
    'First clear listbox
    lbCategories.Clear
   
    'Fill listBox
    sSql = "SELECT DISTINCT Category FROM Links"
    Call frmMain.goControl.FillListBox(lbCategories, frmMain.gDBName, frmMain.goRecordset, sSql)
End Sub

Private Sub btnSelectAll_Click()
    'Select and deselect all listbox items
    
    If bListBoxItemSelected = False Then
        For i = 0 To lbCategories.ListCount - 1
            lbCategories.Selected(i) = Not bListBoxItemSelected
        Next
        bListBoxItemSelected = True
        btnSelectAll.Caption = "Select None"
    Else
        For i = 0 To lbCategories.ListCount - 1
            lbCategories.Selected(i) = Not bListBoxItemSelected
        Next
        bListBoxItemSelected = False
        btnSelectAll.Caption = "Select All"
    End If
End Sub
