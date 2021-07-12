VERSION 5.00
Begin VB.Form EditDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report Attribute"
   ClientHeight    =   2370
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CloseBtn 
      Caption         =   "Close"
      Height          =   360
      Left            =   1860
      TabIndex        =   5
      Top             =   1860
      Width           =   1065
   End
   Begin VB.TextBox Edit2 
      Height          =   960
      Left            =   255
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "EditDialog.frx":0000
      Top             =   750
      Width           =   4260
   End
   Begin VB.TextBox Edit1 
      Height          =   330
      Left            =   262
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   750
      Width           =   4260
   End
   Begin VB.CommandButton CancelBtn 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2655
      TabIndex        =   1
      Top             =   1860
      Width           =   1065
   End
   Begin VB.CommandButton OKBtn 
      Caption         =   "OK"
      Height          =   360
      Left            =   1072
      TabIndex        =   0
      Top             =   1860
      Width           =   1065
   End
   Begin VB.Label Text 
      Caption         =   "Label1"
      Height          =   480
      Left            =   345
      TabIndex        =   2
      Top             =   225
      Width           =   4095
   End
End
Attribute VB_Name = "EditDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lResult As Boolean
Private cResult As String

Public Function InitParams(cName As String, cValue As String, lMLE As Boolean, lReadOnly As Boolean) As Boolean
    
    ' Is the attribute read only?
    If lReadOnly Then
        Text.Caption = "View Attribute:   " + cName
        OKBtn.Visible = False
        CancelBtn.Visible = False
    Else
        Text.Caption = "Edit Attribute:   " + cName
        CloseBtn.Visible = False
    End If
    
    ' Show the multiline edit?
    If lMLE Then
        Edit1.Visible = False
        Edit2.Text = cValue
    Else
        Edit2.Visible = False
        Edit1.Text = cValue
    End If
    
    lResult = False
    
    Show vbModal
    
    If lResult Then
        cValue = cResult
    End If
   
    InitParams = lResult

End Function

Private Sub CancelBtn_Click()
    lResult = False
    Unload Me
End Sub

Private Sub CloseBtn_Click()
    lResult = False
    Unload Me
End Sub

Private Sub OKBtn_Click()
    lResult = True
    If Edit1.Visible Then
        cResult = Edit1.Text
    Else
        cResult = Edit1.Text
    End If
    
    Unload Me
End Sub
