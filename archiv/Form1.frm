VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   ">>>"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnMoveDown 
      Caption         =   "v"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton BtnMoveUp 
      Caption         =   "^"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   5325
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Chars As Collection

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub Form_Load()
    Set m_Chars = New Collection
    Dim i As Long: i = 65
    Dim c As Long: c = i + 25
    For i = i To c
        m_Chars.Add ChrW(i)
    Next
    UpdateView
End Sub

Private Sub BtnMoveUp_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then
        MsgBox "select item first"
        Exit Sub
    End If
    Col_MoveUp m_Chars, i + 1
    UpdateView i - 1
End Sub

Private Sub BtnMoveDown_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then
        MsgBox "select item first"
        Exit Sub
    End If
    Col_MoveDown m_Chars, i + 1
    UpdateView i + 1
End Sub

Sub UpdateView(Optional ByVal SelectedIndex)
    List1.Clear
    Dim i As Long
    For i = 1 To m_Chars.Count
        List1.AddItem m_Chars.Item(i)
    Next
    If IsMissing(SelectedIndex) Then Exit Sub
    If List1.ListCount <= SelectedIndex Then Exit Sub
    List1.ListIndex = SelectedIndex
End Sub
