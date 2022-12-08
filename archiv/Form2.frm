VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   5550
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnMoveDown 
      Caption         =   "v"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
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
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Chars As Collection

Private Sub Form_Load()
    Set m_Chars = New Collection
    Dim i As Integer: i = 65
    Dim c As Integer: c = i + 25
    Dim ch As Char
    For i = i To c
        m_Chars.Add Char(i)
    Next
    UpdateView
End Sub

Private Function Char(ByVal aCharW As Integer) As Char
    Set Char = New Char: Char.New_ aCharW
End Function

Private Sub BtnMoveUp_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then
        MsgBox "select item first"
        Exit Sub
    End If
    i = i + 1
    Col_SwapItems m_Chars, i, i - 1
    UpdateView i - 2
End Sub

Private Sub BtnMoveDown_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then
        MsgBox "select item first"
        Exit Sub
    End If
    i = i + 1
    Col_SwapItems m_Chars, i, i + 1
    UpdateView i '- 2
End Sub

Sub UpdateView(Optional ByVal SelectedIndex)
    List1.Clear
    Dim i As Integer, ch As Char
    For i = 1 To m_Chars.Count
        Set ch = m_Chars.Item(i)
        List1.AddItem ch.ToStr
    Next
    If IsMissing(SelectedIndex) Then Exit Sub
    If List1.ListCount <= SelectedIndex Then Exit Sub
    List1.ListIndex = SelectedIndex
End Sub

