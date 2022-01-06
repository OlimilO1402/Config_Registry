VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Config-Registry"
   ClientHeight    =   5355
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9255
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnDeleteEntry 
      Caption         =   "Delete Entry"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton BtnFileExists 
      Caption         =   "File Exists?"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton BtnWriteVBPRecentFiles 
      Caption         =   "Write vbp Recent Files to Registry"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton BtnReadVBPRecentFiles 
      Caption         =   "Read vbp Recent Files from Registry"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   375
      Width           =   9135
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      ItemData        =   "Form1.frx":1782
      Left            =   0
      List            =   "Form1.frx":1784
      TabIndex        =   0
      Top             =   750
      Width           =   9135
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private regKey As String

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    regKey = "Software\Microsoft\Visual Basic\6.0\RecentFiles\"
End Sub

Private Sub Form_Resize()
    Dim l As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Text1.Height
    If W > 0 And H > 0 Then Text1.Move l, T, W, H
    T = List1.Top
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move l, T, W, H
End Sub

Private Sub BtnReadVBPRecentFiles_Click()
    Registry.RootKey = HKEY_CURRENT_USER
    If Not Registry.OpenKey(regKey, False) Then
        MsgBox "Could not open registry key: " & regKey
        Exit Sub
    End If
    List1.Clear
    Dim i As Long, s As String
    For i = 1 To 50
        'If Registry.ValueExists(CStr(i)) Then
            s = Registry.ReadString(CStr(i))
            List1.AddItem s
        'End If
    Next
    Registry.CloseKey
End Sub
    
Private Sub BtnWriteVBPRecentFiles_Click()
    Registry.RootKey = HKEY_CURRENT_USER
    If Not Registry.OpenKey(regKey, True) Then
        MsgBox "Could not open registry key: " & regKey
        Exit Sub
    End If
    Dim i As Long, c As Long: c = List1.ListCount
    For i = 0 To c - 1
        Registry.WriteString CStr(i + 1), List1.List(i)
    Next
Try: On Error GoTo Catch
    If c < 50 Then
        For i = c To 50
            Registry.WriteString CStr(i + 1), "" 'List1.List(i)
            'If Not Registry.DeleteValue(CStr(i)) Then
            '
            'End If
        Next
    End If
    Resume Finally
Catch:
Finally:
    Registry.CloseKey
End Sub

Private Sub BtnFileExists_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(List1.List(i))
    If pfn.IsPath Then
        If pfn.PathExists Then
            MsgBox "Yes, path does exist:" & vbCrLf & pfn.Value
        End If
    Else
        If pfn.Exists Then
            MsgBox "Yes, file does exist:" & vbCrLf & pfn.Value
        Else
            MsgBox "No, it does not exist:" & vbCrLf & pfn.Value
        End If
    End If
End Sub

Private Sub BtnDeleteEntry_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Dim pfn As String: pfn = List1.List(i)
    If MsgBox("Are you sure to delete this entry?" & vbCrLf & pfn) <> vbOK Then Exit Sub
    List1.RemoveItem i
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim i As Long: i = List1.ListIndex
        If i < 0 Then
            List1.AddItem Text1.Text
        Else
            List1.List(i) = Text1.Text
        End If
    End If
End Sub

Private Sub List1_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then Exit Sub
    Text1.Text = List1.List(i)
End Sub

'Private Function IsPath(pfn As String) As Boolean
'Try: On Error GoTo Catch
'    IsPath = GetAttr(pfn) = vbDirectory Or vbVolume
'Catch:
'End Function
'
'Private Function FileExists(pfn As String) As Boolean
'Try: On Error GoTo Catch
'    FileExists = GetAttr(pfn) <> 0
'Catch:
'End Function
'

