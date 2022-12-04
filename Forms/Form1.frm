VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Config-Registry"
   ClientHeight    =   12375
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14055
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   14055
   StartUpPosition =   3  'Windows-Standard
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
      TabIndex        =   1
      ToolTipText     =   "Return-key will take the changes"
      Top             =   0
      Width           =   9735
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
      Width           =   9735
   End
   Begin VB.Label LblCaption 
      Caption         =   " Nr | Exists | VBP-PathFileName"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   3495
   End
   Begin VB.Menu mnuRegistry 
      Caption         =   "Registry"
      Begin VB.Menu mnuRegistryRecentVBPFilesRead 
         Caption         =   "Read Recent VBP-Files"
      End
      Begin VB.Menu mnuRegistryRecentVBPFilesWrite 
         Caption         =   "Write Recent VBP-Files"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExists 
         Caption         =   "Exists?"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Delete From List"
      End
      Begin VB.Menu mnuFileAddDelAtEnd 
         Caption         =   "Add All Deleted Files"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_regKey   As String
Private m_VbpFiles As Collection 'Of String of VBP-files
Private m_DelFiles As Collection

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    m_regKey = "Software\Microsoft\Visual Basic\6.0\RecentFiles\"
    Set m_VbpFiles = New Collection
    Set m_DelFiles = New Collection
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Text1.Height
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
    T = List1.Top
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then List1.Move L, T, W, H
End Sub

Private Sub mnuRegistryRecentVBPFilesRead_Click()
Try: On Error GoTo Catch
    Registry.RootKey = HKEY_CURRENT_USER
    If Not Registry.OpenKey(m_regKey, False) Then
        MsgBox "Could not open registry key: " & m_regKey
        Exit Sub
    End If
    List1.Clear
    Dim i As Integer, s As String, pfn As PathFileName
    Dim c As Integer: c = Registry.GetKeyCount
    Do
        i = i + 1
        'Dim bV As Boolean: bV = Registry.ValueExistsNoClose(CStr(i))
        'If Registry.ValueExists(CStr(i)) Then
        'If bV Then
            s = Registry.ReadString(CStr(i))
            If Len(s) Then
                Set pfn = MNew.PathFileName(s)
                m_VbpFiles.Add pfn
            End If
            
            'Set pfn = MNew.PathFileName(s)
            's = Int_ToStr2(i) & ": | " & IIf(pfn.Exists, " Yes  ", "  No  ") & " | " & s
            'List1.AddItem Format(i, pfn)
        'End If
        Registry.CloseCurrentKey
    Loop Until i >= c
    GoTo Finally
Catch:
    MsgBox "Error"
Finally:
    Registry.CClose
    UpdateView
End Sub

Sub VbpFiles_Delete()
    
End Sub

Sub UpdateData(ByVal i As Integer, ByVal sPFN As String)
    If i < 1 And 50 < i Then Exit Sub 'i is outside range
    Dim pfn As PathFileName: Set pfn = m_VbpFiles.Item(i)
    pfn.Value = sPFN
    UpdateView i, pfn
End Sub

Sub UpdateView(Optional ByVal i As Integer = -1, Optional ByVal pfn As PathFileName = Nothing)
    If 1 <= i And i <= 50 Then
        If pfn Is Nothing Then Set pfn = m_VbpFiles.Item(i)
        List1.List(i - 1) = Format(i, pfn)
    Else
        List1.Clear
        For i = 1 To m_VbpFiles.Count
            Set pfn = m_VbpFiles.Item(i)
            List1.AddItem Format(i, pfn)
        Next
    End If
End Sub

Function Format(ByVal i As Integer, ByVal pfn As PathFileName) As String
    Format = Int_ToStr2(i) & ": | " & IIf(pfn.Exists, " Yes  ", "  No  ") & " | " & pfn.Value
End Function
Function Int_ToStr2(i As Integer) As String
    Int_ToStr2 = CStr(i): If Len(Int_ToStr2) < 2 Then Int_ToStr2 = "0" & Int_ToStr2
End Function

Private Sub mnuRegistryRecentVBPFilesWrite_Click()
    Registry.RootKey = HKEY_CURRENT_USER
    If Not Registry.OpenKey(m_regKey, True) Then
        MsgBox "Could not open registry key: " & m_regKey
        Exit Sub
    End If
    Dim i As Long, c As Long: c = List1.ListCount
    For i = 0 To c - 1
        Registry.WriteString CStr(i + 1), ParseFileName(List1.List(i))
    Next
Try: On Error GoTo Catch
    If c < 50 Then
        For i = c To 50
            Registry.WriteString CStr(i + 1), ""
        Next
    End If
    GoTo Finally
Catch:
Finally:
    Registry.CloseKey
End Sub

Private Sub mnuFileExists_Click()
    If List1.ListCount = 0 Then
        MsgBox "The list is empty, first click the menu item:" & vbCrLf & """" & mnuRegistry.Caption & """ -> """ & mnuRegistryRecentVBPFilesRead.Caption & """"
        Exit Sub
    End If
    Dim s As String
    'should we take the filename from the listbox,
    's = Text1.Text = ListBox_ParseFileName(List1)
    'or should we take it directly from the textbox?
    s = Text1.Text
    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(s)
    If pfn.IsPath Then
        If pfn.PathExists Then
            MsgBox "Yes, path does exist:" & vbCrLf & pfn.Value
        End If
    Else
        If pfn.Exists Then
            MsgBox "Yes, file does exist:" & vbCrLf & pfn.Value
        Else
            MsgBox "No, file does not exist:" & vbCrLf & pfn.Value
        End If
    End If
End Sub

Private Sub mnuFileDelete_Click()
    If List1.ListCount = 0 Then
        MsgBox "The list is empty, first click the menu item:" & vbCrLf & """" & mnuRegistry.Caption & """ -> """ & mnuRegistryRecentVBPFilesRead.Caption & """"
        Exit Sub
    End If
    List1_KeyDown KeyCodeConstants.vbKeyDelete, 0
'    Dim i As Long: i = List1.ListIndex
'    If i < 0 Then
'        MsgBox "Select entry first!"
'        Exit Sub
'    End If
'    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(List1.List(i))
'    If MsgBox("Are you sure to delete this entry?" & vbCrLf & pfn.Quoted, vbOKCancel) = vbCancel Then Exit Sub
'    If mDeletedFiles Is Nothing Then Set mDeletedFiles = New Collection
'    mDeletedFiles.Add pfn
'    List1.RemoveItem i
End Sub

Private Sub mnuFileAddDelAtEnd_Click()
    'If MsgBox("Add the deleted entries at the end of the list?", vbOKCancel) = vbCancel Then Exit Sub
    Dim c As Long, i As Long, s As String
    Do Until c = 50
        s = List1.List(c)
        c = c + 1
        For i = 0 To mDeletedFiles.Count - 1
            mDeletedFiles.Item (i)
            c = c + 1
        Next
    'Next
    Loop
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case KeyCodeConstants.vbKeyReturn
        Dim i As Integer: i = List1.ListIndex
        If i < 0 Then Exit Sub
        UpdateData i + 1, Text1.Text
        KeyAscii = 0
    'Case Else
    End Select
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyDelete
        Dim i As Integer: i = List1.ListIndex
        If i < 1 Or 50 < i Then Exit Sub 'out of bounds
        'm_VbpFiles.Remove i + 1
        'no we actually do not remove any object just
        '
        Dim pfn As PathFileName
        Set pfn = m_VbpFiles.Item(i)
        UpdateView
        If List1.ListCount <= i Then i = List1.ListCount - 1
        List1.ListIndex = i
    End Select
End Sub

Private Sub List1_Click()
    Text1.Text = ParseFileName(List1.Text)
End Sub

Function ParseFileName(ByVal s As String) As String
Try: On Error GoTo Catch
    'parses the filename in this context from the string
    Dim pos As Long: pos = InStr(1, s, "|")
    If pos > 0 Then pos = InStr(pos + 1, s, "|")
    If pos > 0 Then s = Mid(s, pos + 1)
    ParseFileName = Trim(s)
Catch:
End Function

'Private Sub BtnCheckNDelete_Click()
''   'BtnCheckNDelete.Caption = "Check If File Exists and ask to delete"
'    If mDeletedFiles Is Nothing Then Set mDeletedFiles = New Collection
'
'    Dim i As Long, FNm As String, pfn As PathFileName
'    Dim cn As Long, cd As Long
'    If List1.ListCount = 0 Then
'        MsgBox "Click the button first: " & vbCrLf & """" & BtnReadVBPRecentFiles.Caption & """"
'        Exit Sub
'    End If
'    For i = List1.ListCount - 1 To 0 Step -1
'        FNm = ParseFileName(List1.List(i))
'        If Len(FNm) Then
'            Set pfn = MNew.PathFileName(FNm)
'            If pfn.Exists Then
'                cn = cn + 1
'            Else
'                cd = cd + 1
'                If MsgBox("File does not exist, delete it from the list?" & vbCrLf & pfn.Value, vbOKCancel) = vbOK Then
'                    mDeletedFiles.Add pfn
'                    List1.RemoveItem i
'                End If
'            End If
'        End If
'    Next
'    If cd = 0 Then
'        MsgBox "All files are existing, nothing to delete!"
'    Else
'        MsgBox cd & " files are missing, now click the button" & vbCrLf & """" & BtnWriteVBPRecentFiles.Caption & """" & vbCrLf & "Or the button" & vbCrLf & BtnAddDeletedAtEnd.Caption
'    End If
'End Sub

