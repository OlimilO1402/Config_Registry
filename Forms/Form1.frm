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
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      TabIndex        =   0
      ToolTipText     =   "Press Del to delete, drag'n'drop vbp-files from file-explorer"
      Top             =   720
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileRecentVBPFilesRead 
         Caption         =   "Read Recent VBP-Files From Registry"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileRecentVBPFilesWrite 
         Caption         =   "Write Recent VBP-Files To Registry"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExists 
         Caption         =   "Exists?"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSelectedWrite 
         Caption         =   "Write Selected To Registry"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Delete From List"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileHowManyDeleted 
         Caption         =   "How Many Deleted Files"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuFileAddDelAtEnd 
         Caption         =   "Add All Deleted Files"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopyListToCB 
         Caption         =   "&Copy List To ClipBoard"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditMoveUp 
         Caption         =   "Move &Up ^"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuEditMoveDown 
         Caption         =   "Move &Down v"
         Shortcut        =   ^D
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

Private Sub List1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    Dim i As Long: i = List1.ListIndex
    Dim s As String: s = List1.List(i)
    s = ParseFileName(s)
    Data.Clear
    Data.Files.Add s
    Data.SetData , vbCFFiles
End Sub

Sub UpdateData(ByVal i As Integer, ByVal sPFN As String)
    'If i < 1 And 50 < i Then Exit Sub 'i is outside range
    Dim PFN As PathFileName: Set PFN = m_VbpFiles.Item(i)
    PFN.Value = sPFN
    UpdateView i, PFN
End Sub

Sub UpdateView(Optional ByVal i As Integer = -1, Optional ByVal PFN As PathFileName = Nothing)
    If 1 <= i And i <= 50 Then
        If PFN Is Nothing Then Set PFN = m_VbpFiles.Item(i)
        List1.List(i - 1) = Format(i, PFN)
    Else
        List1.Clear
        For i = 1 To m_VbpFiles.Count
            Set PFN = m_VbpFiles.Item(i)
            List1.AddItem Format(i, PFN)
        Next
    End If
End Sub

Function Format(ByVal i As Integer, ByVal PFN As PathFileName) As String
    Format = Int_ToStr2(i) & ": | " & IIf(PFN.Exists, " Yes  ", "  No  ") & " | " & PFN.Value
End Function
Function Int_ToStr2(i As Integer) As String
    Int_ToStr2 = CStr(i): If Len(Int_ToStr2) < 2 Then Int_ToStr2 = "0" & Int_ToStr2
End Function

' v ############################## v '    mnuFile    ' v ############################## v '
Private Sub mnuFileRecentVBPFilesRead_Click()
Try: On Error GoTo Catch
    Registry.RootKey = HKEY_CURRENT_USER
    If Not Registry.OpenKey(m_regKey, False) Then
        MsgBox "Could not open registry key: " & m_regKey
        Exit Sub
    End If
    List1.Clear
    Set m_VbpFiles = New Collection
    'Set m_DelFiles = New Collection
    Dim i As Integer, s As String, PFN As PathFileName
    Dim c As Integer: c = 100 'Registry.GetKeyCount
    Do
        i = i + 1
        s = Registry.ReadString(CStr(i))
        If Len(s) Then
            Set PFN = MNew.PathFileName(s)
            m_VbpFiles.Add PFN
        End If
    Loop Until i >= c
    GoTo Finally
Catch:
    MsgBox "Error"
Finally:
    Registry.CloseKey
    UpdateView
End Sub

Private Sub mnuFileRecentVBPFilesWrite_Click()
Try: On Error GoTo Catch
    Registry.RootKey = HKEY_CURRENT_USER
    If Not Registry.OpenKey(m_regKey, True) Then
        MsgBox "Could not open registry key: " & m_regKey
        Exit Sub
    End If
    Dim s As String, PFN As PathFileName
    Dim i As Long, c As Long: c = m_VbpFiles.Count
    For i = 1 To c '- 1
        Set PFN = m_VbpFiles.Item(i)
        s = PFN.Value
        Registry.WriteString CStr(i), s
    Next
    GoTo Finally
Catch:
    ErrHandler "mnuRegistryRecentVBPFilesWrite", "i: " & i & " c: " & c & vbCrLf & "RegKey: " & m_regKey & vbCrLf & s
Finally:
    Registry.CloseKey
End Sub

Private Sub mnuFileExists_Click()
    If List1.ListCount = 0 Then
        MsgBox "The list is empty, first click the menu item:" & vbCrLf & """" & mnuFile.Caption & """ -> """ & mnuFileRecentVBPFilesRead.Caption & """"
        Exit Sub
    End If
    Dim s As String
    s = Text1.Text
    Dim PFN As PathFileName: Set PFN = MNew.PathFileName(s)
    If PFN.IsPath Then
        If PFN.PathExists Then
            MsgBox "Yes, path does exist:" & vbCrLf & PFN.Value
        End If
    Else
        If PFN.Exists Then
            MsgBox "Yes, file does exist:" & vbCrLf & PFN.Value
        Else
            MsgBox "No, file does not exist:" & vbCrLf & PFN.Value
        End If
    End If
End Sub

Private Sub mnuFileSelectedWrite_Click()
Try: On Error GoTo Catch
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then
        MsgBox "Select item first"
        Exit Sub
    End If
    Dim s As String, PFN As PathFileName: Set PFN = m_VbpFiles.Item(i + 1)
    Registry.RootKey = HKEY_CURRENT_USER
    If Not Registry.OpenKey(m_regKey, True) Then
        MsgBox "Could not open registry key: " & m_regKey
        Exit Sub
    End If
    s = PFN.Value
    Registry.WriteString CStr(i + 1), s
    GoTo Finally
Catch:
    ErrHandler "mnuFileSelectedWrite_Click", "i: " & i & vbCrLf & vbCrLf & s
Finally:
    Registry.CloseKey
End Sub

Private Sub mnuFileDelete_Click()
    List1_KeyDown KeyCodeConstants.vbKeyDelete, 0
End Sub

Private Sub mnuFileHowManyDeleted_Click()
    MsgBox "Number of entries in the list of deleted files: " & m_DelFiles.Count
End Sub

Private Sub mnuFileAddDelAtEnd_Click()
    Dim i As Long, c As Long: c = m_DelFiles.Count
    If c = 0 Then
        MsgBox "The list of deleted files is empty"
        Exit Sub
    End If
    Dim PFN As PathFileName
    For i = 1 To m_DelFiles.Count '- 1
        Set PFN = m_DelFiles.Item(i)
        m_VbpFiles.Add PFN
    Next
    UpdateView
End Sub

Private Sub mnuFileExit_Click()
    'ask? - just quit
    Unload Me
End Sub
' ^ ############################## ^ '    mnuFile    ' ^ ############################## ^ '

' v ############################## v '    mnuEdit    ' v ############################## v '
Private Sub mnuEditCopyListToCB_Click()
    Dim i As Long, c As Long: c = m_VbpFiles.Count
    Dim s As String, PFN As PathFileName
    For i = 1 To m_VbpFiles.Count
        Set PFN = m_VbpFiles.Item(i)
        s = s & PFN.Value & vbCrLf
    Next
    Clipboard.SetText s
End Sub

Private Sub mnuEditMoveUp_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then
        MsgBox "Select item first"
        Exit Sub
    End If
    Col_MoveUp m_VbpFiles, i + 1 'collections are 1-based
    UpdateView
    If i < 0 Or List1.ListCount <= i + 1 Then Exit Sub
    List1.ListIndex = i - 1
End Sub

Private Sub mnuEditMoveDown_Click()
    Dim i As Long: i = List1.ListIndex
    If i < 0 Then
        MsgBox "Select item first"
        Exit Sub
    End If
    Col_MoveDown m_VbpFiles, i + 1 'collections are 1-based
    UpdateView
    If i < 0 Or List1.ListCount <= (i + 1) Then Exit Sub
    List1.ListIndex = i + 1
End Sub
' ^ ############################## ^ '    mnuEdit    ' ^ ############################## ^ '

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii <> KeyCodeConstants.vbKeyReturn Then Exit Sub
    KeyAscii = 0
    Dim s As String, m As String
    Dim c As Integer: c = List1.ListCount: m = m & IIf(c = 0, "The list ist empty. ", "")
    Dim i As Integer: i = List1.ListIndex: m = m & IIf(i < 0, "Nothing is selected. ", "")
    If i < 0 Then
        s = Text1.Text
        If MsgBox(m & vbCrLf & "Do you want to add this entry? " & vbCrLf & s) = vbCancel Then Exit Sub
        Dim PFN As PathFileName: Set PFN = MNew.PathFileName(s)
        m_VbpFiles.Add PFN
        UpdateView
    Else
        UpdateData i + 1, Text1.Text
    End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case KeyCodeConstants.vbKeyDelete
        Dim i As Integer: i = List1.ListIndex
        If List1.ListCount = 0 Then
            MsgBox "The list is empty, first click the menu item:" & vbCrLf & """" & mnuFile.Caption & """ -> """ & mnuFileRecentVBPFilesRead.Caption & """"
            Exit Sub
        End If
        'we save every object in the list of deleted files befor we delete from list
        Dim PFN As PathFileName
        Set PFN = m_VbpFiles.Item(i + 1)
        m_DelFiles.Add PFN
        m_VbpFiles.Remove i + 1
        
        UpdateView
        If List1.ListCount <= i Then i = List1.ListCount - 1
        'set selection again
        List1.ListIndex = i
    End Select
End Sub

Private Sub List1_Click()
    Text1.Text = ParseFileName(List1.Text)
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim PFN As PathFileName: Set PFN = MNew.PathFileName(Data.Files(1))
    If Not PFN.Exists Then Exit Sub
    Dim ext As String: ext = LCase(PFN.Extension)
    If ext <> ".vbp" Then MsgBox "Wrong fileformat, only vbp-files": Exit Sub
    m_VbpFiles.Add PFN
    UpdateView
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

'copy this same function to every class, form or module
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function


