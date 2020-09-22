VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "ucListView - Test"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   2  'CenterScreen
   Begin LV.ucListView ucListView1 
      Height          =   5535
      Left            =   2280
      TabIndex        =   21
      Top             =   240
      Width           =   6735
      _extentx        =   11880
      _extenty        =   9763
   End
   Begin VB.PictureBox picSetup 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   7335
      Left            =   0
      ScaleHeight     =   7335
      ScaleWidth      =   2085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2085
      Begin VB.CheckBox chkHeaderFixedWidth 
         Caption         =   "HeaderFixedWidth"
         Height          =   285
         Left            =   195
         TabIndex        =   17
         Top             =   5970
         Width           =   1815
      End
      Begin VB.CheckBox chkHeaderHide 
         Caption         =   "HeaderHide"
         Height          =   285
         Left            =   195
         TabIndex        =   19
         Top             =   6540
         Width           =   1815
      End
      Begin VB.CheckBox chkUnderlineHot 
         Caption         =   "UnderlineHot"
         Height          =   285
         Left            =   465
         TabIndex        =   15
         Top             =   5310
         Width           =   1695
      End
      Begin VB.CheckBox chkLabelEdit 
         Caption         =   "LabelEdit"
         Height          =   225
         Left            =   195
         TabIndex        =   5
         Top             =   2190
         Width           =   975
      End
      Begin VB.CheckBox chkTrackSelect 
         Caption         =   "TrackSelect"
         Height          =   285
         Left            =   465
         TabIndex        =   14
         Top             =   5010
         Width           =   1815
      End
      Begin VB.CheckBox chkHideSelection 
         Caption         =   "HideSelection"
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   2580
         Width           =   1815
      End
      Begin VB.CheckBox chkOneClickActivate 
         Caption         =   "OneClickActivate"
         Height          =   285
         Left            =   195
         TabIndex        =   13
         Top             =   4695
         Width           =   1815
      End
      Begin VB.CheckBox chkLabelTips 
         Caption         =   "LabelTips"
         Height          =   285
         Left            =   195
         TabIndex        =   12
         Top             =   4395
         Width           =   1815
      End
      Begin VB.CommandButton cmdFill 
         Caption         =   "&Fill (100 items)"
         Height          =   435
         Left            =   195
         TabIndex        =   3
         Top             =   1005
         Width           =   1245
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   435
         Left            =   195
         TabIndex        =   4
         Top             =   1530
         Width           =   1245
      End
      Begin VB.CheckBox chkMultiselect 
         Caption         =   "Multiselect"
         Height          =   285
         Left            =   195
         TabIndex        =   7
         Top             =   2880
         Width           =   1815
      End
      Begin VB.ComboBox cbViewMode 
         Height          =   315
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   525
         Width           =   1695
      End
      Begin VB.CheckBox chkCheckBoxes 
         Caption         =   "CheckBoxes"
         Height          =   285
         Left            =   195
         TabIndex        =   8
         Top             =   3180
         Width           =   1815
      End
      Begin VB.CheckBox chkGridLines 
         Caption         =   "GridLines"
         Height          =   285
         Left            =   195
         TabIndex        =   10
         Top             =   3780
         Width           =   1815
      End
      Begin VB.CheckBox chkFullRowSelect 
         Caption         =   "FullRowSelect"
         Height          =   285
         Left            =   195
         TabIndex        =   11
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CheckBox chkSubItemImages 
         Caption         =   "SubItemImages"
         Height          =   285
         Left            =   195
         TabIndex        =   9
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox chkHeaderFlat 
         Caption         =   "HeaderFlat"
         Height          =   285
         Left            =   195
         TabIndex        =   18
         Top             =   6255
         Width           =   1815
      End
      Begin VB.CheckBox chkScrollBarFlat 
         Caption         =   "ScrollBarFlat"
         Height          =   285
         Left            =   195
         TabIndex        =   20
         Top             =   6825
         Width           =   1815
      End
      Begin VB.ComboBox cbBorderStyle 
         Height          =   315
         Left            =   195
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   150
         Width           =   1695
      End
      Begin VB.CheckBox chkHeaderDragDrop 
         Caption         =   "HeaderDragDrop"
         Height          =   285
         Left            =   195
         TabIndex        =   16
         Top             =   5685
         Width           =   1815
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const sCHARS As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz "

Private m_ColumnSortOrder(2) As eSortOrderConstants
Private m_CurrentColumn      As Integer



Private Sub Form_Load()

    With cbBorderStyle
        .AddItem "0 - bsNone"
        .AddItem "1 - bsThin"
        .AddItem "2 - bsThick"
    End With
    With cbViewMode
        .AddItem "0 - vmIcon"
        .AddItem "1 - vmDetails"
        .AddItem "2 - vmSmallIcon"
        .AddItem "3 - vmList"
    End With
    
    With ucListView1
    
        Call .Initialize
        
        Call .InitializeImageListSmall
        Call .InitializeImageListLarge
        Call .InitializeImageListHeader
        Call .ImageListSmall_AddBitmap(LoadResPicture("IL16x16", vbResBitmap), vbMagenta)
        Call .ImageListLarge_AddBitmap(LoadResPicture("IL32x32", vbResBitmap), vbMagenta)
        Call .ImageListHeader_AddBitmap(LoadResPicture("ILHEADER", vbResBitmap), vbMagenta)
        
        Call .ColumnAdd(0, "Header 1", 150, [caLeft])
        Call .ColumnAdd(1, "Header 2", 100, [caRight])
        Call .ColumnAdd(2, "Header 3", 100, [caCenter])
        
        .RaiseSubItemPrePaint = True 'Force OnSubItemPrePaint() event
    End With
    
    cbBorderStyle.ListIndex = 2
    cbViewMode.ListIndex = 3
    
    Call Randomize(Timer)
    Call cmdFill_Click
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    Call ucListView1.Move(picSetup.Width + 10, 10, Me.ScaleWidth - (picSetup.Width + 20), Me.ScaleHeight - 20)
End Sub



Private Sub cbBorderStyle_Click()
    ucListView1.BorderStyle = cbBorderStyle.ListIndex
End Sub

Private Sub cbViewMode_Click()
    ucListView1.ViewMode = cbViewMode.ListIndex
End Sub

Private Sub cmdFill_Click()
            
  Dim nIdx As Integer
  Dim nCol As Integer
    
    With ucListView1
        .Visible = False
        For nCol = 0 To 2
            m_ColumnSortOrder(nCol) = [soDefault] '~None
            .ColumnIcon(nCol) = -1                '~None
        Next nCol
        m_CurrentColumn = -1
        For nIdx = .Count To .Count + 99
            Call .ItemAdd(nIdx, pvRandomString(3), 0, 1)
            Call .SubItemSet(nIdx, 1, Format(-50 + Rnd * 100, "0.0000"), 0)
            Call .SubItemSet(nIdx, 2, DateSerial(2004, Rnd * 11 + 1, Rnd * 30 + 1), 0)
        Next nIdx
        .Visible = True
    End With
End Sub

Private Sub cmdClear_Click()
  
  Dim nCol As Integer
  
    For nCol = 0 To 2
        m_ColumnSortOrder(nCol) = [soDefault] '~None
        ucListView1.ColumnIcon(nCol) = -1     '~None
    Next nCol
    m_CurrentColumn = -1
    Call ucListView1.Clear
End Sub



Private Sub chkLabelEdit_Click()
    ucListView1.LabelEdit = CBool(chkLabelEdit)
End Sub

Private Sub chkHideSelection_Click()
    ucListView1.HideSelection = CBool(chkHideSelection)
End Sub

Private Sub chkMultiselect_Click()
    ucListView1.MultiSelect = CBool(chkMultiselect)
End Sub

Private Sub chkCheckBoxes_Click()
    ucListView1.CheckBoxes = CBool(chkCheckBoxes)
End Sub

Private Sub chkSubItemImages_Click()
    ucListView1.SubItemImages = CBool(chkSubItemImages)
End Sub

Private Sub chkGridLines_Click()
    ucListView1.GridLines = CBool(chkGridLines)
End Sub

Private Sub chkFullRowSelect_Click()
    ucListView1.FullRowSelect = CBool(chkFullRowSelect)
End Sub

Private Sub chkHeaderDragDrop_Click()
    ucListView1.HeaderDragDrop = CBool(chkHeaderDragDrop)
End Sub

Private Sub chkLabelTips_Click()
    ucListView1.LabelTips = CBool(chkLabelTips)
End Sub

Private Sub chkOneClickActivate_Click()
    ucListView1.OneClickActivate = CBool(chkOneClickActivate)
End Sub

Private Sub chkTrackSelect_Click()
    ucListView1.TrackSelect = CBool(chkTrackSelect)
End Sub

Private Sub chkUnderlineHot_Click()
    ucListView1.UnderlineHot = CBool(chkUnderlineHot)
End Sub

Private Sub chkHeaderFlat_Click()
    ucListView1.HeaderFlat = CBool(chkHeaderFlat)
End Sub

Private Sub chkHeaderHide_Click()
    ucListView1.HeaderHide = CBool(chkHeaderHide)
End Sub

Private Sub chkHeaderFixedWidth_Click()
    ucListView1.HeaderFixedWidth = chkHeaderFixedWidth
End Sub

Private Sub chkScrollBarFlat_Click()
    ucListView1.ScrollBarFlat = CBool(chkScrollBarFlat)
End Sub



Private Sub ucListView1_GotFocus()
    Debug.Print ">ucListView1_GotFocus"
End Sub

Private Sub ucListView1_LostFocus()
    Debug.Print ">ucListView1_LostFocus"
End Sub

Private Sub ucListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print ">ucListView1_KeyDown"; KeyCode; Shift
End Sub

Private Sub ucListView1_KeyPress(KeyAscii As Integer)
    Debug.Print ">ucListView1_KeyPress"; KeyAscii
End Sub

Private Sub ucListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print ">ucListView1_KeyUp"; KeyCode; Shift
End Sub

Private Sub ucListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print ">ucListView1_MouseDown"; Button; Shift; x; y
End Sub

Private Sub ucListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print ">ucListView1_MouseMove"; Button; Shift; x; y
End Sub

Private Sub ucListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Debug.Print ">ucListView1_MouseUp"; Button; Shift; x; y
End Sub

Private Sub ucListView1_MouseEnter()
    Debug.Print ">ucListView1_MouseEnter"
End Sub

Private Sub ucListView1_MouseLeave()
    Debug.Print ">ucListView1_MouseLeave"
End Sub

Private Sub ucListView1_ColumnClick(Column As Integer)
    Debug.Print ">ucListView1_ColumnClick"; Column

  Dim nCol As Integer
    
    With ucListView1
        If (.Count > 1) Then
            For nCol = 0 To 2
                If (nCol <> Column) Then
                    m_ColumnSortOrder(nCol) = [soDefault] '~None
                    .ColumnIcon(nCol) = -1                '~None
                End If
            Next nCol
            If (m_ColumnSortOrder(Column) = [soAscending]) Then
                m_ColumnSortOrder(Column) = [soDescending]
                .ColumnIcon(Column) = 1 'by User
              Else
                m_ColumnSortOrder(Column) = [soAscending]
                .ColumnIcon(Column) = 0 'by User
            End If
            Select Case Column
                Case 0: Call .Sort(Column, m_ColumnSortOrder(Column), [stStringSensitive])
                Case 1: Call .Sort(Column, m_ColumnSortOrder(Column), [stNumeric])
                Case 2: Call .Sort(Column, m_ColumnSortOrder(Column), [stDate])
            End Select
        End If
        m_CurrentColumn = Column
    End With
End Sub

Private Sub ucListView1_ColumnRightClick(Column As Integer)
    Debug.Print ">ucListView1_ColumnRightClick"; Column
End Sub

Private Sub ucListView1_Click()
    Debug.Print ">ucListView1_Click"
End Sub

Private Sub ucListView1_ItemClick(Item As Integer)
    Debug.Print ">ucListView1_ItemClick"; Item
End Sub

Private Sub ucListView1_ItemCheck(Item As Integer)
    Debug.Print ">ucListView1_ItemCheck"; Item
End Sub

Private Sub ucListView1_DblClick()
    Debug.Print ">ucListView1_DblClick"
End Sub

Private Sub ucListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
'   Cancel = 1
    Debug.Print ">ucListView1_AfterLabelEdit"; Cancel; NewString
End Sub

Private Sub ucListView1_BeforeLabelEdit(Cancel As Integer)
'   Cancel = 1
    Debug.Print ">ucListView1_BeforeLabelEdit"; Cancel
End Sub

Private Sub ucListView1_Resize()
    Debug.Print ">ucListView1_Resize"
End Sub

'// New!
Private Sub ucListView1_OnSubItemPrePaint(ByVal Item As Integer, ByVal SubItem As Integer, NewBackColor As Long, NewForeColor As Long, Process As Boolean)
' Be careful what you add here!

    If (Item Mod 2) Then
        NewBackColor = RGB(150, 200, 250)
        NewForeColor = RGB(0, 0, 250)
        Process = True
    End If
    
'   or uncomment next lines for column highlighting

'    If (SubItem = m_CurrentColumn) Then
'        NewBackColor = RGB(150, 200, 250)
'        NewForeColor = RGB(0, 0, 250)
'      Else
'        NewBackColor = vbWindowBackground
'        NewForeColor = vbWindowText
'    End If
'    Process = True
End Sub

'//

Private Function pvRandomString(ByVal nChars As Integer) As String
    pvRandomString = String$(nChars, Mid$(sCHARS, Int(Rnd * Len(sCHARS)) + 1, 1))
End Function
