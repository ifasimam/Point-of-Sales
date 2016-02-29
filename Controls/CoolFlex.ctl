VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl CoolFlex 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   PropertyPages   =   "CoolFlex.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "CoolFlex.ctx":001E
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   540
      TabIndex        =   9
      Top             =   2490
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   582
      _Version        =   393216
      Format          =   94633986
      CurrentDate     =   37980
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   630
      ScaleHeight     =   825
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   105
         TabIndex        =   2
         Top             =   -30
         Width           =   2115
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   540
            TabIndex        =   4
            Top             =   450
            Width           =   45
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Please Wait ... "
            Height          =   195
            Left            =   525
            TabIndex        =   3
            Top             =   315
            Width           =   1080
         End
      End
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   540
      TabIndex        =   5
      Top             =   1770
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   510
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   1410
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   2100
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      _Version        =   393216
      Format          =   94633985
      CurrentDate     =   37985
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2325
      Left            =   495
      TabIndex        =   0
      Top             =   960
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   4101
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   -2147483642
      BackColorFixed  =   -2147483644
      BackColorSel    =   -2147483646
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BorderStyle     =   0
   End
   Begin VB.Image OptionOff 
      Height          =   210
      Left            =   105
      Picture         =   "CoolFlex.ctx":0330
      Top             =   2700
      Width           =   210
   End
   Begin VB.Image OptionOn 
      Height          =   210
      Left            =   90
      Picture         =   "CoolFlex.ctx":05DA
      Top             =   2280
      Width           =   210
   End
   Begin VB.Image imgUnchecked 
      Height          =   195
      Left            =   165
      Picture         =   "CoolFlex.ctx":0884
      Top             =   1995
      Width           =   195
   End
   Begin VB.Image imgChecked 
      Height          =   195
      Left            =   135
      Picture         =   "CoolFlex.ctx":0ACE
      Top             =   1755
      Width           =   195
   End
   Begin VB.Image imgNone 
      Height          =   330
      Left            =   105
      Top             =   105
      Width           =   330
   End
   Begin VB.Label DummyLabel 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   300
      TabIndex        =   6
      Top             =   210
      Width           =   45
   End
End
Attribute VB_Name = "CoolFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
Private NumbersOnly As Boolean
Private MyDataName As Database
Private MyRecord As Long       'var for total record
Private MyRecordPos As Long    'var for record pos
Private AutoFix As Boolean       'var for automatic fixed
Private ModifyWidth As Long
Private MyAlignment As AlignmentSettings   'var for alignment setting
Private MyEdit As Boolean                          'var for edit flexgrid
Private LoadRecord As Boolean            'var for specify wheter record is loading or not
Private ColumnType() As CoolFlexColType
Private SetColumnTypeArray As Boolean
Private LastCol As Long
Private FrstCol As Long
Private ComboBoxCount As Integer
Private mLaunchForm As String
Private SortOnHeader As Boolean
Private SortOnHeaderValue As CoolFlexSort
Private flexKeyDown As Integer         '--igit value for KeyDown
Private zTxtValue As String               '-- igit 02/11/02
Private txtBoxValue As String             '-- igit 03/11/04
Private zValue As Boolean
Private zKeyPress As Integer
Private zFooter As Integer    '--igit 02/19/02
Private ColumnLock() As Boolean  '---igit 04/23/04
'component activity
Public Event Click()
Public Event EnterCell(Rowsel As Long, Colsel As Long, Value As String)
Public Event DblClick()
Public Event LeaveCell()
Public Event RowColChange()
Public Event ExitFocus()
Public Event CellComboBoxClick(ColIndex As Long, Value As String)
Public Event CellComboBoxChange(ColIndex As Long)

'--- igit -------------------------------------------
Public Event CellComboBoxKeyPress(ColIndex As Long, KeyAscii As Integer)
Public Event CellComboBoxValidate(ColIndex As Long, Cancel As Boolean, Value As String)
Public Event CellComboBoxLostFocus(ColIndex As Long)
Public Event CellComboGotFocus(ColIndex As Long)

Public Event CellTextBoxKeyPress(KeyAscii As Integer)
Public Event CellTextBoxKeyDown(KeyCode As Integer, Shift As Integer)
Public Event CellTextBoxValidate(Cancel As Boolean)
Public Event CellTextBoxLostFocus()
Public Event CellTextBoxChange(Value As Variant)   '---igit 05/29/2002
Public Event CellDtPickerClick()
Public Event DtPickerCallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
Public Event DtPickerChange()
Public Event DtPickerCloseUp()
Public Event CellTmPickerClick()
Public Event CellTmPickerChange()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event DTPickerKeyDown(KeyCode As Integer, Shift As Integer)
Public Event TmPickerKeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event EnterTextBox(Value As Variant)
'---------------------------------------------------
Public Event CellCheckBoxClick(ColIndex As Long, Value As Integer)

Public Enum CoolFlexGridLines
    GridFlat = 1
    GridInset = 2
    GridNone = 0
    GridRaised = 3
End Enum

Public Enum CoolFlexScrollBar
    ScrollBarBoth = 3
    ScrollBarHorizontal = 1
    ScrollBarNone = 0
    ScrollBarVertical = 2
End Enum

Public Enum CoolFlexSort
    SortNone = 0
    SortGenericAscending = 1
    SortGenericDescending = 2
    SortNumericAscending = 3
    SortNumericDescending = 4
    SortStringNoCaseAsending = 5
    SortNoCaseDescending = 6
    SortStringAscending = 7
    SortStringDescending = 8
End Enum

Public Enum CoolFlexColType
    eNone      '--igit
    eTextbox
    eCheckbox
    eCombobox
    eDTPicker  '--- igit
    eTmPicker  '---igit
    eOption
End Enum
'---- 12/08/01 by igit
Public Enum txtBorderStyle
   None = 0
   FixedSingle = 1
End Enum

Public Enum MergeSettings
   flexMergeFree = 0
   flexMergeNever = 1
   flexMergeRestrictAll = 2
   flexMergeRestrictColumns = 3
   flexMergeRestrictRows = 4
End Enum

Public Enum SetAlignment
   flexAlignLeftTop = 0       'The cell content is aligned left, top.
   flexAlignLeftCenter = 1    'Default for strings. The cell content is aligned left, center.
   flexAlignLeftBottom = 2    'The cell content is aligned left, bottom.
   flexAlignCenterTop = 3     'The cell content is aligned center, top.
   flexAlignCenterCenter = 4  'The cell content is aligned center, center.
   flexAlignCenterBottom = 5  'The cell content is aligned center, bottom.
   flexAlignRightTop = 6      'The cell content is aligned right, top.
   flexAlignRightCenter = 7   'Default for numbers. The cell content is aligned right, center.
   flexAlignRightBottom = 8   'The cell content is aligned right, bottom.
   flexAlignGeneral = 9       'The cell content is of general alignment. This is "left, center" for strings and "right, center" for numbers.
End Enum

Public Enum GrdSelection
   flexSelectionFree = 0      'Free. This allows individual cells in the MSHFlexGrid to be selected, spreadsheet style. This is the default.
   flexSelectionByRow = 1     'By Row. This forces selections to span entire rows, as in a multi-column list box or record-based display.
   flexSelectionByColumn = 2  'By Column. This forces selections to span entire columns, as if selecting ranges for a chart or fields for sorting.
End Enum

Public Property Let Footer(ByVal zValue As Integer)
   zFooter = zValue
End Property

Public Property Get Footer() As Integer
   Footer = zFooter
End Property

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
   '---msgbox "Hi"
   frmAbout.show
End Sub

Private Sub Combo1_Change(Index As Integer)
    RaiseEvent CellComboBoxChange(MSFlexGrid1.col)
End Sub

Private Sub Combo1_Click(Index As Integer)
'    RaiseEvent CellComboBoxClick(MSFlexGrid1.Col, MSFlexGrid1.Text)
   RaiseEvent CellComboBoxClick(MSFlexGrid1.col, Combo1(Index).Text)
'    If Combo1(LastCol).Visible = True Then
'       MSFlexGrid1.Text = Combo1(LastCol).Text
'       Combo1(LastCol).Text = ""
'       Combo1(LastCol).Visible = False
'    End If
    
End Sub

'--- igit 02/04/02
Private Sub Combo1_GotFocus(Index As Integer)
   If AutoDropDown Then
      Call SendMessage(Combo1(MSFlexGrid1.col).hwnd, CB_SHOWDROPDOWN, True, 0)
      'SendKeys "%{DOWN}"
   End If
   RaiseEvent CellComboGotFocus(MSFlexGrid1.col)
End Sub
'-------------------------------------
Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   'RaiseEvent CellComboBoxKeyPress(MSFlexGrid1.Col, KeyAscii)
   KeyAscii = AutoFind(Combo1(MSFlexGrid1.col), KeyAscii, False)
End Sub

Private Function AutoFind(ByRef cboCurrent As ComboBox, _
   ByVal KeyAscii As Integer, Optional ByVal LimitToList As Boolean = False)
   
   Dim ICB As Long
   Dim sFindString As String
   
   On Error GoTo ErrTrap
   'If KeyAscii = 8 Then '
      'If cboCurrent.SelStart <= 1 Then
         'cboCurrent = ""
         'KeyAscii = KeyAscii
     '    AutoFind = 0
        ' Exit Function
      'End If
      'If cboCurrent.SelLength = 0 Then
      '   sFindString = UCase(Left(cboCurrent, Len(cboCurrent) - 1))
      'Else
      '   sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart - 1)
      'End If
   'Else
   If KeyAscii <> 8 And KeyAscii < 32 Or KeyAscii > 127 Then
      Exit Function
   Else
      If cboCurrent.SelLength = 0 Then
         sFindString = UCase(cboCurrent.Text & Chr$(KeyAscii))
      Else
         sFindString = Left$(cboCurrent.Text, cboCurrent.SelStart) & Chr$(KeyAscii)
      End If
      ICB = SendMessage(cboCurrent.hwnd, CB_FINDSTRING, -1, ByVal sFindString)
      If ICB <> CB_ERR Then
         cboCurrent.ListIndex = ICB
         cboCurrent.SelStart = Len(sFindString)
         cboCurrent.SelLength = Len(cboCurrent.Text) - cboCurrent.SelStart
         AutoFind = 0
      Else
         If LimitToList = True Then
            AutoFind = 0
         Else
            AutoFind = KeyAscii
         End If
      End If
   End If
exitFunc:
   Exit Function
   
ErrTrap:
   MsgBox err.Number & "-" & err.Description
   err.Clear
   Resume exitFunc
End Function
'--igit-----------------------
Private Sub Combo1_LostFocus(Index As Integer)
   RaiseEvent CellComboBoxLostFocus(MSFlexGrid1.col)
   Combo1(Index).Visible = False
End Sub
   
Private Sub Combo1_Validate(Index As Integer, Cancel As Boolean)
'   If AutoDropDown Then SendKeys "%{UP}"
   RaiseEvent CellComboBoxValidate(MSFlexGrid1.col, Cancel, Combo1(Index).Text)
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
   RaiseEvent DtPickerCallbackKeyDown(KeyCode, Shift, CallbackField, CallbackDate)
End Sub

Private Sub DTPicker1_Change()
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = DTPicker1.Value
   RaiseEvent DtPickerChange
End Sub

Private Sub DTPicker1_Click()
   RaiseEvent CellDtPickerClick
End Sub

Private Sub DTPicker1_CloseUp()
   RaiseEvent DtPickerCloseUp
End Sub

'------------------------------
Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
   '--igit
   Select Case KeyCode
        Case vbKeyEscape        'user press escape key
           DTPicker1.Visible = False
        
        Case vbKeyReturn
            If DTPicker1.Visible = True Then
               MSFlexGrid1.Text = DTPicker1.Value
               DTPicker1.Visible = False
            End If
        
        Case vbKeyDown          'user press arrow down key
'           MSFlexGrid1.SetFocus
'           DoEvents
'           If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'              MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'           End If
        
        Case vbKeyUp            'user press arrow up key
'           MSFlexGrid1.SetFocus
'           DoEvents
'           If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
'              MSFlexGrid1.Row = MSFlexGrid1.Row - 1
'            End If
    
        Case vbKeyLeft
'            If Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = 0 Then
'                MSFlexGrid1.Col = MSFlexGrid1.Col - 1
'            ElseIf Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = Len(Combo1(Index).Text) Then
'                Combo1(Index).SelStart = 0
'            End If
            
        Case vbKeyRight
'            If Combo1(Index).SelStart = Len(Combo1(Index).Text) Then
'                MSFlexGrid1.Col = MSFlexGrid1.Col + 1
'            ElseIf Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = Len(Combo1(Index).Text) Then
'                Combo1(Index).SelStart = Len(Combo1(Index).Text)
'            End If
            
    End Select
    RaiseEvent DTPickerKeyDown(KeyCode, Shift)
End Sub

Private Sub DTPicker1_LostFocus()
   DTPicker1.Visible = False
End Sub

Private Sub DTPicker2_Change()
   MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, MSFlexGrid1.col) = Format(DTPicker2.Value, "hh:mm ampm")
   RaiseEvent CellTmPickerChange
End Sub

Private Sub DTPicker2_Click()
   RaiseEvent CellTmPickerClick
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)
     '--igit
   Select Case KeyCode
        Case vbKeyEscape        'user press escape key
           DTPicker2.Visible = False
        
        Case vbKeyReturn
            If DTPicker2.Visible = True Then
               MSFlexGrid1.Text = Format(DTPicker2.Value, "hh:mm ampm")
               DTPicker2.Visible = False
            End If
        
        Case vbKeyDown          'user press arrow down key
'           MSFlexGrid1.SetFocus
'           DoEvents
'           If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'              MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'           End If
        
        Case vbKeyUp            'user press arrow up key
'           MSFlexGrid1.SetFocus
'           DoEvents
'           If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
'              MSFlexGrid1.Row = MSFlexGrid1.Row - 1
'            End If
    
        Case vbKeyLeft
'            If Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = 0 Then
'                MSFlexGrid1.Col = MSFlexGrid1.Col - 1
'            ElseIf Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = Len(Combo1(Index).Text) Then
'                Combo1(Index).SelStart = 0
'            End If
            
        Case vbKeyRight
'            If Combo1(Index).SelStart = Len(Combo1(Index).Text) Then
'                MSFlexGrid1.Col = MSFlexGrid1.Col + 1
'            ElseIf Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = Len(Combo1(Index).Text) Then
'                Combo1(Index).SelStart = Len(Combo1(Index).Text)
'            End If
            
    End Select
    RaiseEvent TmPickerKeyDown(KeyCode, Shift)
End Sub

Private Sub DTPicker2_LostFocus()
   DTPicker2.Visible = False
End Sub

Private Sub MSFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
    
      With MSFlexGrid1
        If (.col + 1) <= (.Cols - 1) Then
          .col = .col + 1
        Else
          If (.Row + 1) <= (.Rows - 1) Then
            .Row = .Row + 1
            .col = 1
          End If
        End If
            
      End With
    End If
    
    'MSFlexGrid1_DblClick
    
   flexKeyDown = KeyCode
   RaiseEvent KeyDown(KeyCode, Shift)  '--igit
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
   '--igit
   RaiseEvent KeyPress(KeyAscii)
   If Not ColumnLock(MSFlexGrid1.col) Then
      If ColumnType(Colsel) = eCombobox Then
         KeyAscii = AutoFind(Combo1(MSFlexGrid1.col), KeyAscii, False)
      End If
      
      GridColEdit MSFlexGrid1, KeyAscii
   End If
   KeyAscii = 0
End Sub

Private Sub GridColEdit(ctrlGrid As CONTROL, ziKeyPress As Integer)
   'On Error Resume Next
   
   Dim Rowsel As Long
   Dim Colsel As Long
   Dim Value As String
   Dim SelStart As Long
   
   LastCol = MSFlexGrid1.Colsel
   Rowsel = MSFlexGrid1.Rowsel
   Colsel = MSFlexGrid1.Colsel
  
   Select Case ziKeyPress
      Case 0 To 32 'spacebar
            Value = MSFlexGrid1.TextMatrix(MSFlexGrid1.Rowsel, MSFlexGrid1.Colsel)
            SelStart = 0
      Case Else
         If ColumnType(Colsel) = eDTPicker Then Exit Sub
         If ColumnType(Colsel) = eTmPicker Then Exit Sub
         SelStart = 1
         Value = Chr(ziKeyPress)
      End Select
      
   If ctrlGrid.Row < ctrlGrid.Rows - Me.Footer Then   '---last row

'      If MSFlexGrid1.MouseRow = 0 And SortOnHeader = True Then
'         RemoveItem (MSFlexGrid1.Rows)
'         MSFlexGrid1.Sort = SortOnHeaderValue
'         AddItem ""
'      End If

      If MyEdit = True And LoadRecord = False And MSFlexGrid1.MouseCol > 0 And MSFlexGrid1.MouseRow > 0 Then
         'If MSFlexGrid1.MouseCol > 0 And MSFlexGrid1.MouseRow > 0 Then
         Select Case ColumnType(Colsel)
            Case eTextbox 'default
                    zTxtValue = MSFlexGrid1.TextMatrix(MSFlexGrid1.Rowsel, MSFlexGrid1.Colsel)
                    Text1.BackColor = MSFlexGrid1.BackColor
                    Text1.ForeColor = MSFlexGrid1.ForeColor
                    Set Text1.Font = MSFlexGrid1.Font
                    Text1.Width = IIf(MSFlexGrid1.CellWidth < 0, 0, MSFlexGrid1.CellWidth)
                    Text1.Height = MSFlexGrid1.CellHeight
                    Text1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                    Text1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                    Text1.Text = Value
                    
                    If MSFlexGrid1.CellWidth > 0 Then
                    Text1.Visible = True
                    Text1.SetFocus
                    Text1.SelStart = SelStart
                    Text1.SelLength = Len(Text1.Text)
                    End If
               ' End If
                     RaiseEvent EnterTextBox(MSFlexGrid1.TextMatrix(MSFlexGrid1.Rowsel, MSFlexGrid1.Colsel))
            Case eCheckbox
                'If MSFlexGrid1.MouseCol > 0 And MSFlexGrid1.MouseRow > 0 Then
                    MSFlexGrid1.CellPictureAlignment = 4 'center x center
                    If MSFlexGrid1.CellPicture = imgChecked.Picture Then  'MSFlexGrid1.Text = "C" Then
                        Set MSFlexGrid1.CellPicture = imgUnchecked.Picture  'LoadPicture(App.Path & "\Checked.bmp")
                        'MSFlexGrid1.Text = "U"
                        RaiseEvent CellCheckBoxClick(MSFlexGrid1.col, 0)
                    Else
                        Set MSFlexGrid1.CellPicture = imgChecked.Picture  'LoadPicture(App.Path & "\Checked.bmp")
                        'MSFlexGrid1.Text = "C"
                        RaiseEvent CellCheckBoxClick(MSFlexGrid1.col, 1)
                    End If
                    '--- this routine is just for checking the true value
'                    If MSFlexGrid1.CellPicture = imgUnchecked.Picture Then
'                        MsgBox "Unchecked"
'                    ElseIf MSFlexGrid1.CellPicture = imgChecked.Picture Then
'                        MsgBox "Check"
'                    End If
                    '-------------------------------------------------
                'End If
            Case eOption
                  MSFlexGrid1.CellPictureAlignment = 4 'center x center
                  SetOptionButton
                  If MSFlexGrid1.CellPicture = OptionOn.Picture Then   'MSFlexGrid1.Text = "C" Then
                      Set MSFlexGrid1.CellPicture = OptionOff.Picture   'LoadPicture(App.Path & "\Checked.bmp")
                      'MSFlexGrid1.Text = "U"
                      'RaiseEvent CellCheckBoxClick(MSFlexGrid1.Col, 0)
                  Else
                      Set MSFlexGrid1.CellPicture = OptionOn.Picture   'LoadPicture(App.Path & "\Checked.bmp")
                      'MSFlexGrid1.Text = "C"
                     ' RaiseEvent CellCheckBoxClick(MSFlexGrid1.Col, 1)
                  End If
            
            
            
            
            Case eCombobox
                'If MSFlexGrid1.MouseCol > 0 And MSFlexGrid1.MouseRow > 0 Then
                    Combo1(Colsel).BackColor = MSFlexGrid1.BackColor
                    Combo1(Colsel).ForeColor = MSFlexGrid1.ForeColor
                    Set Combo1(Colsel).Font = MSFlexGrid1.Font
                    Combo1(Colsel).Width = MSFlexGrid1.CellWidth
                    Combo1(Colsel).Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                    Combo1(Colsel).Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                    
                    If Combo1(Colsel).Style = 2 Then
                      cboFindList Combo1(Colsel), Value
                    Else
                      Combo1(Colsel).Text = Value
                    End If
                    Combo1(Colsel).Visible = True
                    Combo1(Colsel).ZOrder
                    Combo1(Colsel).SetFocus
                    If Combo1(Colsel).Style < 2 Then
                    Combo1(Colsel).SelStart = SelStart
                    Combo1(Colsel).SelLength = Len(Combo1(Colsel).Text)
                    End If
                'End If
            Case eDTPicker
                'If MSFlexGrid1.MouseCol > 0 And MSFlexGrid1.MouseRow > 0 Then
                    On Error Resume Next
                    DTPicker1.Width = MSFlexGrid1.CellWidth
                    DTPicker1.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                    DTPicker1.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                    If Value <> "" Then DTPicker1.Value = Value
                    DTPicker1.Visible = True
                    DTPicker1.ZOrder
                    DTPicker1.SetFocus
                'End If
               Case eTmPicker
                'If MSFlexGrid1.MouseCol > 0 And MSFlexGrid1.MouseRow > 0 Then
                    On Error Resume Next
                    DTPicker2.Width = MSFlexGrid1.CellWidth
                    DTPicker2.Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
                    DTPicker2.Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
                    If Value <> "" Then DTPicker2.Value = Value
                    DTPicker2.Visible = True
                    DTPicker2.ZOrder
                    DTPicker2.SetFocus
                'End If
         End Select
      End If
   End If
End Sub

Private Sub MSFlexGrid1_RowColChange()
    RaiseEvent RowColChange
End Sub

Private Sub MSFlexGrid1_Click()
   RaiseEvent Click
   With MSFlexGrid1
      FrstCol = .col
      'If ColumnType(.col) = eCombobox And .MouseRow <> 0 Then MSFlexGrid1_KeyPress (32)
      If ColumnType(.col) = eCheckbox And .MouseRow <> 0 Then MSFlexGrid1_KeyPress (32)
      'If ColumnType(.col) = eDTPicker And .MouseRow <> 0 Then MSFlexGrid1_KeyPress (32)
      If ColumnType(.col) = eTmPicker And .MouseRow <> 0 Then MSFlexGrid1_KeyPress (32)
      If ColumnType(.col) = eOption And .MouseRow <> 0 Then MSFlexGrid1_KeyPress (32)
      
   End With
End Sub

Private Sub MSFlexGrid1_DblClick()
   RaiseEvent DblClick
   If ColumnType(MSFlexGrid1.col) = eDTPicker And MSFlexGrid1.MouseRow <> 0 Then MSFlexGrid1_KeyPress (32)
   If ColumnType(MSFlexGrid1.col) = eCombobox And MSFlexGrid1.MouseRow <> 0 Then MSFlexGrid1_KeyPress (32)
   If MSFlexGrid1.MouseCol <> 0 And MSFlexGrid1.MouseRow <> 0 Then
      MSFlexGrid1_KeyPress (32)  '-- staff Keyascii with spacebar
   End If
End Sub

'end of component activity

'component methods
Public Sub Clear()
    Dim X As Long
    MSFlexGrid1.Clear
    
    Text1.Text = ""
    Text1.Visible = False
    
    For X = 0 To Combo1.UBound
        Combo1(X).Text = ""
        Combo1(X).Visible = False
    Next
    
    DTPicker1.Visible = False '--igit
    DTPicker2.Visible = False '--igit
    SetCheckBoxes
    SetOptionButton
End Sub

Public Sub RemoveItem(ByVal Index As Long)
    MSFlexGrid1.RemoveItem Index
    Text1.Text = ""
    Text1.Visible = False
    Combo1(LastCol).Visible = False
    'Combo1(LastCol).Text = ""
       
    DTPicker1.Visible = False    '--igit
    DTPicker2.Visible = False '--igit
End Sub

Public Sub AddItem(Item As String, Optional Index As Integer)

    'MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
    If Index = 0 Then
      MSFlexGrid1.AddItem Item
    Else
      MSFlexGrid1.AddItem Item, Index
    End If
    Text1.Text = ""
    Text1.Visible = False
    Combo1(LastCol).Visible = False
    'Combo1(LastCol).Text = ""
        
    DTPicker1.Visible = False    '--igit
    DTPicker2.Visible = False '--igit
    SetCheckBoxes
    SetOptionButton
End Sub

Public Sub ColType(ByVal ColNumber As Long, ByVal eType As CoolFlexColType)
    ColumnType(ColNumber) = eType
    Select Case eType
        Case eNone 'default
        
        Case eTextbox
        
        Case eCheckbox
            SetCheckBoxes
        
        Case eCombobox
        
        Case eDTPicker  '--igit
        Case eTmPicker
        Case eOption '-04/29/04
            SetOptionButton
    End Select
End Sub

Public Sub ComboBoxAddItem(ByVal col As Long, ByVal Item As String)
    Combo1(col).AddItem Item
End Sub

'Public Sub ComboBoxStyle(ByVal col As Integer, ByVal sty As Integer)
'  Combo1(col).Style = sty
'End Sub

Public Sub ComboBoxClear(ByVal col As Long)
    Combo1(col).Clear
End Sub

Public Sub allownumbersonly(Optional ByVal b As Boolean = False)
  NumbersOnly = b
End Sub

Public Sub ComboBoxRemoveItem(ByVal col As Long, ByVal Index As Integer)
    Combo1(col).RemoveItem Index
End Sub

Public Sub SortOnHeaderClick(ByVal SortOn As Boolean, ByVal newValue As CoolFlexSort)
    SortOnHeader = SortOn
    SortOnHeaderValue = newValue
End Sub

'end component methods

Private Sub MSFlexGrid1_EnterCell()
    Dim Rowsel As Long
    Dim Colsel As Long
    Dim Value As String
    zTxtValue = ""
    LastCol = MSFlexGrid1.Colsel
    Rowsel = MSFlexGrid1.Rowsel
    Colsel = MSFlexGrid1.Colsel
    RaiseEvent EnterCell(Rowsel, Colsel, Value)
End Sub

Private Sub MSFlexGrid1_LeaveCell()
    
    If Text1.Visible = True Then
       MSFlexGrid1.Text = Text1.Text
       zTxtValue = MSFlexGrid1.Text
       Text1.Text = ""
       Text1.Visible = False
    End If
    
    If Combo1(LastCol).Visible = True Then
       MSFlexGrid1.Text = Combo1(LastCol).Text
       'Combo1(LastCol).Text = ""
       Combo1(LastCol).Visible = False
    End If
    
    If DTPicker1.Visible = True Then
       MSFlexGrid1.Text = DTPicker1.Value
       DTPicker1.Visible = False
    End If
    
    If DTPicker2.Visible = True Then
       MSFlexGrid1.Text = Format(DTPicker2.Value, "hh:mm ampm")
       DTPicker2.Visible = False
    End If
    
    RaiseEvent LeaveCell
End Sub

Private Sub MSFlexGrid1_Scroll()
    Text1.Visible = False
    Combo1(LastCol).Visible = False
    
    DTPicker1.Visible = False    '--igit
    DTPicker2.Visible = False    '--igit
End Sub

Private Sub Combo1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape        'user press escape key
           Combo1(Index).Visible = False
        
        Case vbKeyReturn
            If Combo1(LastCol).Visible = True Then
               MSFlexGrid1.Text = Combo1(LastCol).Text
               'Combo1(LastCol).Text = ""
               Combo1(LastCol).Visible = False
            End If
        Case vbKeyDown          'user press arrow down key
'           MSFlexGrid1.SetFocus
'           DoEvents
'           If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
'              MSFlexGrid1.Row = MSFlexGrid1.Row + 1
'           End If

        Case vbKeyUp            'user press arrow up key
'           MSFlexGrid1.SetFocus
'           DoEvents
'           If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
'              MSFlexGrid1.Row = MSFlexGrid1.Row - 1
'            End If

        Case vbKeyLeft
'            If Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = 0 Then
'                MSFlexGrid1.Col = MSFlexGrid1.Col - 1
'            ElseIf Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = Len(Combo1(Index).Text) Then
'                Combo1(Index).SelStart = 0
'            End If

        Case vbKeyRight
'            If Combo1(Index).SelStart = Len(Combo1(Index).Text) Then
'                MSFlexGrid1.Col = MSFlexGrid1.Col + 1
'            ElseIf Combo1(Index).SelStart = 0 And Len(Combo1(Index).SelText) = Len(Combo1(Index).Text) Then
'                Combo1(Index).SelStart = Len(Combo1(Index).Text)
'            End If

    End Select
End Sub

Private Sub MSFlexGrid1_SelChange()
   Text1.Visible = False
End Sub

'---igit 05/29/2002
Private Sub Text1_Change()
   'If zTxtValue = "" Then
   'zTxtValue = Text1.Text
   'End If
   Dim cValue As String
   cValue = Text1.Text
   txtBoxValue = Text1.Text
   RaiseEvent CellTextBoxChange(cValue)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bcancel As Boolean
    RaiseEvent CellTextBoxKeyDown(KeyCode, Shift)
    zKeyPress = KeyCode
    Select Case KeyCode
        Case vbKeyEscape        'user press escape key
           Text1.Visible = False
        
        Case vbKeyReturn
          ' zTxtValue = Text1.Text
            '-----igit 04/23/04
            RaiseEvent CellTextBoxValidate(bcancel)
            If Not bcancel Then MSFlexGrid1_LeaveCell
        
        Case vbKeyDown          'user press arrow down key
          ' zTxtValue = Text1.Text
           '-----igit 04/23/04
            RaiseEvent CellTextBoxValidate(bcancel)
            If bcancel Then Exit Sub
            
           
           MSFlexGrid1_LeaveCell
           MSFlexGrid1.SetFocus
           DoEvents
           If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
              MSFlexGrid1.Row = MSFlexGrid1.Row + 1
           End If
        
        Case vbKeyUp            'user press arrow up key
          ' zTxtValue = Text1.Text
          '-----igit 04/23/04
            RaiseEvent CellTextBoxValidate(bcancel)
            If bcancel Then Exit Sub
            
           MSFlexGrid1_LeaveCell
           MSFlexGrid1.SetFocus
           DoEvents
           If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
              MSFlexGrid1.Row = MSFlexGrid1.Row - 1
            End If
    
        Case vbKeyLeft
            '-----igit 04/23/04
            RaiseEvent CellTextBoxValidate(bcancel)
            If bcancel Then Exit Sub
            
            If Text1.SelStart = 0 And Len(Text1.SelText) = 0 Then
                MSFlexGrid1.col = MSFlexGrid1.col - 1
            ElseIf Text1.SelStart = 0 And Len(Text1.SelText) = Len(Text1.Text) Then
                Text1.SelStart = 0
            End If
            
        Case vbKeyRight
            '-----igit 04/23/04
            RaiseEvent CellTextBoxValidate(bcancel)
            If bcancel Then Exit Sub
            
            If Text1.SelStart = Len(Text1.Text) Then
                If MSFlexGrid1.col <> MSFlexGrid1.Cols - 1 Then MSFlexGrid1.col = MSFlexGrid1.col + 1
                If MSFlexGrid1.col = MSFlexGrid1.Cols - 1 Then MSFlexGrid1.col = MSFlexGrid1.col
            ElseIf Text1.SelStart = 0 And Len(Text1.SelText) = Len(Text1.Text) Then
                Text1.SelStart = Len(Text1.Text)
            End If
            
    End Select
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   RaiseEvent CellTextBoxKeyPress(KeyAscii)
   
   If NumbersOnly Then
   If ((KeyAscii <> 8) And (KeyAscii <> vbKeyDelete) And _
  (KeyAscii <> 46)) And ((KeyAscii < 48 Or KeyAscii > 57)) Then
    KeyAscii = 0
  Else
    If KeyAscii = 46 Then
      If InStr(Text1, ".") Then
        KeyAscii = 0
        Exit Sub
      End If
    End If
    KeyAscii = KeyAscii
  End If
  End If
End Sub

Private Sub Text1_LostFocus()
   Text1.Visible = False
   RaiseEvent CellTextBoxLostFocus
   
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   RaiseEvent CellTextBoxValidate(Cancel)
   If Cancel = False Then
      Select Case zKeyPress
      Case 13
      Case 38
      Case 40
      Case Else
      End Select
      If zTxtValue = "" Then
         zTxtValue = Text1.Text
      End If
      MSFlexGrid1_LeaveCell
   End If
   
End Sub

Private Sub UserControl_EnterFocus()
   On Error GoTo ErrTrap
  
   MSFlexGrid1.SetFocus
   'GridColEdit MSFlexGrid1, 32
   Exit Sub
ErrTrap:
   MsgBox err.Number & " - " & err.Description
   Resume Next
End Sub

Private Sub UserControl_ExitFocus()
    '---igit
    On Error GoTo ErrTrap
    MSFlexGrid1_LeaveCell
    Text1.Text = ""
    Text1.Visible = False
    'Combo1(LastCol).Text = ""
    Combo1(LastCol).Visible = False
    DTPicker1.Visible = False '--igit
    DTPicker2.Visible = False
    RaiseEvent ExitFocus  '--igit
    Exit Sub
    
ErrTrap:
   MsgBox err.Number & " - " & err.Description
   Resume Next
End Sub

Private Sub UserControl_Initialize()
   On Error GoTo ErrTrap
    'initialize control in design time
    MSFlexGrid1.Top = 0
    MSFlexGrid1.Left = 0
    MSFlexGrid1.Width = UserControl.Width - 60
    MSFlexGrid1.Height = UserControl.Height - 60
    'MSFlexGrid1.RowHeight(1) = 315
    'coordinate progress
    Picture1.Left = (UserControl.Width / 2) - (Picture1.Width / 2)
    Picture1.Top = (UserControl.Height / 2) - (Picture1.Height / 2)
    DTPicker1.Value = Date
    DTPicker2.Value = Time
    'zFooter = 1
    Exit Sub
ErrTrap:
   MsgBox err.Number & " - " & err.Description
   Resume Next
End Sub

'------------------------------------------------------
Private Sub UserControl_Resize()
    On Error Resume Next 'GoTo ErrTrap
    MSFlexGrid1.Top = 0
    MSFlexGrid1.Left = 0
    MSFlexGrid1.Width = UserControl.Width - 60
    MSFlexGrid1.Height = UserControl.Height - 60
    Picture1.Left = (UserControl.Width / 2) - (Picture1.Width / 2)
    Picture1.Top = (UserControl.Height / 2) - (Picture1.Height / 2)
'    Exit Sub
'ErrTrap:
'    'MsgBox Err.Number & " - " & Err.Description
'    Resume Next
End Sub

Public Sub Show_Record(ByVal SQLCommand As String)
    'On Error GoTo errorhandler
    LoadRecord = True
    Dim Maindb As Database
    Dim theset As Object
    Dim c As Long, No As Long
    Dim DynamicCol() As Long
    Dim TotalColoumn As Long
    Dim MyData As String
    Dim DataWidth As Long
    
    'open recordset
    Set theset = MyDataName.OpenRecordset(SQLCommand)
    
    If theset.EOF Then Exit Sub  'if no record exist
    'calculate total field
    TotalColoumn = theset.Fields.Count
    
    Set_Grid (TotalColoumn)
    'recreate array in run time
    
    For c = 1 To theset.Fields.Count
       MSFlexGrid1.TextMatrix(0, c) = theset.Fields(c - 1).Name
    Next c
    
    theset.MoveLast
    MyRecord = theset.AbsolutePosition + 1
    theset.MoveFirst
    
    If AutoFixCol = False Then
    Do While Not theset.EOF
       DoEvents
       No = No + 1
       MyRecordPos = theset.AbsolutePosition + 1
       Label2.Caption = Format(MyRecordPos / MyRecord * 100, "##") & "  % Completed"
       MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
       MSFlexGrid1.TextMatrix(No, 0) = Str(No)
       For c = 1 To theset.Fields.Count
          MSFlexGrid1.col = c
          MSFlexGrid1.Row = No
          MSFlexGrid1.CellAlignment = MyAlignment
          MSFlexGrid1.TextMatrix(No, c) = theset.Fields(c - 1).Value & ""
       Next c
       theset.MoveNext
    Loop
    'when select autofixcol=true
    Else
    ReDim DynamicCol(TotalColoumn)
       Do While Not theset.EOF
          DoEvents
          No = No + 1
          MyRecordPos = theset.AbsolutePosition + 1
          Label2.Caption = Format(MyRecordPos / MyRecord * 100, "##") & "  % Completed"
          MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
          MSFlexGrid1.TextMatrix(No, 0) = Str(No)
          For c = 1 To theset.Fields.Count
          MSFlexGrid1.col = c
          MSFlexGrid1.Row = No
          MSFlexGrid1.CellAlignment = MyAlignment
             MyData = Trim(theset.Fields(c - 1).Value) & ""
             MSFlexGrid1.TextMatrix(No, c) = MyData
             'get the width value
             DummyLabel.Caption = MyData
             DataWidth = DummyLabel.Width
             If DynamicCol(c) < DataWidth + ModifyWidth + 100 Then
                DynamicCol(c) = DataWidth + ModifyWidth + 100
             End If
          
          MSFlexGrid1.ColWidth(c) = DynamicCol(c)
          Next c
          theset.MoveNext
       Loop
    End If
    
    Set theset = Nothing
    LoadRecord = False
    Exit Sub
errorhandler:
    
    MsgBox err.Number & "  " & err.Description
    
End Sub

Private Sub Set_Grid(ByVal mycol As Long)
    'setting msflexgrid control
    MSFlexGrid1.Clear
    MSFlexGrid1.Cols = mycol + 1
    MSFlexGrid1.Rows = 1
    MSFlexGrid1.TextMatrix(0, 0) = "No."
    MSFlexGrid1.ColWidth(0) = 500
End Sub

'property
'text matrix property
Public Property Get TextMatrix(ByVal Row As Long, ByVal col As Long) As String
    TextMatrix = MSFlexGrid1.TextMatrix(Row, col)
End Property

Public Property Let TextMatrix(ByVal Row As Long, ByVal col As Long, ByVal NewText As String)
    MSFlexGrid1.TextMatrix(Row, col) = NewText
End Property

'----------Set Col & Row  creat comboBox array
Public Property Let Cols(ByVal NewCols As Long)
    On Error Resume Next
    
    Dim X As Integer
    
    MSFlexGrid1.Cols = NewCols
    ReDim ColumnType(NewCols - 1)
    ReDim ColumnLock(NewCols - 1) '---igit 04/23/04
    SetColumnTypeArray = True
    
    For X = 1 To NewCols - 1
        Load Combo1(X)
    Next
    PropertyChanged "Cols"
End Property

Public Property Get Cols() As Long
Attribute Cols.VB_ProcData.VB_Invoke_Property = "General"
    Cols = MSFlexGrid1.Cols
End Property

Public Property Let col(ByVal NewCol As Long)
    MSFlexGrid1.col = NewCol
End Property

Public Property Get col() As Long
    col = MSFlexGrid1.col
End Property

Public Property Let Rows(ByVal NewRows As Long)
    MSFlexGrid1.Rows = NewRows
    PropertyChanged "Rows"
    SetCheckBoxes
    SetOptionButton
End Property

Public Property Get Rows() As Long
Attribute Rows.VB_ProcData.VB_Invoke_Property = "General"
    Rows = MSFlexGrid1.Rows
End Property

Public Property Let Row(ByVal NewRow As Long)
  On Error GoTo errhandler
    MSFlexGrid1.Row = NewRow
    Exit Property
errhandler:
    'MsgBox "Error: " & err.Description & vbCr _
    & "User Control: Coolflex" & vbCr _
    & "Property: Row"
    err.Clear
End Property

Public Property Get Row() As Long
    Row = MSFlexGrid1.Row
End Property

Private Sub SetCheckBoxes(Optional ByVal iCol As Integer, Optional ByVal iRow As Integer)
    Dim X As Long
    Dim Y As Long
    Dim TempRow As Long
    Dim TempCol As Long
            
    If SetColumnTypeArray = False Then
        Exit Sub
    End If
    
    TempRow = MSFlexGrid1.Row
    TempCol = MSFlexGrid1.col
    
    For X = 1 To MSFlexGrid1.Rows - 1
        For Y = 0 To MSFlexGrid1.Cols - 1
            If ColumnType(Y) = eCheckbox Then
                If Len(MSFlexGrid1.TextMatrix(X, Y)) = 0 Then
                    MSFlexGrid1.Row = X
                    MSFlexGrid1.col = Y
                    MSFlexGrid1.CellPictureAlignment = 4 'center x center
                    Set MSFlexGrid1.CellPicture = imgUnchecked.Picture
                    MSFlexGrid1.Text = " "
                    'MSFlexGrid1.CellForeColor = vbWhite
                End If
            End If
        Next Y
    Next X
    MSFlexGrid1.Row = TempRow
    MSFlexGrid1.col = TempCol
End Sub

Private Sub SetOptionButton(Optional ByVal iCol As Integer, Optional ByVal iRow As Integer)
    Dim X As Long
    Dim Y As Long
    Dim TempRow As Long
    Dim TempCol As Long
            
    If SetColumnTypeArray = False Then
        Exit Sub
    End If
    
    TempRow = MSFlexGrid1.Row
    TempCol = MSFlexGrid1.col
    MSFlexGrid1.Redraw = False
    For X = 1 To MSFlexGrid1.Rows - 1
        For Y = 0 To MSFlexGrid1.Cols - 1
            If ColumnType(Y) = eOption Then
                'If Len(MSFlexGrid1.TextMatrix(X, Y)) = 0 Then
                    MSFlexGrid1.Row = X
                    MSFlexGrid1.col = Y
                    MSFlexGrid1.CellPictureAlignment = 4 'center x center
                    Set MSFlexGrid1.CellPicture = OptionOff.Picture
                    MSFlexGrid1.Text = " "
                    'MSFlexGrid1.CellForeColor = vbWhite
                'End If
            End If
        Next Y
    Next X
    MSFlexGrid1.Row = TempRow
    MSFlexGrid1.col = TempCol
    MSFlexGrid1.Redraw = True
End Sub

Public Property Let ColWidth(ByVal col As Long, ByVal newWidth As Long)
    If newWidth <= 0 Then
      MSFlexGrid1.ColWidth(col) = 0
    Else
      MSFlexGrid1.ColWidth(col) = newWidth
    End If
    PropertyChanged "ColWidth"
    
End Property

Public Property Get ColWidth(ByVal col As Long) As Long
    ColWidth = MSFlexGrid1.ColWidth(col)
    
End Property


'-----igit 11/06/01----------------------------

Public Property Get AllowUserResizing() As AllowUserResizeSettings
   AllowUserResizing = MSFlexGrid1.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As AllowUserResizeSettings)
   MSFlexGrid1.AllowUserResizing() = New_AllowUserResizing
   PropertyChanged "AllowUserResizing"
End Property

Public Property Let RowHeight(ByVal Index As Long, ByVal newHeight As Long)
   MSFlexGrid1.RowHeight(Index) = newHeight
   PropertyChanged "RowHeight"
End Property

Public Property Get CellHeight() As Long
   CellHeight = MSFlexGrid1.CellHeight
End Property

Public Property Get CellWidth() As Long
  CellWidth = MSFlexGrid1.CellWidth
End Property

Public Property Get CellTop() As Long
  CellTop = MSFlexGrid1.CellTop
End Property

Public Property Get CellLeft() As Long
  CellLeft = MSFlexGrid1.CellLeft
End Property

Public Property Let TopRow(ByVal NewTopRow As Long)
  MSFlexGrid1.TopRow = NewTopRow
End Property

Public Property Get TopRow() As Long
  TopRow = MSFlexGrid1.TopRow
End Property

Public Property Get Top() As Long
  Top = MSFlexGrid1.Top
End Property

Public Property Let LeftCol(ByVal NewLeftCol As Long)
  MSFlexGrid1.LeftCol = NewLeftCol
End Property

Public Property Get LeftCol() As Long
  LeftCol = MSFlexGrid1.LeftCol
End Property

Public Property Get Left() As Long
  Left = MSFlexGrid1.Left
End Property

Public Property Get CellBackColor() As OLE_COLOR
  CellBackColor = MSFlexGrid1.CellBackColor
End Property

Public Property Let CellBackColor(ByVal Newcolor As OLE_COLOR)
  MSFlexGrid1.CellBackColor = Newcolor
  PropertyChanged "CellBackColor"
End Property

Public Property Let CellForeColor(ByVal Newcolor As OLE_COLOR)
  MSFlexGrid1.CellForeColor = Newcolor
  PropertyChanged "CellForeColor"
End Property

Property Get CellForeColor() As OLE_COLOR
  CellForeColor = MSFlexGrid1.CellForeColor
End Property
 
Property Get GrdKeyDown() As Integer
  GrdKeyDown = flexKeyDown
End Property

Property Get CellTextStyle() As TextStyleSettings
  CellTextStyle = MSFlexGrid1.CellTextStyle
End Property

Property Let CellTextStyle(ByVal newValue As TextStyleSettings)
  MSFlexGrid1.CellTextStyle = newValue
  PropertyChanged "CellTextStyle"
End Property

Property Get FillStyle() As FillStyleSettings
  FillStyle = MSFlexGrid1.FillStyle
End Property

Property Let FillStyle(ByVal newValue As FillStyleSettings)
  MSFlexGrid1.FillStyle = newValue
  PropertyChanged "FillStyle"
End Property
           '--- 02/01/02 igit
           
Property Get ComboValue() As String
  'ComboValue = Combo1(FrstCol).Text
  ComboValue = CallByName(Combo1(MSFlexGrid1.col), "Text", VbGet)
End Property

Property Get ComboText() As String
  ComboText = Combo1(MSFlexGrid1.col).List(Combo1(MSFlexGrid1.col).ListIndex)
End Property
  
Property Get GrdTxtBoxValue() As String
   GrdTxtBoxValue = txtBoxValue
End Property
  
Property Get TextBoxValue() As String
  TextBoxValue = zTxtValue
End Property
  
Property Let AutoDropDown(ByVal newValue As Boolean)
  zValue = newValue
  PropertyChanged "AutoDropDown"
End Property

Property Get AutoDropDown() As Boolean
  AutoDropDown = zValue
End Property

'----------------------------------------------------

Public Property Get Rowsel() As Long
    Rowsel = MSFlexGrid1.Rowsel
End Property

Public Property Let Rowsel(ByVal NewRowSel As Long)
    MSFlexGrid1.Rowsel = NewRowSel
End Property

Public Property Get Colsel() As Long
    Colsel = MSFlexGrid1.Colsel
End Property

Public Property Let Colsel(ByVal NewColSel As Long)
    MSFlexGrid1.Colsel = NewColSel
End Property

'number of recordset
Public Property Get TotalRecord() As Long
    TotalRecord = MyRecord
End Property

'set view progress
Public Property Get ViewProgress() As Boolean
    ViewProgress = Picture1.Visible
End Property

Public Property Let ViewProgress(ByVal NewViewProgress As Boolean)
    Picture1.Visible = NewViewProgress
    PropertyChanged "ViewProgress"
End Property

'set redraw
Public Property Get Redraw() As Boolean
    Redraw = MSFlexGrid1.Redraw
End Property

Public Property Let Redraw(ByVal NewRedraw As Boolean)
    MSFlexGrid1.Redraw = NewRedraw
    PropertyChanged "Redraw"
End Property

'set color property
Public Property Let BackColor(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.BackColor = Newcolor
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = MSFlexGrid1.BackColor
End Property

Public Property Let BackColorBkg(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.BackColorBkg = Newcolor
    PropertyChanged "BackColorBkg"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
    BackColorBkg = MSFlexGrid1.BackColorBkg
End Property

Public Property Let BackColorFixed(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.BackColorFixed = Newcolor
    PropertyChanged "BackColorFixed"
End Property

Public Property Get BackColorFixed() As OLE_COLOR
    BackColorFixed = MSFlexGrid1.BackColorFixed
End Property

Public Property Let BackColorSel(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.BackColorSel = Newcolor
    PropertyChanged "BackColorSel"
End Property

Public Property Get BackColorSel() As OLE_COLOR
    BackColorSel = MSFlexGrid1.BackColorSel
End Property

Public Property Let ForeColor(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.ForeColor = Newcolor
    PropertyChanged "ForeColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = MSFlexGrid1.ForeColor
End Property

Public Property Let ForeColorFixed(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.ForeColorFixed = Newcolor
    PropertyChanged "ForeColorFixed"
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
    ForeColorFixed = MSFlexGrid1.ForeColorFixed
End Property

Public Property Let ForeColorSel(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.ForeColorSel = Newcolor
    PropertyChanged "ForeColorSel"
End Property

Public Property Get ForeColorSel() As OLE_COLOR
    ForeColorSel = MSFlexGrid1.ForeColorSel
End Property

Public Property Let GridColor(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.GridColor = Newcolor
    PropertyChanged "GridColor"
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = MSFlexGrid1.GridColor
End Property

Public Property Let GridColorFixed(ByVal Newcolor As OLE_COLOR)
    MSFlexGrid1.GridColorFixed = Newcolor
    PropertyChanged "GridColorFixed"
End Property

Public Property Get GridColorFixed() As OLE_COLOR
    GridColorFixed = MSFlexGrid1.GridColorFixed
End Property
'end of set color property
'set font
Public Property Get Font() As IFontDisp
   Set Font = MSFlexGrid1.Font
End Property

Public Property Set Font(ByVal New_Font As IFontDisp)
    Set MSFlexGrid1.Font = New_Font
    PropertyChanged "Font"
End Property
'end of set font

'set mousepointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ' Validation is supplied by UserControl.
    Let UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property
'end of set mousepointer

'set mouseicon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
'end of set mouseicon

'set record alignment for all record display in msflexgrid
Public Property Get RecordAlignment() As AlignmentSettings
   RecordAlignment = MyAlignment
End Property

Public Property Let RecordAlignment(ByVal NewAlignment As AlignmentSettings)
   MyAlignment = NewAlignment
   PropertyChanged "RecordAlignment"
End Property
'end of set record

'set database name
Public Property Let DataName(ByVal newValue As Database)
    Set MyDataName = newValue
End Property

Public Property Get DataName() As Database
    Set MyDataName = DataName
End Property
'end of database setting

'set samadaa boleh edit atau tidak
Public Property Get EditEnable() As Boolean
Attribute EditEnable.VB_ProcData.VB_Invoke_Property = "General"
   EditEnable = MyEdit
End Property

Public Property Let EditEnable(ByVal NewEditEnable As Boolean)
   MyEdit = NewEditEnable
   'popertyChanged "EditEnable"
End Property
'end of edit enable

Public Property Get EditTextLenght() As Long
    EditTextLenght = Text1.MaxLength
End Property

Public Property Let EditTextLenght(ByVal NewEditLenght As Long)
    Text1.MaxLength = NewEditLenght
    PropertyChanged "EditTextLenght"
End Property

'property autofix
Public Property Get AutoFixCol() As Boolean
   AutoFixCol = AutoFix
End Property

Public Property Let AutoFixCol(ByVal NewAutoFixCol As Boolean)
   AutoFix = NewAutoFixCol
   PropertyChanged "AutoFixCol"
End Property

Public Property Get AddWidth() As Long
   AddWidth = ModifyWidth
End Property

Public Property Let AddWidth(ByVal NewModifyValue As Long)
   ModifyWidth = NewModifyValue
   PropertyChanged "AddWidth"
End Property

Public Property Get GridEnabled() As Boolean
Attribute GridEnabled.VB_ProcData.VB_Invoke_Property = "General"
   GridEnabled = MSFlexGrid1.Enabled
End Property

Public Property Let GridEnabled(ByVal newValue As Boolean)
   MSFlexGrid1.Enabled = newValue
   Text1.Text = ""
   If newValue = False Then
     Text1.Visible = newValue
     Combo1(LastCol).Visible = newValue
     DTPicker1.Visible = newValue      '--igit
     DTPicker2.Visible = newValue      '--igit
   End If
End Property

Public Property Get HideCol(ByVal col As Long) As Boolean
    If MSFlexGrid1.ColWidth(col) = 0 Then
        HideCol = True
    Else
        HideCol = False
    End If
End Property
Public Property Let HideCol(ByVal col As Long, ByVal newValue As Boolean)
    Dim X As Long
    If newValue = True Then
        If ColumnType(col) = eCheckbox Then
            With MSFlexGrid1
                .col = col
                For X = 1 To .Rows - 1
                    .Row = X
                    If .CellPicture = imgUnchecked.Picture Then
                        .Text = "U"
                    ElseIf .CellPicture = imgChecked.Picture Then
                        .Text = "C"
                    End If
                    Set .CellPicture = Nothing
                Next X
            End With
        End If
        
        If ColumnType(col) = eOption Then
            With MSFlexGrid1
                .col = col
                For X = 1 To .Rows - 1
                    .Row = X
                    If .CellPicture = OptionOff.Picture Then
                        .Text = "0"
                    ElseIf .CellPicture = OptionOn.Picture Then
                        .Text = "1"
                    End If
                    Set .CellPicture = Nothing
                Next X
            End With
        End If
        MSFlexGrid1.ColWidth(col) = 0
    Else
        If ColumnType(col) = eCheckbox Then
            With MSFlexGrid1
                .col = col
                For X = 1 To .Rows - 1
                    .Row = X
                    If .Text = "U" Then
                        Set .CellPicture = imgUnchecked.Picture
                    ElseIf .Text = "C" Then
                        Set .CellPicture = imgChecked.Picture
                    End If
                    .Text = " "
                Next X
            End With
        End If
        
        If ColumnType(col) = eOption Then
            With MSFlexGrid1
                .col = col
                For X = 1 To .Rows - 1
                    .Row = X
                    If .Text = "0" Then
                        Set .CellPicture = OptionOff.Picture
                    ElseIf .Text = "1" Then
                        Set .CellPicture = OptionOn.Picture
                    End If
                    .Text = " "
                Next X
            End With
        End If
        MSFlexGrid1.ColWidth(col) = MSFlexGrid1.ColWidth(col - 1)
    End If
End Property

Public Property Get HideRow(ByVal Row As Long) As Boolean
    If MSFlexGrid1.RowHeight(Row) = 0 Then
        HideRow = True
    Else
        HideRow = False
    End If
End Property

Public Property Let HideRow(ByVal Row As Long, ByVal newValue As Boolean)
    Dim X As Long
    With MSFlexGrid1
        If newValue = True Then
            .Row = Row
            For X = 1 To .Cols - 1
                If ColumnType(X) = eCheckbox Then
                    .col = X
                    If .CellPicture = imgUnchecked.Picture Then
                        .Text = "U"
                    ElseIf .CellPicture = imgChecked.Picture Then
                        .Text = "C"
                    End If
                    Set .CellPicture = Nothing
                End If
                
                If ColumnType(X) = eOption Then
                    .col = X
                    If .CellPicture = OptionOff.Picture Then
                        .Text = "0"
                    ElseIf .CellPicture = OptionOn.Picture Then
                        .Text = "1"
                    End If
                    Set .CellPicture = Nothing
                End If
                
            Next X
            
            
            
            MSFlexGrid1.RowHeight(Row) = 0
        Else
            .Row = Row
            For X = 1 To .Cols - 1
                If ColumnType(X) = eCheckbox Then
                    .col = X
                    If .Text = "U" Then
                        Set .CellPicture = imgUnchecked.Picture
                    ElseIf .Text = "C" Then
                        Set .CellPicture = imgChecked.Picture
                    End If
                    .Text = " "
                End If
                
                If ColumnType(X) = eOption Then
                    .col = X
                    If .Text = "0" Then
                        Set .CellPicture = OptionOff.Picture
                    ElseIf .Text = "1" Then
                        Set .CellPicture = OptionOn.Picture
                    End If
                    .Text = " "
                End If
            Next X
            MSFlexGrid1.RowHeight(Row) = MSFlexGrid1.RowHeight(Row - 1)
        End If
    End With
End Property

Public Property Get GridLines() As CoolFlexGridLines
   GridLines = MSFlexGrid1.GridLines
End Property

Public Property Let GridLines(ByVal newValue As CoolFlexGridLines)
   MSFlexGrid1.GridLines = newValue
End Property

'--- igit 11/06/01-------------------
Public Property Get CellAlignment() As SetAlignment 'FlexCellAlignment
   CellAlignment = MSFlexGrid1.CellAlignment
End Property

Public Property Let CellAlignment(ByVal newValue As SetAlignment)
   MSFlexGrid1.CellAlignment = newValue
   PropertyChanged "CellAlignment"
End Property

Public Property Let ColAlignment(ByVal Index As Long, ByVal newValue As AlignmentSettings)  'FlexCellAlignment
   MSFlexGrid1.ColAlignment(Index) = newValue
End Property

Public Property Let CellFontBold(ByVal NewVal As Boolean)
   MSFlexGrid1.CellFontBold = NewVal
   PropertyChanged "CellFontBold"
End Property

Public Property Get CellFontBold() As Boolean
Attribute CellFontBold.VB_ProcData.VB_Invoke_Property = "General"
   CellFontBold = MSFlexGrid1.CellFontBold
End Property

Public Property Let CellFontItalic(ByVal NewVal As Boolean)
   MSFlexGrid1.CellFontItalic = NewVal
   PropertyChanged "CellFontItalic"
End Property

Public Property Get CellFontItalic() As Boolean
Attribute CellFontItalic.VB_ProcData.VB_Invoke_Property = "General"
   CellFontItalic = MSFlexGrid1.CellFontItalic
End Property

Public Property Let CellFontSize(ByVal NewVal As Integer)
   MSFlexGrid1.CellFontSize = NewVal
   PropertyChanged "CellFontSize "
End Property

Public Property Get CellFontSize() As Integer
Attribute CellFontSize.VB_ProcData.VB_Invoke_Property = "General"
   CellFontSize = MSFlexGrid1.CellFontSize
End Property

Public Property Get MergeCells() As MergeSettings
   MergeCells = MSFlexGrid1.MergeCells
End Property

Public Property Let MergeCells(ByVal NewVal As MergeSettings)
   MSFlexGrid1.MergeCells = NewVal
   PropertyChanged "MergeCells"
End Property

Public Property Get MergeCol(ByVal Index As Long) As Boolean
   MergeCol = MSFlexGrid1.MergeCol(Index)
End Property

Public Property Let MergeCol(ByVal Index As Long, ByVal NewVal As Boolean)
   MSFlexGrid1.MergeCol(Index) = NewVal
   PropertyChanged "MergeCol"
End Property

Public Property Get MergeRow(ByVal Index As Long) As Boolean
   MergeRow = MSFlexGrid1.MergeRow(Index)
End Property

Public Property Let MergeRow(ByVal Index As Long, ByVal NewVal As Boolean)
   MSFlexGrid1.MergeRow(Index) = NewVal
   PropertyChanged "MergeRow"
End Property

'-------------------------------------
Public Property Get MouseCol() As Integer
   MouseCol = MSFlexGrid1.MouseCol
End Property

Public Property Get MouseRow() As Integer
   MouseRow = MSFlexGrid1.MouseRow
End Property

Public Property Get ScrollBars() As CoolFlexScrollBar
   ScrollBars = MSFlexGrid1.ScrollBars
End Property

Public Property Let ScrollBars(ByVal newValue As CoolFlexScrollBar)
   MSFlexGrid1.ScrollBars = newValue
End Property

Public Property Let Sort(ByVal newValue As CoolFlexSort)
   Text1.Text = ""
   Text1.Visible = False
   'Combo1(LastCol).Text = ""
   Combo1(LastCol).Visible = False
   DTPicker1.Visible = False     '--igit
   DTPicker2.Visible = False     '--igit
   MSFlexGrid1.Sort = newValue
End Property

Public Property Get Tag() As String
   Tag = MSFlexGrid1.Tag
End Property

Public Property Let Tag(ByVal newValue As String)
   MSFlexGrid1.Tag = newValue
End Property

Public Property Get Text() As String
   Text = MSFlexGrid1.Text
   'zTxtValue = MSFlexGrid1.Text
End Property

Public Property Let Text(ByVal newValue As String)
   MSFlexGrid1.Text = newValue
   'zTxtValue = NewValue
   If ColumnType(MSFlexGrid1.col) = eCheckbox Then
      If MSFlexGrid1.Row <> 0 Then
         MSFlexGrid1.CellPictureAlignment = 4 'center x center
         If MSFlexGrid1.CellPicture = imgUnchecked.Picture Then 'MSFlexGrid1.Text = "U" Then
            Set MSFlexGrid1.CellPicture = imgUnchecked.Picture  'LoadPicture(App.Path & "\Checked.bmp")
         Else
            Set MSFlexGrid1.CellPicture = imgChecked.Picture  'LoadPicture(App.Path & "\Checked.bmp")
            'MSFlexGrid1.Text = "C"
         End If
      End If
   End If
End Property

Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_ProcData.VB_Invoke_Property = "General"
   WordWrap = MSFlexGrid1.WordWrap
End Property

Public Property Let WordWrap(ByVal newValue As Boolean)
   MSFlexGrid1.WordWrap = newValue
End Property

Public Property Get ComboBoxListCount(ByVal col As Long) As Long
   ComboBoxListCount = Combo1(col).ListCount
End Property

Public Property Get ComboBoxListIndex(ByVal col As Long) As Long
   ComboBoxListIndex = Combo1(col).ListIndex
End Property

Public Property Let ComboBoxListIndex(ByVal col As Long, ByVal newValue As Long)
   Combo1(col).ListIndex = newValue
End Property

'-- igit-----------------------------
Public Property Get ComboBoxNewIndex(ByVal col As Long) As Long
   ComboBoxNewIndex = Combo1(col).NewIndex
End Property

Public Property Get ComboBoxList(ByVal col As Long, ByVal Index As Integer) As String
   ComboBoxList = Combo1(col).List(Index)
End Property

Public Property Let ComboBoxList(ByVal col As Long, ByVal Index As Integer, ByVal newValue As String)
   Combo1(col).List(Index) = newValue
End Property

Public Property Let TxtBorder(ByVal newValue As txtBorderStyle)
   Text1.BorderStyle = newValue
End Property

Public Property Get TxtBorder() As txtBorderStyle
   TxtBorder = Text1.BorderStyle
End Property

Public Property Let GridSelectionMode(ByVal ModeValue As GrdSelection)
   MSFlexGrid1.SelectionMode = ModeValue
   PropertyChanged "GridSelectionMode"
End Property

Public Property Get GridSelectionMode() As GrdSelection
   GridSelectionMode = MSFlexGrid1.SelectionMode
End Property

Public Property Let AllowBigSelection(ByVal newValue As Boolean)
   MSFlexGrid1.AllowBigSelection = newValue
   PropertyChanged "AllowBigSelection"
End Property

Public Property Get AllowBigSelection() As Boolean
Attribute AllowBigSelection.VB_ProcData.VB_Invoke_Property = "General"
   AllowBigSelection = MSFlexGrid1.AllowBigSelection
End Property

'------------------------------------
Public Property Get ComboBoxItemData(ByVal col As Long, ByVal Index As Integer) As Long
   ComboBoxItemData = Combo1(col).ItemData(Index)
End Property

Public Property Let ComboBoxItemData(ByVal col As Long, ByVal Index As Integer, ByVal newValue As Long)
   Combo1(col).ItemData(Index) = newValue
End Property

Public Property Get FixedCols() As Long
Attribute FixedCols.VB_ProcData.VB_Invoke_Property = "General"
   FixedCols = MSFlexGrid1.FixedCols
End Property

Public Property Let FixedCols(ByVal newValue As Long)
   MSFlexGrid1.FixedCols = newValue
End Property

Public Property Get FixedRows() As Long
Attribute FixedRows.VB_ProcData.VB_Invoke_Property = "General"
   FixedRows = MSFlexGrid1.FixedRows
End Property

Public Property Let FixedRows(ByVal newValue As Long)
   MSFlexGrid1.FixedRows = newValue
End Property
'---igit 02/05/02
Public Sub Headers(Alignment As SetAlignment, ParamArray Hdr() As Variant)
   Dim i
   With MSFlexGrid1
      Cols = UBound(Hdr) + 1
      .Row = 0
      For i = LBound(Hdr) To UBound(Hdr)
         .col = i: .Text = Hdr(i): .CellAlignment = Alignment
      Next
   End With
End Sub

Public Sub ColSize(ParamArray zSize() As Variant)
   Dim i
   With MSFlexGrid1
      Cols = UBound(zSize) + 1
      .Row = 0
      For i = LBound(zSize) To UBound(zSize)
         .ColWidth(i) = zSize(i)
      Next
   End With
End Sub
'-------------------------------------

'----igit 05/12/04-------------
Public Sub SetFocus()
   MSFlexGrid1.SetFocus
End Sub
'-----------------------------

Private Sub UserControl_Terminate()
   If Text1.Visible = True Then
      'MSFlexGrid1.Text = Text1.Text
       'zTxtValue = Text1.Text
       'Text1.Text = ""
       Text1.Visible = False
    End If
    
    If Combo1(LastCol).Visible = True Then
       'MSFlexGrid1.Text = Combo1(LastCol).Text
       'Combo1(LastCol).Text = ""
       Combo1(LastCol).Visible = False
    End If
    
    If DTPicker1.Visible = True Then
       'MSFlexGrid1.Text = DTPicker1.Value
       DTPicker1.Visible = False
    End If
       
     If DTPicker2.Visible = True Then
       'MSFlexGrid1.Text = DTPicker1.Value
       DTPicker2.Visible = False
    End If
End Sub

'--- igit 02/19/02----------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "Footer", Me.Footer, 1
   PropBag.WriteProperty "ForeColor", Me.ForeColor, &H80000006
   PropBag.WriteProperty "ForeColorFixed", Me.ForeColorFixed, &H80000012
   PropBag.WriteProperty "ForeColorSel", Me.ForeColorSel, &H8000000E
   PropBag.WriteProperty "WordWrap", Me.WordWrap, False
   PropBag.WriteProperty "BackColor", Me.BackColor, &HFFFFFF
   PropBag.WriteProperty "BackColorBkg", Me.BackColorBkg, &HC0C0C0
   PropBag.WriteProperty "BackColorFixed", Me.BackColorFixed, &HE0E0E0
   PropBag.WriteProperty "BackColorSel", Me.BackColorSel, &H80000002
   PropBag.WriteProperty "AllowBigSelection", Me.AllowBigSelection, False
   PropBag.WriteProperty "Cols", Me.Cols, 2
   PropBag.WriteProperty "Rows", Me.Rows, 2
   PropBag.WriteProperty "EditEnable", Me.EditEnable, False
   PropBag.WriteProperty "AllowUserResizing", MSFlexGrid1.AllowUserResizing, 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Me.Footer = PropBag.ReadProperty("Footer", 1)
   Me.ForeColor = PropBag.ReadProperty("ForeColor", &H80000006)
   Me.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", &H80000012)
   Me.ForeColorSel = PropBag.ReadProperty("ForeColorSel", &H8000000E)
   Me.WordWrap = PropBag.ReadProperty("WordWrap", False)
   Me.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
   Me.BackColorBkg = PropBag.ReadProperty("BackColorBkg", &HC0C0C0)
   Me.BackColorFixed = PropBag.ReadProperty("BackColorFixed", &HE0E0E0)
   Me.BackColorSel = PropBag.ReadProperty("BackColorSel", &H80000002)
   Me.AllowBigSelection = PropBag.ReadProperty("AllowBigSelection", False)
   Me.Cols = PropBag.ReadProperty("Cols", 2)
   Me.Rows = PropBag.ReadProperty("Rows", 2)
   Me.EditEnable = PropBag.ReadProperty("EditEnable", False)
   MSFlexGrid1.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", 1)
End Sub
'-------------------------------------
'---igit 04/28/02----------------------------
Public Property Get ChkValue(ByVal Row, ByVal col As Long) As Boolean
   If ColumnType(col) = eCheckbox Then
      MSFlexGrid1.Row = Row
      MSFlexGrid1.col = col
      If MSFlexGrid1.Text = "U" Or MSFlexGrid1.CellPicture = imgUnchecked.Picture Then
         ChkValue = False
      Else
         ChkValue = True
      End If
   Else
      MsgBox "The column you specify in not a checkbox....", 64, "Message"
   End If
End Property
Public Property Let ChkValue(ByVal Row, ByVal col As Long, newValue As Boolean)
   If ColumnType(col) = eCheckbox Then
      MSFlexGrid1.Row = Row
      MSFlexGrid1.col = col
      If newValue Then 'MSFlexGrid1.Text = "C" Then
          Set MSFlexGrid1.CellPicture = imgChecked.Picture  'LoadPicture(App.Path & "\Checked.bmp")
         MSFlexGrid1.Text = "C"
      Else
          Set MSFlexGrid1.CellPicture = imgUnchecked.Picture  'LoadPicture(App.Path & "\Checked.bmp")
          MSFlexGrid1.Text = "U"
      End If
   Else
      MsgBox "The column you specify in not a checkbox....", 64, "Message"
   End If
End Property

Public Property Get ScrollTrack() As Boolean
   ScrollTrack = MSFlexGrid1.ScrollTrack
End Property

Public Property Let ScrollTrack(ByVal vNewValue As Boolean)
   MSFlexGrid1.ScrollTrack = vNewValue
End Property
'----------igit 04/23/04--------------
Public Property Get ColLock(ByVal nIndex As Integer) As Boolean
   ColLock = ColumnLock(nIndex)
End Property

Public Property Let ColLock(ByVal nIndex As Integer, ByVal nValue As Boolean)
   ColumnLock(nIndex) = nValue
End Property

