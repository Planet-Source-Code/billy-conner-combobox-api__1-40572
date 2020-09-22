VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combo Box Tool"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox GridCb 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox GridTb 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Button 
      Caption         =   "Search"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Button 
      Caption         =   "Clear Items"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Button 
      Caption         =   "Remove Item"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Button 
      Caption         =   "Add Item"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid Grid1 
      Height          =   6375
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   11245
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   2
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   3720
   End
   Begin VB.CommandButton Button 
      Caption         =   "Setup && Start"
      Height          =   375
      Index           =   0
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Cb As ComboBoxAPI

Private Sub Button_Click(Index As Integer)
Dim iCaption As String
Dim iData As Long
Dim i As Long
Dim SearchStr As String
'Dim iStr As String

Select Case Index
    
    Case 0 'start
        Cb.hwnd = CLng(Val(InputBox("Enter Handle of the ComboBox to Modify", "Handle Needed", Combo1.hwnd)))
        If Cb.hwnd = 0 Then
            MsgBox "Invalid Handle: This Handle is not a Combobox Control"
            Call DoDisableButtons
            Exit Sub
        End If
        Call DoEnableButtons

        Timer1.Enabled = True
        
   Case 1 'AddItem
        Do
            iCaption = InputBox("Enter a Caption For The List Item", "Add Item", "New Item")
        Loop Until iCaption <> vbNullString
        
        iData = Val(InputBox("Enter Item Data (If Any)", "Add Item", "0"))
        
        Cb.AddItem iCaption, iData
        If Cb.hwnd <> Combo1.hwnd Then
            Combo1.AddItem iCaption, iData
        End If
        
    Case 2 'remove item
        Cb.RemoveItem Cb.ListIndex
        If Cb.hwnd <> Combo1.hwnd Then
            Combo1.AddItem iCaption, iData
        End If

    Case 3 'clear
        Cb.Clear
    
    Case 4 'search
        SearchStr = InputBox("Search for what string?", "Search For String", "New")
        If Cb.Search(SearchStr) < 0 Then
            MsgBox "Search Completed, No matches found."
        Else
            MsgBox "First found in ListIndex " & Cb.Search(SearchStr)
        End If
End Select
End Sub

Private Sub Form_Load()

Set Cb = New ComboBoxAPI
Cb.hwnd = Combo1.hwnd
With Grid1
    .BackColorFixed = .BackColor
    .GridLinesFixed = .GridLines
    .GridColorFixed = .GridColor
    .RowHeightMin = GridCb.Height
    .Rows = 0
    .Cols = 2
    .ColWidth(0) = .Width * 0.4 'take 40% of width
    .ColWidth(1) = .Width * 0.575 'take 50% (also allow selection to show)
    'add our items
    .AddItem ("Hwnd")
        .RowData(0) = 99
    .AddItem ("Parent")
        .RowData(1) = 0
    .AddItem ("Style")
        .RowData(2) = 99
    .AddItem ("Visible")
        .RowData(3) = 1
    .AddItem ("Enabled")
        .RowData(4) = 1
    .AddItem ("Sorted")
        .RowData(5) = 99
    .AddItem ("DroppedState")
        .RowData(6) = 1
    .AddItem ("Count")
        .RowData(7) = 99
    .AddItem ("ListIndex")
        .RowData(8) = 0
    .AddItem ("List")
        .RowData(9) = 0
    .AddItem ("Text")
        .RowData(10) = 0
    .AddItem ("ItemData")
        .RowData(11) = 0
    .AddItem ("SelText")
        .RowData(12) = 99
    .AddItem ("SelStart")
        .RowData(13) = 0
    .AddItem ("SelLength")
        .RowData(14) = 0
    .AddItem ("Scrollable")
        .RowData(15) = 99
    .AddItem ("Top")
        .RowData(16) = 0
    .AddItem ("Left")
        .RowData(17) = 0
    .AddItem ("Height")
        .RowData(18) = 99
    .AddItem ("Width")
        .RowData(19) = 0
    
    .Height = .RowHeight(1) * .Rows + 75
    .Col = 1
End With
'Timer1.Enabled = True

End Sub

Private Sub Grid1_DblClick()
If Grid1.RowData(Grid1.Row) = 1 Then
    GridCb.ListIndex = IIf(GridCb.ListIndex = 0, 1, 0)
End If
End Sub

Private Sub Grid1_EnterCell()
'Grid1.CellBackColor = vbRed
GridCb.Visible = False
GridTb.Visible = False
End Sub

Private Sub Grid1_SelChange()
Static InUse As Boolean
Static LastRow As Long
Dim NewRow As Long
NewRow = Grid1.Row

If InUse = False Then
    'ModifyOld one
    InUse = True 'to show the event not to do this when we trigger lastrow
    Grid1.Col = 0
    Grid1.Row = LastRow
    Grid1.CellBackColor = vbWhite
    Grid1.CellForeColor = vbBlack

    Grid1.Row = NewRow
    If Grid1.RowData(NewRow) <> 99 Then
        Grid1.CellBackColor = vbBlue
    Else
        Grid1.CellBackColor = RGB(127, 127, 127) 'gray
    End If
    Grid1.CellForeColor = vbWhite
    Grid1.Col = 1
    
Select Case Grid1.RowData(NewRow)
    Case 1 'its combobox dropdown
        GridCb.Top = Grid1.RowPos(NewRow) + 40
        GridCb.Left = Grid1.ColPos(1) + 35
        GridCb.Width = Grid1.ColWidth(1)
        GridCb.Clear
        GridCb.AddItem "True"
        GridCb.AddItem "False"
        GridCb.ListIndex = IIf(Grid1.TextMatrix(Grid1.Row, Grid1.Col) = "True", 0, 1)
        GridCb.Visible = True
        GridCb.SetFocus
        
    Case 0
        GridTb.Top = Grid1.RowPos(NewRow) + 100
        GridTb.Left = Grid1.ColPos(1) + 40
        GridTb.Width = Grid1.ColWidth(1) - 25
        GridTb.Height = Grid1.RowHeightMin - 70
        GridTb.Text = Grid1.TextMatrix(Grid1.Row, Grid1.Col)
        GridTb.SelStart = 0
        GridTb.SelLength = Len(GridTb.Text)
        GridTb.Visible = True
        GridTb.SetFocus
    End Select
    LastRow = NewRow
    InUse = False
End If
End Sub

Private Sub GridCb_Click()
Static Lastone As Long

If Lastone = GridCb.ListIndex Then
    Exit Sub
End If

Lastone = GridCb.ListIndex
Select Case Grid1.Row
    Case 3
        Cb.Visible = IIf(GridCb.ListIndex = 0, True, False)
    Case 4
        Cb.Enabled = IIf(GridCb.ListIndex = 0, True, False)
    Case 6
        Cb.DroppedState = IIf(GridCb.ListIndex = 0, True, False)
    
End Select
GridCb_Click 'call again 1 more time to make new listindex effective
End Sub

Private Sub GridTb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Select Case Grid1.Row
    Case 1 'parent
        Cb.Parent = Val(GridTb.Text)
    Case 8 'listindex
        Cb.ListIndex = Val(GridTb.Text)
    Case 9 'list
        Cb.List(Cb.ListIndex) = GridTb.Text
    Case 10 'text
        Cb.Text = GridTb.Text
    Case 11 'itemdata
        Cb.ItemData(Cb.ListIndex) = Val(GridTb.Text)
    Case 13 'selStart
        Cb.SelStart = Val(GridTb.Text)
    Case 14 'selLength
        Cb.SelLength = Val(GridTb.Text)
    Case 16 'top
        Cb.Top = Val(GridTb.Text)
    Case 17 'left
        Cb.Left = Val(GridTb.Text)
    Case 19 'width
        Cb.Width = Val(GridTb.Text)
End Select
End If
End Sub

Private Sub Timer1_Timer()
Dim RetVal As eStyle
Dim CbStyle As String
RetVal = Cb.Style
Select Case RetVal
    Case Is = eStyle.DropDownCombo
        CbStyle = "Drop Down Combo"
    Case Is = eStyle.DropDownList
        CbStyle = "Drop Down List"
    Case Is = eStyle.SimpleCombo
        CbStyle = "Simple Combo"
    Case Else
        CbStyle = vbNullString
    End Select
LockWindowUpdate Grid1.hwnd
Grid1.TextMatrix(0, 1) = CStr(Cb.hwnd)
Grid1.TextMatrix(1, 1) = CStr(Cb.Parent)
Grid1.TextMatrix(2, 1) = CbStyle
Grid1.TextMatrix(3, 1) = Cb.Visible
Grid1.TextMatrix(4, 1) = Cb.Enabled
Grid1.TextMatrix(5, 1) = Cb.Sorted
Grid1.TextMatrix(6, 1) = Cb.DroppedState
Grid1.TextMatrix(7, 1) = Cb.Count
Grid1.TextMatrix(8, 1) = Cb.ListIndex
Grid1.TextMatrix(9, 1) = Cb.List(Cb.ListIndex)
Grid1.TextMatrix(10, 1) = Cb.Text
Grid1.TextMatrix(11, 1) = Cb.ItemData(Cb.ListIndex)
Grid1.TextMatrix(12, 1) = Cb.SelText
Grid1.TextMatrix(13, 1) = Cb.SelStart
Grid1.TextMatrix(14, 1) = Cb.SelLength
Grid1.TextMatrix(15, 1) = Cb.Scrollable
Grid1.TextMatrix(16, 1) = Cb.Top
Grid1.TextMatrix(17, 1) = Cb.Left
Grid1.TextMatrix(18, 1) = Cb.Height
Grid1.TextMatrix(19, 1) = Cb.Width
LockWindowUpdate False
End Sub
Private Sub DoDisableButtons()
Dim i As Long

For i = 1 To Button.Count - 1
    Button(i).Enabled = False
Next i
Grid1.Enabled = False
End Sub
Private Sub DoEnableButtons()
Dim i As Long

For i = 1 To Button.Count - 1
    Button(i).Enabled = True
Next i
Grid1.Enabled = True
End Sub
