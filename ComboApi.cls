VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboBoxAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Author: Billy Conner
'---------------------------------
'   ABOUT
'---------------------------------
    'This allows the user to modify comboboxes via API
    'that includes your controls as well as external apps.

'---------------------------------
'   MORE INFO
'---------------------------------
    'The Friend procedures are not really part of a real vb Combobox
    'It took me a while to figure out the SelStart,SelLength,SelText
    'There are still some Minor bugs to work out, and a few more things i can add to this.

'---------------------------------
'   PLEASE NOTE
'---------------------------------
    'The combobox class will not operate properly unless you specify a valid handle (HWND)
    'In other words make sure you set Class.Hwnd = (Valid combobox handle)

Option Explicit

Public Enum eStyle
    DropDownCombo = 2
    DropDownList = 3
    SimpleCombo = 1
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
     x As Long
     y As Long
End Type

Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal aBOOL As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Const RDW_INVALIDATE = &H1
Private Const WM_PAINT As Long = &HF
Private Const GWL_STYLE = (-16)
Private Const WS_VISIBLE As Long = &H10000000

' Combo Box return Values (may use these more in the future)
Const CB_OKAY = 0
Const CB_ERR = (-1)
Const CB_ERRSPACE = (-2)

' Combo Box Notification Codes (I probably wont use these)
Const CBN_ERRSPACE = (-1)
Const CBN_SELCHANGE = 1
Const CBN_DBLCLK = 2
Const CBN_SETFOCUS = 3
Const CBN_KILLFOCUS = 4
Const CBN_EDITCHANGE = 5
Const CBN_EDITUPDATE = 6
Const CBN_DROPDOWN = 7
Const CBN_CLOSEUP = 8
Const CBN_SELENDOK = 9
Const CBN_SELENDCANCEL = 10

' Combo Box styles
Const CBS_SIMPLE = &H1& 'Using This
Const CBS_DROPDOWN = &H2& 'Using This
Const CBS_DROPDOWNLIST = &H3& 'Using This
Const CBS_OWNERDRAWFIXED = &H10&
Const CBS_OWNERDRAWVARIABLE = &H20&
Const CBS_AUTOHSCROLL = &H40&
Const CBS_OEMCONVERT = &H80&
Const CBS_SORT = &H100& 'Using This
Const CBS_HASSTRINGS = &H200&
Const CBS_NOINTEGRALHEIGHT = &H400&
Const CBS_DISABLENOSCROLL = &H800& 'Using This

' Combo Box messages
Const CB_GETEDITSEL = &H140 'Using This
Const CB_LIMITTEXT = &H141
Const CB_SETEDITSEL = &H142 'Using This
Const CB_ADDSTRING = &H143 'No need for this since i use Insert String
Const CB_DELETESTRING = &H144 'Using This
Const CB_DIR = &H145
Const CB_GETCOUNT = &H146 'Using This
Const CB_GETCURSEL = &H147 'Using This
Const CB_GETLBTEXT = &H148 'Using This
Const CB_GETLBTEXTLEN = &H149
Const CB_INSERTSTRING = &H14A 'Using This
Const CB_RESETCONTENT = &H14B 'Using This
Const CB_FINDSTRING = &H14C 'Using This
Const CB_SELECTSTRING = &H14D
Const CB_SETCURSEL = &H14E 'Using This
Const CB_SHOWDROPDOWN = &H14F 'Using This
Const CB_GETITEMDATA = &H150 'Using This
Const CB_SETITEMDATA = &H151 'Using This
Const CB_GETDROPPEDCONTROLRECT = &H152
Const CB_SETITEMHEIGHT = &H153
Const CB_GETITEMHEIGHT = &H154
Const CB_SETEXTENDEDUI = &H155
Const CB_GETEXTENDEDUI = &H156
Const CB_GETDROPPEDSTATE = &H157 'Using This
Const CB_FINDSTRINGEXACT = &H158 'Using This
Const CB_SETLOCALE = &H159 'Using This
Const CB_GETLOCALE = &H15A 'Using This
Const CB_MSGMAX = &H15B


Private MyHwnd As Long
Private iData As Long
'Public NewIndex As Long
'--Selecting Vars
Private sStart As Long
Private sLength As Long

Public Property Get hwnd() As Long
    
    hwnd = MyHwnd

End Property

Public Property Let hwnd(Value As Long)

    MyHwnd = Value
    
    If IsComboBox = False Then 'if not a combobox
        MyHwnd = 0 'set to an invalid handle
    End If
    
End Property


Public Property Get Text() As String
    
    Dim sItemText As String * &H7FFF
    
    GetWindowText MyHwnd, sItemText, &H7FFF
    Text = sItemText

End Property

Public Property Let Text(pText As String)

    SetWindowText MyHwnd, pText

End Property

Public Property Get SelStart() As Long
    
    sStart = SendMessageLong(MyHwnd, CB_GETEDITSEL, 0&, 0&) Mod &H10000
    SelStart = sStart

End Property

Public Property Let SelStart(Value As Long)
    
    sStart = Value
    SendMessageLong MyHwnd, CB_SETEDITSEL, 0&, (sStart * &H10000) + sStart

End Property

Public Property Get SelLength() As Long
    
    Dim FirstPos As Long
    Dim LastPos As Long
    
    sLength = SendMessageLong(MyHwnd, CB_GETEDITSEL, 0&, 0&)
    If sLength = CB_ERR Then
        Exit Property
    End If
    LastPos = sLength \ &H10000
    FirstPos = sLength Mod &H10000
    SelLength = LastPos - FirstPos
    
End Property

Public Property Let SelLength(Value As Long)
    
    sLength = Value
    SendMessageLong MyHwnd, CB_SETEDITSEL, 0&, (sStart * &H10000) + (sLength + sStart)

End Property

Public Property Get SelText() As String
    
    'basically Get The Text in the Selected String (if any)
    Dim FirstPos As Long
    Dim LastPos As Long
    Dim sItemText As String * &H7FFF
    Dim iText As String

    sLength = SendMessageLong(MyHwnd, CB_GETEDITSEL, 0&, 0&)
    If sLength = CB_ERR Then
        Exit Property
    End If
    LastPos = (sLength \ &H10000) + 1
    FirstPos = (sLength Mod &H10000) + 1
    If (LastPos - FirstPos <= 0) Then
        Exit Property
    End If
    GetWindowText MyHwnd, sItemText, &H7FFF
    iText = Left(sItemText, InStr(sItemText, vbNullChar) - 1)
    If iText = vbNullString Then
        Exit Property
    End If
    SelText = Mid(iText, FirstPos, LastPos - FirstPos)


End Property

Public Property Let SelText(pText As String)
    
    'basically insert the string where the Start is..and if any Selected replace them also..
    
    Dim FirstPos As Long
    Dim LastPos As Long
    Dim iText As String
    Dim sItemText As String * &H7FFF
    
    sLength = SendMessageLong(MyHwnd, CB_GETEDITSEL, 0&, 0&)
    If sLength = CB_ERR Then
        Exit Property
    End If
    LastPos = sLength \ &H10000
    FirstPos = sLength Mod &H10000
    GetWindowText MyHwnd, sItemText, &H7FFF
    iText = Left(sItemText, InStr(sItemText, vbNullChar) - 1)
    iText = Left(iText, FirstPos) & pText & Mid(iText, LastPos + 1)
    SetWindowText MyHwnd, iText
    Me.SelLength = 0

End Property

Public Property Get ListIndex() As Long
 
  ListIndex = SendMessageLong(MyHwnd, CB_GETCURSEL, 0&, 0&)

End Property

Public Property Let ListIndex(Index As Long)

    SendMessageLong MyHwnd, CB_SETCURSEL, Index, 0&

End Property

'Search for first finding of Existance of a string starting at the left (this is not like Instr(), its like Left())
Public Function Search(pText As String, Optional StartIndex As Long = 0) As Long

    Search = SendMessageStr(MyHwnd, CB_FINDSTRING, StartIndex, pText)

End Function

'Search for first finding of Existance of an exact string
Public Function SearchExact(pText As String, Optional StartIndex As Long = 0) As Long

    SearchExact = SendMessageStr(MyHwnd, CB_FINDSTRINGEXACT, StartIndex, pText)

End Function

Public Property Get Sorted() As Boolean

    Sorted = GetWindowStyle(CBS_SORT)

End Property

Public Property Get Count() As Long
    
    Count = SendMessageLong(MyHwnd, CB_GETCOUNT, 0&, 0&)

End Property

Public Sub AddItem(pText As String, Data As Long, Optional pListIndex As Long = -1)
  
   'add to the combobox
    If pListIndex = -1 Then 'if no listitem selected
        pListIndex = Me.Count 'set it = count (to be added last in list)
    End If
    SendMessageStr MyHwnd, CB_INSERTSTRING, ByVal pListIndex, pText
    Me.ItemData(pListIndex) = Data 'dont forget the data

End Sub

Public Property Get ItemData(pListIndex As Long) As Long

    ItemData = SendMessage(MyHwnd, CB_GETITEMDATA, ByVal pListIndex, ByVal 0&)

End Property

Public Property Let ItemData(pListIndex As Long, Data As Long)

    SendMessage MyHwnd, CB_SETITEMDATA, ByVal pListIndex, ByVal Data

End Property

Public Sub RemoveItem(pListIndex As Long)

    SendMessageLong MyHwnd, CB_DELETESTRING, pListIndex, 0&

End Sub

Public Property Get List(pListIndex As Long) As String
    
    Dim sItemText As String * &H7FFF
    
    SendMessageStr MyHwnd, CB_GETLBTEXT, pListIndex, ByVal sItemText
    List = Left(sItemText, InStr(sItemText, Chr(0)) - 1)
End Property

Public Property Let List(pListIndex As Long, pText As String)
    
    Dim tmpData As Long
    
    If Me.Count Then
        If Len(pText) = 0 Then
            pText = Chr(32)
        End If
        'Get The old Data
        tmpData = Me.ItemData(ListIndex)
        'delete the current item
        Me.RemoveItem pListIndex
        'insert a new one
        'Add the old Data also
        Me.AddItem pText, tmpData, pListIndex
    End If
    
End Property

Public Property Get Style() As eStyle
    
    'this is working for my apps. but i know its probably not correct
    Dim iStyle As Long
    Dim DDStyle As eStyle
    
    iStyle = GetWindowLong(MyHwnd, GWL_STYLE)
    If ((iStyle And CBS_DROPDOWNLIST) = CBS_DROPDOWNLIST) Then
        DDStyle = eStyle.DropDownList
        Style = DDStyle
        Exit Property
    End If
    
    If ((iStyle And CBS_DROPDOWN) = CBS_DROPDOWN) Then
        DDStyle = eStyle.DropDownCombo
        Style = DDStyle
        Exit Property
    End If
    
    If ((iStyle And CBS_SIMPLE) = CBS_SIMPLE) Then
        DDStyle = eStyle.SimpleCombo
        Style = DDStyle
        Exit Property
    End If
    Style = DDStyle 'if this is reached its assuming the Object is Invalid

End Property

Public Sub Clear()

    SendMessage MyHwnd, CB_RESETCONTENT, CLng(True), 0&

End Sub

Public Property Get Enabled() As Boolean

    Enabled = IsWindowEnabled(MyHwnd)

End Property

Public Property Let Enabled(Value As Boolean)

    EnableWindow MyHwnd, Value

End Property

Public Property Get Visible() As Boolean

    Visible = IsWindowVisible(MyHwnd)

End Property
Public Property Let Visible(Value As Boolean)

    Dim rct As RECT
    
    SetWindowStyle WS_VISIBLE, Value
    RedrawWindow Me.Parent, rct, 1&, RDW_INVALIDATE
End Property

Public Property Get Parent() As Long

    Parent = GetParent(MyHwnd)

End Property

Public Property Let Parent(Handle As Long)

    SetParent MyHwnd, Handle

End Property


Friend Function Exists() As Boolean

    Exists = IsWindow(MyHwnd)

End Function

Private Function IsComboBox() As Boolean

    'Lets Check To see if it has a Combobox type Style
    If (Me.Style <> 0) Then 'if style is 0 then its not a combobox i assume(or its a lsitbox)
        IsComboBox = True
        Exit Function
    End If
    'Style Passed, Lets see if Logic works out(only listboxes make it this far i believe)
    '   basically if no items exist then lets test to see if index(0) has a string in it,
    '   if so then its not a combobox
    If Me.Count = 0 And Len(Me.List(0)) > 0 Then
        IsComboBox = False
    End If

End Function

'this checks to see if dropdownlist showing
Friend Property Get DroppedState() As Boolean

    DroppedState = CBool(SendMessageLong(MyHwnd, CB_GETDROPPEDSTATE, 0&, 0&))

End Property

'This allows you to set whether the DropDownbox is showing
Friend Property Let DroppedState(Value As Boolean)

    SendMessageLong MyHwnd, CB_SHOWDROPDOWN, CLng(Value), 0&

End Property

Friend Property Get Scrollable() As Boolean
    
    Scrollable = GetWindowStyle(CBS_DISABLENOSCROLL)
    
End Property

Public Sub Refresh()
    
    SendMessage MyHwnd, WM_PAINT, -1&, 0&

End Sub

Public Sub SetFocus()

    Putfocus MyHwnd

End Sub

Public Property Get Top() As Long
    
    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    Top = Wpl.rcNormalPosition.Top

End Property
Public Property Let Top(Value As Long)

    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    'we have to move the bottom side first in order to keep the same height
    Wpl.rcNormalPosition.Bottom = Wpl.rcNormalPosition.Bottom + (Value - Wpl.rcNormalPosition.Top)
    Wpl.rcNormalPosition.Top = Value
    SetWindowPlacement MyHwnd, Wpl
    
End Property

Public Property Get Left() As Long
    
    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    Left = Wpl.rcNormalPosition.Left
    
End Property
Public Property Let Left(Value As Long)

    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    'we have to move the right side first in order to keep the same width
    Wpl.rcNormalPosition.Right = Wpl.rcNormalPosition.Right + (Value - Wpl.rcNormalPosition.Left)
    Wpl.rcNormalPosition.Left = Value
    SetWindowPlacement MyHwnd, Wpl
 
End Property

Public Property Get Height() As Long
    
    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    Height = Abs(Wpl.rcNormalPosition.Bottom - Wpl.rcNormalPosition.Top)
    
End Property
Public Property Let Height(Value As Long)

    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    Wpl.rcNormalPosition.Bottom = Value + Wpl.rcNormalPosition.Top
    SetWindowPlacement MyHwnd, Wpl
 
End Property

Public Property Get Width() As Long

    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    Width = Abs(Wpl.rcNormalPosition.Right - Wpl.rcNormalPosition.Left)
    
End Property
Public Property Let Width(Value As Long)

    Dim Wpl As WINDOWPLACEMENT
    
    Wpl.Length = Len(Wpl)
    Call GetWindowPlacement(MyHwnd, Wpl)
    Wpl.rcNormalPosition.Right = Abs(Value + Wpl.rcNormalPosition.Left)
    SetWindowPlacement MyHwnd, Wpl
 
End Property

Private Sub SetWindowStyle(pStyle As Long, Value As Boolean)
    
    Dim iStyle As Long
    
    iStyle = GetWindowLong(MyHwnd, GWL_STYLE)
    Select Case Value
        Case True
            iStyle = iStyle Or pStyle
        Case False
            iStyle = iStyle And (Not pStyle)
    End Select
    SetWindowLong MyHwnd, GWL_STYLE, iStyle
End Sub

Private Function GetWindowStyle(pStyle As Long) As Boolean
    
    GetWindowStyle = GetWindowLong(MyHwnd, GWL_STYLE) And pStyle

End Function
