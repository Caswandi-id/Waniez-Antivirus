VERSION 5.00
Begin VB.UserControl Tab 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   1650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1650
   ScaleWidth      =   3855
   ToolboxBitmap   =   "Tab.ctx":0000
   Begin VB.Image imgBack 
      Height          =   255
      Left            =   720
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblAllTabs 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   195
   End
   Begin VB.Image ioc2 
      Height          =   210
      Left            =   3120
      Picture         =   "Tab.ctx":0312
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image i4 
      Height          =   210
      Index           =   0
      Left            =   1080
      Picture         =   "Tab.ctx":05BC
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image ioc 
      Height          =   210
      Left            =   3360
      Picture         =   "Tab.ctx":0668
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Shape Box 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image io3 
      Height          =   300
      Left            =   3720
      Picture         =   "Tab.ctx":0714
      Top             =   960
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image io2 
      Height          =   300
      Left            =   2640
      Picture         =   "Tab.ctx":07F6
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image io1 
      Height          =   300
      Left            =   2520
      Picture         =   "Tab.ctx":0888
      Top             =   960
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Label l1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab0"
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image i3 
      Height          =   330
      Index           =   0
      Left            =   1440
      Picture         =   "Tab.ctx":096A
      Top             =   1200
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image i1 
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "Tab.ctx":0A5C
      Top             =   1200
      Visible         =   0   'False
      Width           =   30
   End
   Begin VB.Image i2 
      Height          =   315
      Index           =   0
      Left            =   240
      Picture         =   "Tab.ctx":0B4E
      Stretch         =   -1  'True
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Menu PM 
      Caption         =   "PM"
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   1
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   2
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   3
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   4
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   5
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   6
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   7
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   8
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   9
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   10
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   11
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   12
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   13
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   14
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   15
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   16
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   17
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   18
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   19
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   20
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   21
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   22
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   23
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   24
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   25
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   26
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   27
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   28
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   29
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   30
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   31
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   32
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   33
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   34
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   35
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   36
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   37
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   38
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   39
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   40
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   41
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   42
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   43
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   44
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   45
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   46
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   47
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   48
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   49
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   50
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   51
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   52
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   53
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   54
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   55
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   56
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   57
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   58
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   59
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   60
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   61
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   62
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   63
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   64
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   65
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   66
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   67
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   68
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   69
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   70
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   71
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   72
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   73
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   74
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   75
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   76
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   77
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   78
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   79
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   80
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   81
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   82
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   83
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   84
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   85
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   86
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   87
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   88
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   89
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   90
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   91
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   92
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   93
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   94
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   95
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   96
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   97
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   98
      End
      Begin VB.Menu PMTabArray 
         Caption         =   ""
         Index           =   99
      End
   End
End
Attribute VB_Name = "Tab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'[Tab]
Option Explicit

Const TabSpacing = -15
Const Padding = 180

Dim TheLastTabYouWereOver As Integer, TheActiveTab As Integer
Dim MoveAlong As Long 'offset for moving the tabs around when there are too many

Dim myCloseButton As Boolean
Dim myActiveForeColor As OLE_COLOR, myBlurForeColor As OLE_COLOR
Dim TabHeight As Long

Public Event Click(tIndex As Integer)
Public Event DblClick(tIndex As Integer)
Public Event RightClick(tIndex As Integer)
Public Event TabClose(tIndex As Integer)

Private Sub i1_DblClick(Index As Integer)
    RaiseEvent DblClick(Index)
End Sub

Private Sub i2_DblClick(Index As Integer): i1_DblClick Index: End Sub
Private Sub i3_DblClick(Index As Integer): i1_DblClick Index: End Sub

Private Sub i4_Click(Index As Integer)
    i4(Index).Visible = False 'cos the tab is already gone
    RaiseEvent TabClose(Index)
    Redraw
End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    UserControl_MouseDown Button, Shift, x, y
End Sub

Private Sub l1_DblClick(Index As Integer): i1_DblClick Index: End Sub

Private Sub i1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Index = TheLastTabYouWereOver And i1.Count > 2 Then Exit Sub 'if im already on it then do nothing. No flicker now is there?
    
    If TheLastTabYouWereOver = 0 Then TheLastTabYouWereOver = 1
    
    ShowDefaultImg TheLastTabYouWereOver
    ShowHoverImg Index
    TheLastTabYouWereOver = Index
    
    If myCloseButton Then i4(Index).Picture = ioc.Picture

    
End Sub
Private Sub i2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): i1_MouseMove Index, Button, Shift, x, y: End Sub
Private Sub i3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): i1_MouseMove Index, Button, Shift, x, y: End Sub
Private Sub l1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): i1_MouseMove Index, Button, Shift, x, y: End Sub

Private Sub i1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        RaiseEvent RightClick(Index)
    End If
End Sub
Private Sub i2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): i1_MouseDown Index, Button, Shift, x, y: End Sub
Private Sub i3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): i1_MouseDown Index, Button, Shift, x, y: End Sub
Private Sub l1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single): i1_MouseDown Index, Button, Shift, x, y: End Sub


Private Sub i1_Click(Index As Integer)
    On Error Resume Next
    Box.Visible = True
    Box.ZOrder 0 'bring this one to top
    i1(Index).ZOrder 0
    i2(Index).ZOrder 0
    i3(Index).ZOrder 0
    l1(Index).ZOrder 0
    ShowDefaultImg Index
    ShowDefaultImg TheLastTabYouWereOver
    TheActiveTab = Index
    Redraw
    RaiseEvent Click(Index)
End Sub
Private Sub i2_Click(Index As Integer): i1_Click Index: End Sub
Private Sub i3_Click(Index As Integer): i1_Click Index: End Sub
Private Sub l1_Click(Index As Integer): i1_Click Index: End Sub

Private Function ShowHoverImg(Index As Integer)
    On Error Resume Next
    If Index = 0 Then Exit Function 'preventing original from getting overwritten
    If Index = TheActiveTab Then Exit Function
    i1(Index).Picture = io1.Picture
    i2(Index).Picture = io2.Picture
    i3(Index).Picture = io3.Picture
    
    i4(Index).Picture = ioc.Picture
End Function

Private Function ShowDefaultImg(Index As Integer)
    On Error Resume Next
    If Index = 0 Then Exit Function 'preventing original from getting overwritten
    i1(Index).Picture = i1(0).Picture
    i2(Index).Picture = i2(0).Picture
    i3(Index).Picture = i3(0).Picture
    i4(Index).Picture = ioc2.Picture
End Function

Private Function Redraw()
'this is it... the heart of the tab drawing.

    On Error Resume Next
    Dim Ix As Integer
    Dim DT As Long, dy As Long
    Dim J As Long
    
    DT = MoveAlong
    If TabHeight = 0 Then TabHeight = 300
    Box.Move 0, TabHeight, UserControl.Width, UserControl.Height - TabHeight
    Box.ZOrder 0
    
    For Ix = 1 To i1.UBound Step 1
        
        If Ix = TheActiveTab Then dy = 0 Else dy = 15
        
        i1(Ix).Move DT, dy, 30, TabHeight + 15
        If dy = 0 Then i1(Ix).ZOrder 0
        
        DT = DT + i1(Ix).Width
        
        With l1(Ix)
            .FontBold = (Ix = TheActiveTab)
            .ForeColor = IIf(Ix = TheActiveTab, myActiveForeColor, myBlurForeColor)
            .Move DT + Padding \ 2, (TabHeight + 15 + dy - .Height) \ 2, .Width
            J = .Width + Padding
            If myCloseButton Then J = J + i4(Ix).Width
            i2(Ix).Move .Left - TabSpacing - Padding \ 2 + TabSpacing, dy, J, TabHeight + 15
            If dy = 0 Then i2(Ix).ZOrder 0
            .ZOrder 0
            DT = DT + J + TabSpacing
        End With
        
        i3(Ix).Move DT - TabSpacing, dy, 30, TabHeight + 15
        If dy = 0 Then i3(Ix).ZOrder 0
        
        i4(Ix).Visible = myCloseButton
        If myCloseButton Then
            i4(Ix).Move i3(Ix).Left - i4(Ix).Width - IIf(dy = 0, 30, 15), IIf(dy = 0, 60, 75)
            i4(Ix).ZOrder 0
        End If
        
        DT = DT + i3(Ix).Width
    Next
    
    
End Function

Public Property Let ActiveTab(Index As Integer)
    On Error Resume Next
    TheActiveTab = Index
    i1(Index).ZOrder 0
    i2(Index).ZOrder 0
    i3(Index).ZOrder 0
    i4(Index).ZOrder 0
    'i1_Click Index
    Redraw
End Property

Public Property Get ActiveTab() As Integer
    On Error Resume Next
    ActiveTab = TheActiveTab
End Property

Public Property Let TabCaption(Index As Integer, Text As String)
    On Error Resume Next
    l1(Index).Caption = Text
    Redraw
End Property

Public Property Get TabCaption(Index As Integer) As String
    On Error Resume Next
    TabCaption = l1(Index).Caption
End Property

Public Property Let TabTag(Index As Integer, Text As String)
    On Error Resume Next
    l1(Index).Tag = Text
    Redraw
End Property

Public Property Get TabTag(Index As Integer) As String
    On Error Resume Next
    TabTag = l1(Index).Tag
End Property

Public Property Let BackColor(ByVal What As OLE_COLOR)
    On Error Resume Next
    UserControl.BackColor = What
    PropertyChanged "BackColor"
End Property

Public Property Get BackColor() As OLE_COLOR
    On Error Resume Next
    BackColor = UserControl.BackColor
End Property

Public Property Let AllTabsForeColor(ByVal What As OLE_COLOR)
    On Error Resume Next
    lblAllTabs.ForeColor = What
    PropertyChanged "AllTabsForeColor"
End Property

Public Property Get AllTabsForeColor() As OLE_COLOR
    On Error Resume Next
    AllTabsForeColor = lblAllTabs.ForeColor
End Property

Public Property Let TabTooltip(Index As Integer, Text As String)
    On Error Resume Next
    l1(Index).ToolTipText = Text
    Redraw
End Property

Public Property Get TabTooltip(Index As Integer) As String
    On Error Resume Next
    TabTooltip = l1(Index).ToolTipText
End Property

Public Property Let TH(What As Long)
    On Error Resume Next
    TabHeight = What
    Redraw
End Property

Public Property Get TH() As Long
    On Error Resume Next
    TH = TabHeight
End Property

Public Property Let CloseButton(Really As Boolean)
    On Error Resume Next
    myCloseButton = Really
    PropertyChanged "CloseButton"
    Redraw
End Property

Public Property Get CloseButton() As Boolean
    On Error Resume Next
    CloseButton = myCloseButton
End Property

Public Property Let ActiveForeColor(What As OLE_COLOR)
    On Error Resume Next
    myActiveForeColor = What
    PropertyChanged "ActiveForeColor"
End Property

Public Property Get BlurForeColor() As OLE_COLOR
    On Error Resume Next
    BlurForeColor = myBlurForeColor
End Property

Public Property Let BlurForeColor(What As OLE_COLOR)
    On Error Resume Next
    myBlurForeColor = What
    PropertyChanged "BlurForeColor"
End Property

Public Property Get ActiveForeColor() As OLE_COLOR
    On Error Resume Next
    ActiveForeColor = myActiveForeColor
End Property

Public Property Get Picture() As StdPicture
    On Error Resume Next
    Set Picture = imgBack.Picture
End Property

Public Property Set Picture(ByVal newPic As StdPicture)
    On Error Resume Next
    imgBack.Picture = newPic
    PropertyChanged "picture"
End Property

Public Property Get Font() As String
    On Error Resume Next
    Font = l1(0).FontName
End Property

Public Property Let Font(ByVal newFont As String)
    On Error Resume Next
    l1(0).FontName = newFont
    PropertyChanged "FontName"
End Property





Public Function AddTab(Optional Caption As String) As Integer
    On Error Resume Next
    Dim Idx As Integer
    'left of tab
    Load i1(i1.UBound + 1)
    i1(i1.UBound).Visible = True
    i1(i1.UBound).ZOrder 0
    Load i2(i2.UBound + 1)
    i2(i2.UBound).Visible = True
    i2(i2.UBound).ZOrder 0
    Load i3(i3.UBound + 1)
    i3(i3.UBound).Visible = True
    i3(i3.UBound).ZOrder 0
    Load i4(i4.UBound + 1)
        i4(i4.UBound).Visible = myCloseButton
        i4(i4.UBound).ZOrder 0
    Load l1(l1.UBound + 1)
    l1(l1.UBound).Visible = True
    l1(l1.UBound).ZOrder 0
    
    ShowDefaultImg l1.UBound
    
    l1(l1.UBound).Caption = IIf(Caption = "", "Tab " & l1.UBound, Caption)
    Redraw
    
    AddTab = l1.UBound
    
End Function

Public Sub AddTabs(ParamArray Caption() As Variant)
    On Error Resume Next
    Dim i As Integer
    For i = LBound(Caption()) To UBound(Caption()) Step 1
        AddTab CStr(Caption(i))
    Next
End Sub

Public Sub RemoveTab(Index As Integer)
    On Error Resume Next
    i1(Index).Visible = False
    Set i1(Index) = Nothing
    Unload i1(Index)
    i2(Index).Visible = False
    Set i2(Index) = Nothing
    Unload i2(Index)
    i3(Index).Visible = False
    Set i3(Index) = Nothing
    Unload i3(Index)
    
    i4(Index).Visible = False
    Set i4(Index) = Nothing
    Unload i4(Index)
    
    Unload l1(Index)    'Labels get unloaded easily
    Redraw
End Sub

Public Sub RemoveAllTabs()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To i1.UBound Step 1 'static 1 is put there on purpose (0 is base)
        RemoveTab i
    Next
End Sub

Public Function TabUBound() As Integer
    On Error Resume Next
    TabUBound = i1.UBound
End Function

Public Function TabLBound() As Integer
    On Error Resume Next
    TabLBound = i1.LBound
End Function

Private Function IsLoaded(Index As Integer) As Boolean
    On Error Resume Next
    IsLoaded = (i1(Index).Name = i1(Index).Name)
End Function

Private Sub lblAllTabs_Click()
    On Error Resume Next
    Dim i As Long
    Dim K As String
    
    For i = 0 To PMTabArray.UBound Step 1
        With PMTabArray(i)
            K = l1(i + 1).Caption 'leak prevention
            .Caption = K
            .Visible = (Len(.Caption) > 0)
            K = "" 'leak prevention
        End With
    Next
    PopupMenu PM, , lblAllTabs.Left, lblAllTabs.Top
End Sub

Private Sub PMTabArray_Click(Index As Integer)
    On Error Resume Next
    i1_Click Index + 1
End Sub

Private Sub UserControl_Initialize()
    Box.Visible = True
    TabHeight = 300 'why do i have to do this
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 Then
        RaiseEvent RightClick(0) '0 is impossible in tabs, so 0 is the body
    End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        UserControl.BackColor = .ReadProperty("BackColor", &H8000000F)
        lblAllTabs.ForeColor = .ReadProperty("AllTabsForeColor", &H80000007)
        myBlurForeColor = .ReadProperty("BlurForeColor", 0)
        myActiveForeColor = .ReadProperty("ActiveForeColor", 0)
        lblAllTabs.BackColor = UserControl.BackColor
        myCloseButton = .ReadProperty("CloseButton", False)
        Set imgBack.Picture = .ReadProperty("picture", Nothing)
        TabHeight = .ReadProperty("TH", 300)
        l1(0).FontName = .ReadProperty("FontName", "MS Shell Dlg")
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Box.Move 0, TabHeight, UserControl.Width, UserControl.Height - TabHeight
    lblAllTabs.Move Width - lblAllTabs.Width, (Box.Top - lblAllTabs.Height) / 2
    imgBack.Move 0, 0, Width, TabHeight
    lblAllTabs.ZOrder 0
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "BackColor", UserControl.BackColor
        .WriteProperty "CloseButton", myCloseButton
        .WriteProperty "BlurForeColor", myBlurForeColor
        .WriteProperty "ActiveForeColor", myActiveForeColor
        .WriteProperty "picture", imgBack.Picture
        .WriteProperty "TH", TabHeight, 300
        .WriteProperty "AllTabsForeColor", lblAllTabs.ForeColor
        .WriteProperty "FontName", l1(0).FontName, "MS Shell Dlg"
    End With
End Sub

