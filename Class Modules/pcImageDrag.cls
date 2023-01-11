VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcImageDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==================================================================================================
'pcImageDrag.cls      4/20/05
'
'           PURPOSE:
'               Manage an image drag operation using the comctl32 ImageList_Drag* functions.
'
'           LINEAGE:
'               LVDrag.vbp by Brad Martinez http://www.mvps.org
'
'==================================================================================================

Option Explicit



Private mhIml      As Long
Private mhWndLock  As Long

Private Sub Class_Terminate()
    pDestroy
End Sub


Private Sub pMove(ByVal x As Long, ByVal y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Move the image drag to the new location.
    '---------------------------------------------------------------------------------------
    Dim lHwnd      As Long
    
    lHwnd = RootParent(WindowFromPoint(x, y))
    'lhWnd = WindowFromPoint(x, y)
    
    Dim tR      As RECT
    If GetWindowRect(lHwnd, tR) Then
        x = x - tR.Left
        y = y - tR.Top
    End If
    
    If lHwnd <> mhWndLock Then
        ImageList_DragLeave mhWndLock
        ImageList_DragEnter lHwnd, x, y
        mhWndLock = lHwnd
    End If
    
    ImageList_DragMove x, y
    
End Sub

Private Sub pDestroy()
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : End the image drag operation.
    '---------------------------------------------------------------------------------------
    If mhIml Then
        ImageList_DragLeave mhWndLock
        ImageList_EndDrag
        ImageList_Destroy mhIml
   
       ' Hook_Remove Me, WH_MOUSE
    End If
End Sub

Public Sub StartDrag(ByVal oDib As pcDibSection, ByVal x As Long, ByVal y As Long)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Begin a new image drag operation.
    '---------------------------------------------------------------------------------------
    pDestroy
    
    mhIml = ImageList_Create(oDib.Width, oDib.Height, ILC_MASK, OneL, ZeroL)
    
    If mhIml Then
        
        ImageList_AddMasked mhIml, oDib.hBitmap, ImageDrag_TransColor
        ImageList_BeginDrag mhIml, ZeroL, x, y
       ' Hook_Install Me, WH_MOUSE
        
    End If
    
End Sub

Public Sub StopDrag()
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : End the image drag operation.
    '---------------------------------------------------------------------------------------
    pDestroy
End Sub

Public Sub Show(ByVal bVal As Boolean)
    '---------------------------------------------------------------------------------------
    ' Date      : 4/20/05
    ' Purpose   : Show or hide the image.
    '---------------------------------------------------------------------------------------
    ImageList_DragShowNolock -CLng(bVal)
End Sub
