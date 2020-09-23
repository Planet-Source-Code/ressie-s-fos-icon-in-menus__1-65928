VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ressie S. Fos"
   ClientHeight    =   3975
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   1005
      Left            =   1185
      ScaleHeight     =   945
      ScaleWidth      =   1740
      TabIndex        =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   1800
      Begin VB.Image imgOpen 
         Height          =   210
         Left            =   300
         Picture         =   "Form1.frx":0000
         Top             =   75
         Width           =   210
      End
      Begin VB.Image imgNew 
         Height          =   210
         Left            =   105
         Picture         =   "Form1.frx":02AA
         Top             =   105
         Width           =   210
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*******************************************
' The following are API declarations and   '
' constants that set menu images or bitmaps'
'*******************************************

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long        ':( Missing Scope
Private Const MF_BYPOSITION = &H400&
Private mHandle As Long
Private lRet As Long
Private sHandle As Long
'*******************************************

Private Sub PaintMenuBitmaps()  'function to set bitmaps to menus
    On Error Resume Next
    AssignMenuBitmaps Me, imgNew, 0, 0 'New
    AssignMenuBitmaps Me, imgOpen, 0, 1 'Open
End Sub

Private Sub Form_Load()
    PaintMenuBitmaps    'calls the function to set menu bitmaps
End Sub

'Function that assign bitmaps to menu
Private Sub AssignMenuBitmaps(ByRef frm As Form, ByRef IMG As Image, ByVal Menu_Position As Integer, ByVal Sub_Menu_Position As Integer)
   mHandle = GetMenu(frm.hWnd)
   sHandle = GetSubMenu(mHandle, Menu_Position)
   lRet = SetMenuItemBitmaps(sHandle, Sub_Menu_Position, MF_BYPOSITION, IMG.Picture, IMG.Picture)
  
End Sub
