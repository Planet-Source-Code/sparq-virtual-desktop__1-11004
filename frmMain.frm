VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3795
      Left            =   9180
      TabIndex        =   5
      Top             =   0
      Width           =   2370
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1860
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Background Color"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   75
         TabIndex        =   6
         Top             =   120
         Width           =   2205
      End
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   10680
      TabIndex        =   4
      Top             =   8640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   10680
      TabIndex        =   3
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ImageList IMGLarge 
      Left            =   9540
      Top             =   8880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList IMGSmall 
      Left            =   10140
      Top             =   8880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4035
      Left            =   180
      TabIndex        =   2
      Top             =   1740
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7117
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox PixSmall 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   250
      Left            =   10200
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   8280
      Visible         =   0   'False
      Width           =   250
   End
   Begin VB.PictureBox pixLarge 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9540
      ScaleHeight     =   495
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   8280
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Menu mnuPopup 
      Caption         =   ""
      Begin VB.Menu mnuGhost 
         Caption         =   "(un)&Ghost Item"
         Shortcut        =   ^G
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Dim ItemCount As Integer
Dim WinDir As String

Private Sub LoadDesktop()
    ListView1.ListItems.Clear
    ListView1.View = 2
    Dim imgX As ListImage
    Dim FileName As String
    Dim Spot As Integer
    Dim SpotB As Integer
    Dim x As Integer
        
        Dir1.Path = WinDir & "desktop\"
        File1.Path = Dir1.Path
        
        For x = 0 To Dir1.ListCount - 1 '-------LOAD FOLDERS
            DoEvents
            ItemCount = ItemCount + 1
            FileName = Dir1.List(x)
            DisplayIcons (FileName)
            Set imgX = IMGSmall.ListImages.Add(ItemCount, , PixSmall.Picture)
            Set imgX = IMGLarge.ListImages.Add(ItemCount, , pixLarge.Picture)
            
            ListView1.Icons = IMGLarge
            ListView1.SmallIcons = IMGSmall
            
            Spot = 1
            Do Until Spot = 0
                SpotB = Spot
                Spot = InStr(SpotB + 1, FileName, "\")
                DoEvents
            Loop
            'SpotB = Spot
            
            Dim ItemX As ListItem
            Set ItemX = ListView1.ListItems.Add()
            ItemX.Text = Right$(FileName, Len(FileName) - SpotB)
            ItemX.Icon = ItemCount
            ItemX.SmallIcon = ItemCount
            ItemX.Tag = ItemCount
        Next x

        For x = 0 To File1.ListCount - 1 '-------LOAD FILES
            DoEvents
            ItemCount = ItemCount + 1
            FileName = Dir1.Path & "\" & File1.List(x)
            DisplayIcons (FileName)
            Set imgX = IMGSmall.ListImages.Add(ItemCount, , PixSmall.Picture)
            Set imgX = IMGLarge.ListImages.Add(ItemCount, , pixLarge.Picture)
            
            ListView1.Icons = IMGLarge
            ListView1.SmallIcons = IMGSmall
            
            Spot = 1
            Do Until Spot = 0
                SpotB = Spot
                Spot = InStr(SpotB + 1, FileName, "\")
                DoEvents
            Loop
            'SpotB = Spot
            
            Set ItemX = ListView1.ListItems.Add()
            ItemX.Text = Right$(FileName, Len(FileName) - SpotB)
            ItemX.Icon = ItemCount
            ItemX.SmallIcon = ItemCount
            ItemX.Tag = ItemCount
        Next x
        Form_Resize
        Visible = True
End Sub

Function DisplayIcons(Fname As String) As Long
    Dim hImgSmall As Long
    Dim hImgLarge As Long
    Dim info1 As String
    Dim info2 As String
    On Local Error GoTo cmdLoadErrorHandler

        hImgSmall = SHGetFileInfo(Fname$, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        info1 = Left$(shinfo.szDisplayName, InStr(shinfo.szDisplayName, Chr$(0)) - 1)
        info2 = Left$(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr$(0)) - 1)
        PixSmall.Picture = LoadPicture()
        PixSmall.AutoRedraw = True
        ImageList_Draw hImgSmall&, shinfo.iIcon, PixSmall.hDC, 0, 0, ILD_TRANSPARENT
        PixSmall.Picture = PixSmall.Image
       
        hImgLarge = SHGetFileInfo(Fname, 0&, shinfo, Len(shinfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
        info1 = Left$(shinfo.szDisplayName, InStr(shinfo.szDisplayName, Chr$(0)) - 1)
        info2 = Left$(shinfo.szTypeName, InStr(shinfo.szTypeName, Chr$(0)) - 1)
        pixLarge.Picture = LoadPicture()
        pixLarge.AutoRedraw = True
        ImageList_Draw hImgLarge&, shinfo.iIcon, pixLarge.hDC, 0, 0, ILD_TRANSPARENT
        pixLarge.Picture = pixLarge.Image
        Exit Function
        
cmdLoadErrorHandler:
  pixLarge.Picture = LoadPicture()

End Function

Private Sub Form_Load()
  Dim buf As String * 256
  Dim return_len As Long
  Dim wid1 As Single
  
  return_len = GetWindowsDirectory(buf, Len(buf))
  WinDir = Left$(buf, return_len) & "\"
  LoadDesktop
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With ListView1
        .Top = 30
        .Left = 120
        .Width = Width - 2700
        .Height = Height - 760
    End With
    With Frame1
        .Top = 30
        .Left = Width - (.Width)
        .Width = 2500
        .Height = Height - 760
    End With
End Sub

Private Sub Label1_Click()
    On Error Resume Next
    With CommonDialog1
        .CancelError = True
        .ShowColor
        Label1.BackColor = .Color
        ListView1.BackColor = .Color
        PixSmall.BackColor = .Color
        pixLarge.BackColor = .Color
    End With
    LoadDesktop
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuGhost_Click()
    ListView1.ListItems(Val(ListView1.SelectedItem.Tag)).Ghosted = Not (ListView1.ListItems(Val(ListView1.SelectedItem.Tag)).Ghosted)
End Sub
