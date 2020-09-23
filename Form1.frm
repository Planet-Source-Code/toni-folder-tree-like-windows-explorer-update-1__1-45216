VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   11310
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   7830
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14314
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Bevel           =   0
            Picture         =   "Form1.frx":0000
            TextSave        =   "6:06 ìì"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6855
      Left            =   3360
      TabIndex        =   3
      Top             =   360
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   12091
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "file:///D:/My%20VB%20Programs/Uploads%20For%20Planet%20Source/Tree%20To%20Explorer/Step%202"
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Icon Size To 32"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
   Begin VB.PictureBox Pi1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   13996
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList Im16 
      Left            =   120
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label Spliter 
      Height          =   2055
      Left            =   3120
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1680
      Picture         =   "Form1.frx":01DC
      Top             =   7440
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   960
      Picture         =   "Form1.frx":0766
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim m_dwLrgIconWidth As Long
Dim m_dwLrgIconHeight As Long
Dim m_dwSmIconWidth As Long
Dim m_dwSmIconHeight As Long

Public Enum IcSize
            Icons16
            Icons32
End Enum
Dim m_IconsSize As IcSize
Dim StartNod As Node
Dim SplEnab As Boolean
Dim X1, Y1, SpL

Private Sub Command1_Click()
If m_IconsSize = Icons32 Then
    Command1.Caption = "Change Icon Size To 32"
    m_IconsSize = Icons16
Else
    Command1.Caption = "Change Icon Size To 16"
    m_IconsSize = Icons32
End If
ChangeIconSize
FillDrivers
End Sub

Private Sub Form_Load()

ChangeIconSize
FillDrivers

End Sub
Private Sub ChangeIconSize()
Tree.Nodes.Clear
Im16.ListImages.Clear

If m_IconsSize = Icons16 Then
    Im16.ImageHeight = 16
    Im16.ImageWidth = 16
    Pi1.Width = 240
    Pi1.Height = 240
    Set Pi1.Picture = Image2.Picture
Else
    Im16.ImageHeight = 32
    Im16.ImageWidth = 32
    Pi1.Width = 480
    Pi1.Height = 480
    Set Pi1.Picture = Image1.Picture
End If

Im16.ListImages.Add 1, "Start Icon", Pi1.Picture
Im16.ListImages.Add 2, "Start Icon1", Pi1.Picture
Set Tree.ImageList = Im16
Tree.Indentation = 19 * Screen.TwipsPerPixelX
Tree.ImageList = Im16
Tree.Refresh

End Sub
Private Sub FillDrivers()
Dim Nod As Node, Nod1 As Node, DrvNod As Node
Dim DRV As String
Dim Index

Set Nod = Tree.Nodes.Add(, , , "My Computer", "Start Icon", "Start Icon")
Nod.Expanded = True
    If DriveExist("A:\") = True Then
        GetIcon ("A:\"), False, Index1
        Set DrvNod = Tree.Nodes.Add(Nod.Index, tvwChild, , MakeNames("3.5 Floppy (A:)"), Index1, Index1)
        DrvNod.Sorted = True
        Tree.Nodes.Add DrvNod.Index, tvwChild, , "Dummy"
    End If
        
    If DriveExist("B:\") = True Then
        GetIcon ("B:\"), False, Index1
        DrvNod.Sorted = True
        Set DrvNod = Tree.Nodes.Add(Nod.Index, tvwChild, , MakeNames("3.5 Floppy (B:)"), Index1, Index1)
        Tree.Nodes.Add DrvNod.Index, tvwChild, , "Dummy"
    End If

For X = 67 To 85
 DRV = Chr(X) + ":"
    If DriveExist(DRV) Then
        GetIcon DRV + "\", False, Index
        Set Nod1 = Tree.Nodes.Add(Nod.Index, tvwChild, , MakeNames(GetDriveName(DRV) + " (" + DRV + ")"), Index, Index)
        Nod1.Sorted = True
        If HasFolderInside(DRV) = True Then
            Tree.Nodes.Add Nod1.Index, tvwChild, , "Dummy"
        End If
        If X = 67 Then
            Set StartNod = Nod1
            StartNod.Expanded = True
            StartNod.Selected = True
            Tree_NodeClick StartNod
        End If
    End If
Next

End Sub

Public Function HasFolderInside(Path As String) As Boolean
On Error GoTo HASFErr
Dim Path1 As String
Path1 = Path
    Dim Fso, F, F1, FC, M, ArtH, Att, HasF As Boolean
    If Len(Path1) < 3 Then Path1 = Path1 + "\"
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set F = Fso.GetFolder(Path1)
    Set FC = F.SubFolders
    If FC.Count > 0 Then
        HasFolderInside = True
    Else
        HasFolderInside = False
    End If
Exit Function
HASFErr:
        HasFolderInside = False
End Function

Private Sub Form_Resize()
Tree.Move 0, 0, Tree.Width, Me.ScaleHeight - StatusBar1.Height
Command1.Move Tree.Width, 0, Me.ScaleWidth - Tree.Width
Spliter.Move Tree.Width, 0, 60, Me.ScaleHeight - StatusBar1.Height
StatusBar1.Panels.Item(1).Width = Spliter.Left
WebBrowser1.Move Spliter.Left + Spliter.Width, Command1.Height, Me.ScaleWidth - (Spliter.Left + Spliter.Width), Me.ScaleHeight - Command1.Height - StatusBar1.Height
End Sub

Public Function DriveExist(DRV) As Boolean
    Dim Fso, msg
    Set Fso = CreateObject("Scripting.FileSystemObject")
    DriveExist = Fso.DriveExists(DRV)
End Function

Public Function FolderExists(FolderName) As Boolean
    Dim Fso, msg
    Set Fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = Fso.FolderExists(FolderName)
End Function

Public Function GetDriveName(DRV)
On Error GoTo NotDrive
    Dim Fso, D, s
    Set Fso = CreateObject("Scripting.FileSystemObject")
    Set D = Fso.GetDrive(Fso.GetDriveName(DRV))
    If D.DriveType = "Remote" Then
        s = D.ShareName
    ElseIf D.IsReady Then
      s = D.VolumeName
    End If
    GetDriveName = s
Exit Function
NotDrive:
GetDriveName = ""
End Function


Private Function CheckImageKey(Key As String, Optional Index As Long) As Boolean
Dim D As Long
For D = 1 To Im16.ListImages.Count
    If LCase(Key) = LCase(Im16.ListImages(D).Key) Then
        Index = D
        CheckImageKey = True
        Exit Function
    End If
Next
Index = 0
CheckImageKey = False
End Function

Private Sub GetIcon(sFilePath As String, OpenIcon As Boolean, Index)
Dim Inde As Long, SysIn As Long, VBart As VbFileAttribute
Dim charc As String
charc = "Icon"

Index = 0
If OpenIcon = True Then
    If m_IconsSize = Icons16 Then
        GetIconSize sFilePath, m_dwSmIconWidth, m_dwSmIconHeight, SHGFI_SMALLICON Or SHGFI_OPENICON, Inde
    Else
        GetIconSize sFilePath, m_dwLrgIconWidth, m_dwLrgIconHeight, SHGFI_OPENICON, Inde
    End If
    If CheckImageKey(charc + Str(Inde), SysIn) = True Then
        Index = SysIn
    Else
        Set Pi1.Picture = Nothing
        If m_IconsSize = Icons16 Then
            ShowFileIcon sFilePath, SHGFI_SMALLICON, Pi1, True
        Else
            ShowFileIcon sFilePath, SHGFI_LARGEICON, Pi1, True
        End If
        
        Set Pi1.Picture = Pi1.Image
        
        Im16.ListImages.Add Im16.ListImages.Count, charc + Str(Inde), Pi1.Picture
        Index = Im16.ListImages.Count - 1
        
    End If
Else
    If m_IconsSize = Icons16 Then
        GetIconSize sFilePath, m_dwSmIconWidth, m_dwSmIconHeight, SHGFI_SMALLICON, Inde
    Else
        GetIconSize sFilePath, m_dwLrgIconWidth, m_dwLrgIconHeight, SHGFI_LARGEICON, Inde
    End If
    If CheckImageKey(charc + Str(Inde), SysIn) = True Then
        Index = SysIn
    Else
        Set Pi1.Picture = Nothing
        If m_IconsSize = Icons16 Then
            ShowFileIcon sFilePath, SHGFI_SMALLICON, Pi1, False
        Else
            ShowFileIcon sFilePath, SHGFI_LARGEICON, Pi1, False
        End If
        
        Set Pi1.Picture = Pi1.Image
        
        Im16.ListImages.Add Im16.ListImages.Count, charc + Str(Inde), Pi1.Picture
        
        Index = Im16.ListImages.Count - 1
    End If
End If

End Sub

Private Sub GetIconSize(sFilePath As String, _
                                    dwWidth As Long, _
                                    dwHeight As Long, _
                                    dwFlags As Long, SysIndex As Long)
  Dim shfi As SHFILEINFO, hSysImgLst As Long
  
  hSysImgLst = SHGetFileInfo(ByVal sFilePath, 0&, shfi, Len(shfi), _
                                            SHGFI_SYSICONINDEX Or dwFlags)
   SysIndex = shfi.iIcon
  ImageList_GetIconSize hSysImgLst, dwWidth, dwHeight
End Sub

Private Sub ShowFileIcon(sFilePath As String, _
                                      uFlags As Long, _
                                      objPB As PictureBox, OpenIcon As Boolean)
  Dim shfi As SHFILEINFO
  
  objPB.Cls   ' clear prev icon
  If OpenIcon = True Then
    SHGetFileInfo ByVal sFilePath, 0&, shfi, Len(shfi), SHGFI_ICON Or SHGFI_OPENICON Or uFlags
  Else
    SHGetFileInfo ByVal sFilePath, 0&, shfi, Len(shfi), SHGFI_ICON Or uFlags
  End If
  ' DrawIconEx() will shrink (or stretch) the
  ' icon per it's cxWidth & cyWidth params
  If uFlags And SHGFI_SMALLICON Then
    DrawIconEx objPB.hDC, 0, 0, shfi.Hicon, _
                      m_dwSmIconWidth, m_dwSmIconHeight, 0, 0, DI_NORMAL
  Else
    DrawIconEx objPB.hDC, 0, 0, shfi.Hicon, _
                      m_dwLrgIconWidth, m_dwLrgIconHeight, 0, 0, DI_NORMAL
  End If
  objPB.Refresh
  
  ' Clean up! -16x16 icons = 380 bytes, 32x32 icons = 1184 bytes
  DestroyIcon shfi.Hicon
End Sub

Private Sub Spliter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SplEnab = True
X1 = X
Y1 = Y
SpL = Spliter.Left
End Sub

Private Sub Spliter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If SplEnab = True Then
    Spliter.Left = SpL + (X - X1)
    SpL = Spliter.Left
    Tree.Width = Spliter.Left
    StatusBar1.Panels.Item(1).Width = Spliter.Left
    Command1.Move Spliter.Left + Spliter.Width, 0, Me.ScaleWidth - (Spliter.Left + Spliter.Width)
    WebBrowser1.Move Spliter.Left + Spliter.Width, Command1.Height, Me.ScaleWidth - (Spliter.Left + Spliter.Width)
End If
End Sub

Private Sub Spliter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SplEnab = False
End Sub

Private Sub Tree_Collapse(ByVal Node As MSComctlLib.Node)
Dim NodPath As String
NodPath = GetPath(Tree.SelectedItem)
If FolderExists(NodPath) = True Then
    StatusBar1.Panels.Item(1).Text = NodPath
    WebBrowser1.Navigate NodPath
End If
End Sub

Private Sub Tree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim NodPath As String
NodPath = GetPath(Node)
If FolderExists(NodPath) = True Then
    StatusBar1.Panels.Item(1).Text = NodPath
    WebBrowser1.Navigate NodPath
End If
End Sub

Private Sub Tree_Expand(ByVal Node As MSComctlLib.Node)
Dim NodPath As String, Index1, Index2
Dim Fso, F, F1, FC
Dim Nod As Node
'Debug.Print Node.FullPath, Node.Child.Text
NodPath = GetPath(Node)
If Node.Child.Text = "Dummy" Then
    Tree.Nodes.Remove (Node.Child.Index)
    
        Set Fso = CreateObject("Scripting.FileSystemObject")
        Set F = Fso.GetFolder(NodPath + "\")
        Set FC = F.SubFolders
        For Each F1 In FC
            GetIcon (NodPath + "\" + F1.Name), False, Index1 ' Close Icon
            GetIcon (NodPath + "\" + F1.Name), True, Index2 'Open Icon
            Set Nod = Tree.Nodes.Add(Node.Index, tvwChild, , MakeNames(F1.Name), Index1, Index2)
            Nod.Sorted = True
            
            If HasFolderInside(NodPath + "\" + F1.Name) = True Then
                If UCase(F1.Name) <> "RECYCLED" Then ' Only English Versions
                    Tree.Nodes.Add Nod.Index, tvwChild, , "Dummy"
                End If
            End If
        Next
End If
End Sub

Private Function GetPath(Nod As Node) As String
On Error GoTo NotPath
    GetPath = Mid(Nod.FullPath, InStr(1, Nod.FullPath, ":", vbTextCompare) - 1, Len(Nod.FullPath))
    If Mid(GetPath, 3, 1) = ")" Then
        GetPath = Left(GetPath, 2) + Mid(GetPath, 4, Len(GetPath))
    End If
Exit Function
NotPath:
GetPath = ""
End Function

Private Function MakeNames(Tex As String) As String
Dim X, t
    t = UCase(Left(Tex, 1)) + LCase(Right(Tex, Len(Tex) - 1))
                
    For X = 1 To Len(t)
        If Mid(t, X, 1) = " " Or Mid(t, X, 1) = "-" Or Mid(t, X, 1) = "_" Or Mid(t, X, 1) = "," _
            Or Mid(t, X, 1) = "." Or Mid(t, X, 1) = "(" Or Mid(t, X, 1) = ")" Or Mid(t, X, 1) = "[" _
            Or Mid(t, X, 1) = "]" Or Mid(t, X, 1) = "\" Or Mid(t, X, 1) = "/" Or Mid(t, X, 1) = "=" _
            Or Mid(t, X, 1) = "+" Then
                t = Left(t, X) + UCase(Mid(t, X + 1, 1)) + Mid(t, X + 2, Len(t))
        End If
    Next
    MakeNames = t
End Function
