VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmImageForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DFS Imager"
   ClientHeight    =   7305
   ClientLeft      =   3285
   ClientTop       =   1005
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmImageForm.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ComboDiskList 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Text            =   "Disk #"
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.OptionButton OptionSide1 
      BackColor       =   &H00C0C000&
      Caption         =   "Side 1"
      Height          =   255
      Left            =   4320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   735
   End
   Begin VB.OptionButton OptionSide0 
      BackColor       =   &H00C0C000&
      Caption         =   "Side 0"
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdUnlockFile 
      Caption         =   "Unlock"
      Height          =   315
      Left            =   3840
      TabIndex        =   9
      Top             =   1560
      Width           =   780
   End
   Begin VB.CommandButton cmdLockFile 
      Caption         =   "Lock"
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Top             =   1560
      Width           =   780
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   6930
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListIcons 
      Left            =   1440
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":030A
            Key             =   "UnformattedDisk"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":0626
            Key             =   "BasicFileLocked"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":0942
            Key             =   "FileLocked"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":0C5E
            Key             =   "File"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":0F7A
            Key             =   "LockedDisk"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":1296
            Key             =   "Disk"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":15B2
            Key             =   "BasicFile"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox DiskTitle 
      Height          =   315
      Left            =   240
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   3
      Top             =   1080
      Width           =   1665
   End
   Begin VB.ComboBox ComboOption 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Select All"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   900
   End
   Begin VB.CommandButton cmdCompact 
      Caption         =   "Compact"
      Height          =   315
      Left            =   2040
      TabIndex        =   7
      Top             =   1560
      Width           =   900
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   1560
      Width           =   780
   End
   Begin MSComctlLib.ImageList ImageListSmallIcons 
      Left            =   1440
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":18CE
            Key             =   "UnformattedDisk"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":1BEA
            Key             =   "Disk"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":1F06
            Key             =   "LockedDisk"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":2222
            Key             =   "BasicFile"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":253E
            Key             =   "BasicFileLocked"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":285A
            Key             =   "File"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImageForm.frx":2B76
            Key             =   "FileLocked"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListViewFiles 
      Height          =   4815
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   8493
      View            =   1
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageListIcons"
      SmallIcons      =   "ImageListSmallIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Filename"
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   "LoadAddr"
         Text            =   "Load"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   "ExecAddr"
         Text            =   "Exec"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   "Length"
         Text            =   "Length"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Sector"
         Text            =   "Sector"
         Object.Width           =   1235
      EndProperty
   End
   Begin VB.Image DiskUnlocked 
      Height          =   480
      Left            =   240
      Picture         =   "frmImageForm.frx":2E92
      Stretch         =   -1  'True
      Top             =   200
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image DiskLocked 
      Height          =   480
      Left            =   3000
      Picture         =   "frmImageForm.frx":319C
      Stretch         =   -1  'True
      Top             =   195
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label LabelCycleNo 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cycle"
      Height          =   315
      Left            =   1950
      TabIndex        =   15
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Title:"
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cycle No:"
      Height          =   315
      Left            =   1965
      TabIndex        =   12
      Top             =   840
      Width           =   840
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Boot Option:"
      Height          =   315
      Left            =   2790
      TabIndex        =   11
      Top             =   840
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   6135
      Left            =   120
      Top             =   720
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   5175
   End
   Begin VB.Menu FileMenu 
      Caption         =   "&File"
      Begin VB.Menu OpenImage 
         Caption         =   "&Open Image"
      End
      Begin VB.Menu NewSDSImage 
         Caption         =   "&New Single Sided Image"
      End
      Begin VB.Menu NewDSDImage 
         Caption         =   "New &Double Sided Image"
      End
      Begin VB.Menu SplitDSD 
         Caption         =   "&Split Double Sided Disk"
      End
      Begin VB.Menu CombineSSD 
         Caption         =   "Combine &Two Single Sided Disks"
      End
      Begin VB.Menu CloseImage 
         Caption         =   "&Close Image"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu DiskMenu 
      Caption         =   "&Disk"
      Begin VB.Menu cmdRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu DiskSmallIcons 
         Caption         =   "&Small Icons"
         Checked         =   -1  'True
      End
      Begin VB.Menu DiskLargeIcons 
         Caption         =   "&Large Icons"
         Checked         =   -1  'True
      End
      Begin VB.Menu DiskInfo 
         Caption         =   "&Info"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu ProtectionMenu 
      Caption         =   "&Protection"
      Begin VB.Menu ProtectionDisabled 
         Caption         =   "&Disabled"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu HelpMenu 
      Caption         =   "&Help"
      Begin VB.Menu AboutMe 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmImageForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Image Form
' Created/Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

Private mDisk As disk_type
Private mDiskTable As disktable_type

Private mintSIK As Integer ' Selected Item Key

Private Const FileAct_Delete = 1
Private Const FileAct_Lock = 2
Private Const FileAct_Unlock = 3

Private m_iButton As Integer ' BUG Workaround - MS Article ID : 240946

Private Sub SaveCat(boolTitle As Boolean)
    ' Save catalogue
    If mDisk.ImageFile <> "" Then
        DiskCatalogue_Save mDisk, mDiskTable, boolTitle
        LabelCycleNo.Caption = mDisk.Catalogue.CycleNo
    End If
End Sub

Private Sub AboutMe_Click()
    ' Show 'About' form
    frmAbout.Visible = True
End Sub

Private Sub CloseDisk(boolClosingForm As Boolean)
    ' Close Disk Image
    If mDisk.Open Then
        ComboOption_LostFocus
        DiskTitle_LostFocus
        Disk_Close mDisk
    End If

    If Not boolClosingForm Then
        DiskSmallIcons_Click
        OpenDisk "", 0
    End If
End Sub

Private Sub CloseImage_Click()
    ' Close disk
    CloseDisk False
End Sub

Private Sub cmdCompact_Click()
    ' Disk: Compact
    DiskCatalogue_Refresh mDisk
    CompactDisk mDisk, mDiskTable
    
    ' Make sure changes written
    ComboOption_LostFocus
    DiskTitle_LostFocus
    OpenDisk "", 0
End Sub

Private Sub cmdLockFile_Click()
    ' Files: Lock
    SelFilesAction FileAct_Lock
End Sub

Private Sub cmdRefresh_Click()
    ' Refresh disk
    Dim strFile As String
    Dim intDiskNo As Integer
    Dim bytSide As Byte
    
    If mDisk.Open Then
        
        With mDisk
            strFile = .ImageFile
            intDiskNo = .DiskNo
            bytSide = .Side
        End With
        
        CloseDisk True
        OpenDisk strFile, bytSide, intDiskNo
    Else
        OpenDisk "", 0
    End If
End Sub

Private Sub cmdUnlockFile_Click()
    ' Files: Lock
    SelFilesAction FileAct_Unlock
End Sub

Private Sub CombineSSD_Click()
    ' Combine two .SSD's in to a .DSD file
    Dim varFile As Variant
    Dim strFiles(1 To 2) As String
    Dim strTgt As String
    Dim x As Integer
    
    For x = 1 To 2
        ' Get source names
        varFile = glrCommonFileOpenSave( _
                Filter:=glrAddFilterItem("", _
                "DFS Images", "*.ssd;*.img"), _
                Hwnd:=Me.Hwnd, _
                DialogTitle:="Select Single Sided Disk #" & x)
        If IsNull(varFile) Then
            Exit For
        Else
            strFiles(x) = varFile
        End If
    Next
    
    If strFiles(1) <> "" And strFiles(2) <> "" Then
        ' Get target name
        varFile = glrCommonFileOpenSave( _
            Flags:=glrOFN_OVERWRITEPROMPT, _
            Filter:=glrAddFilterItem("", _
            "DFS Images", "*.dsd"), _
            Hwnd:=Me.Hwnd, OpenFile:=False, _
            DialogTitle:="Save As")
            
        If Not IsNull(varFile) Then
            strTgt = varFile
            OpenDisk CombineSSDs(strFiles, strTgt), 0
        End If
    End If
End Sub

Private Sub ComboDiskList_Click()
    ' Change disk (MMB image)
    Dim intDiskNo As Integer
    Dim x As Integer
    Dim n As String
        
    n = ComboDiskList
    If n = "" Then
        xBox "No disk!"
        OpenDisk mDisk.ImageFile, 0, -1
    Else
        x = InStr(n, ":")
        intDiskNo = CInt(Left(n, x - 1))
        If intDiskNo <> mDisk.DiskNo Then
            OpenDisk mDisk.ImageFile, 0, intDiskNo
        End If
    End If
End Sub

Private Sub ComboDiskList_KeyPress(KeyAscii As Integer)
    ' Disk List: Prevent typing, must use list
    KeyAscii = 0
End Sub

Private Sub ComboOption_KeyPress(KeyAscii As Integer)
    ' Boot Option: Prevent typing, must use list
    KeyAscii = 0
End Sub

Private Sub ComboOption_LostFocus()
    ' Boot Option changed?
    Dim o As Byte
    o = BootOptNo(ComboOption.Text)
    With mDisk.Catalogue
        If .Option <> o Then
            .Option = o
            SaveCat False
        End If
    End With
End Sub

Private Sub DiskInfo_Click()
    ' Files: Info
    DiskSmallIcons.Checked = False
    DiskLargeIcons.Checked = False
    DiskInfo.Checked = True
    ListViewFiles.View = lvwSmallIcon Or lvwReport
End Sub

Private Sub DiskLargeIcons_Click()
    ' Files: Large icons
    DiskSmallIcons.Checked = False
    DiskLargeIcons.Checked = True
    DiskInfo.Checked = False
    ListViewFiles.View = lvwIcon
End Sub

Private Sub DiskSmallIcons_Click()
    ' Files: Small Icons
    DiskSmallIcons.Checked = True
    DiskLargeIcons.Checked = False
    DiskInfo.Checked = False
    ListViewFiles.View = lvwSmallIcon
End Sub

Private Sub DiskTitle_KeyPress(KeyAscii As Integer)
    ' DiskTitle: Input mask
    Dim l As Integer
    l = Len(DiskTitle.Text)
    If KeyAscii <> 8 And ((l = 0 And KeyAscii = 32) Or l = 11 _
                    Or KeyAscii < 32) Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        DiskTitle_LostFocus
    End If
End Sub

Private Sub DiskTitle_LostFocus()
    ' DiskTitle: Changed
    If DiskTitle.Text <> mDisk.Catalogue.DiskTitle Then
        mDisk.Catalogue.DiskTitle = DiskTitle.Text
        SaveCat True
    End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
    ' Close About form
    Unload frmAbout
End Sub

Private Sub Form_Load()
    ' Initialise form
    Dim x As Integer
    Dim strOpenFile As String
    Dim intDiskNo As Integer
    
    strOpenFile = Trim(Command()) ' Command line argument
    
    'Debug.Print "Command: "; strOpenFile
    
    If strOpenFile <> "" Then
        ' Strip marks
        If Left(strOpenFile, 1) = """" Then
            strOpenFile = Mid(strOpenFile, 2)
        End If
        If Right(strOpenFile, 1) = """" Then
            strOpenFile = Left(strOpenFile, Len(strOpenFile) - 1)
        End If
    End If
    
    ' Get disk no.
    If strOpenFile <> "" Then
        x = InStr(strOpenFile, "#")
        If x > 0 Then
            intDiskNo = Mid(strOpenFile, x + 1)
            strOpenFile = Left(strOpenFile, x - 1)
        Else
            intDiskNo = -1
        End If
    End If
    
    Debug.Print "Open: "; strOpenFile, intDiskNo
    
    DiskLocked.Left = DiskUnlocked.Left
    
    DiskInfo_Click
    For x = 0 To 3
        ComboOption.AddItem BootOpt(CByte(x)), x
    Next
   
    OpenDisk strOpenFile, 0, intDiskNo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Make sure changes written
    CloseDisk True
End Sub

Private Sub ListViewFiles_Click()
    ' Files selected?
    Dim boolSel As Boolean
    
    boolSel = LVSelectCount(ListViewFiles) > 0 And mDisk.Unprotected
    
    cmdDelete.Enabled = boolSel
    cmdLockFile.Enabled = boolSel
    cmdUnlockFile.Enabled = boolSel
End Sub

Private Sub ListViewFiles_DblClick()
    ' Dump file
    Dim f As String
    
    f = ListViewFiles.SelectedItem
    DumpDFSFile f, mDisk
End Sub

Private Sub ListViewFiles_KeyUp(KeyCode As Integer, Shift As Integer)
    ListViewFiles_Click
End Sub

Private Sub ListViewFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And (ListViewFiles.HitTest(x, y) Is Nothing) Then
        LVSelectNothing ListViewFiles
    End If
End Sub

Private Sub ListViewFiles_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ListViewFiles_Click
End Sub

Private Sub NewDSDImage_Click()
    ' Create new .DSD image and open it
    NewDisk 2
End Sub

Private Sub NewSDSImage_Click()
    ' Create new .SSD image and open it
    NewDisk 1
End Sub

Private Sub NewDisk(bytSides As Byte)
    ' Create new disk
    Dim d As disk_type
    
    d = Disk_New(mDiskTable, bytSides, Me.Hwnd)
    If d.ImageFile <> "" Then
        CloseDisk True
        mDisk = d
        OpenDisk "", 0
    End If
End Sub

Private Sub OpenImage_Click()
    ' Open Image
    Dim varFile As Variant
    Dim strFile As String
    
    varFile = glrCommonFileOpenSave(Filter:=glrAddFilterItem("", _
                "DFS Images", "*.ssd;*.img;*.dsd;*.mmb"), Hwnd:=Me.Hwnd)
    
    If Not IsNull(varFile) Then
        strFile = varFile
        CloseDisk True
        OpenDisk strFile, 0
    End If
End Sub

Private Sub UpdateInfoBox()
    ' Status: Show info about disk
    Dim strInf

    With mDisk.Catalogue
        If mDisk.ImageFile <> "" Then
            strInf = ShowInt(.FileCount, " file") & _
                 ", " & .Sectors & " Sectors" & _
                 ", " & ShowSize(.SectorsUsed * SecSize) & " used" & _
                 ", " & ShowSize((.Sectors - .SectorsUsed - 2) * SecSize) & " free"
        End If
    End With
    
    StatusBar1.SimpleText = strInf
End Sub

Public Sub OpenDisk(strFile As String, bytSide As Byte, _
                            Optional intDiskNo As Integer = 0)
    ' Open disk
    Dim x As Integer
    Dim strInf As String
    Dim boolOK As Boolean
    Dim strExt As String
    
    DiskTitle.Text = ""
    LabelCycleNo.Caption = ""
    ComboOption.Text = ""
    ListViewFiles.ListItems.Clear
    
    If strFile <> "" Or mDisk.ImageFile <> "" Then
        If strFile <> "" Then
            ' MMB file?
            strExt = UCase(Right(strFile, 4))
            If strExt = ".MMB" Then
                ' Read disk table
                If mDiskTable.ImageName <> strFile Then
                    mDiskTable = ReadDiskTable(strFile)
                    'Debug.Print "Disk count="; mDiskTable.ValidDiskCount
                
                    UpdateDiskTableList
                End If
                
                ' Get first formatted sisk
                If intDiskNo = -1 Then
                    For x = 0 To MaxDisks = -1
                        If mDiskTable.Disk(x).Formatted Then
                            intDiskNo = x
                            Exit For
                        End If
                    Next
                End If
                
                ' Validate disk
                boolOK = intDiskNo >= 0 And intDiskNo < MaxDisks
                If boolOK Then
                    boolOK = mDiskTable.Disk(intDiskNo).Formatted
                End If
                
                If Not boolOK Then
                    xBox "Disk No. " & intDiskNo & " invalid or disk unformated!"
                End If
            Else
                boolOK = True
            End If
            
            If boolOK Then
                ' Open disk
                mDisk = Disk_Open(strFile, intDiskNo, bytSide)
                SetDiskTableListValue
                If mDisk.MMB Then
                    mDisk.Locked = mDiskTable.Disk(mDisk.DiskNo).ReadOnly
                End If
            End If
        End If
        
        With mDisk.Catalogue
            DiskTitle.Text = .DiskTitle
            For x = 1 To .FileCount
                FItem (x)
            Next
            LabelCycleNo.Caption = .CycleNo
            ComboOption.Text = BootOpt(.Option)
        End With
    End If
    
    boolOK = mDisk.ImageFile <> ""
    ListViewFiles.Enabled = boolOK
    ListViewFiles.BackColor = IIf(boolOK, vbWhite, Shape1.BackColor)
    DiskTitle.Enabled = boolOK
    DiskTitle.BackColor = ListViewFiles.BackColor
    ComboOption.Enabled = boolOK
    ComboOption.BackColor = ListViewFiles.BackColor
    cmdSelectAll.Enabled = boolOK
    LVSelectNothing ListViewFiles
    CloseImage.Enabled = boolOK
    
    ComboDiskList.Visible = boolOK And mDisk.MMB
    DiskLocked.Visible = boolOK And mDisk.MMB And mDisk.Locked
    DiskUnlocked.Visible = boolOK And mDisk.MMB And Not mDisk.Locked
    
    OptionSide0.Enabled = boolOK And mDisk.DoubleSided
    OptionSide1.Enabled = boolOK And mDisk.DoubleSided
    OptionSide0.Value = mDisk.Side = 0
    OptionSide1.Value = mDisk.Side = 1
    
    ProtectionDisabled.Enabled = boolOK
    mDisk.DisableProtection = Not mDisk.DisableProtection
    ProtectionDisabled_Click

    UpdateInfoBox
    If mDisk.ImageFile = "" Then
        Caption = "DFS Imager"
    Else
        Caption = Dir(mDisk.ImageFile, vbNormal)
    End If
End Sub

Private Function FIcon(x As Integer) As String
    ' File: Return name of icon
    FIcon = "File"
    With mDisk.Catalogue.Files(x)
        If .Locked Then FIcon = "FileLocked"
        If IsBASIC(.Exec) Then
            FIcon = "Basic" & FIcon
        ElseIf .Exec = &HFFCCCC Then
            FIcon = "Puc" & FIcon
        End If
    End With
End Function

Private Sub FItem(x As Integer)
    ' File: Add file to ListViewFiles
    Dim f As ListItem
    
    With mDisk.Catalogue.Files(x)
        Set f = ListViewFiles.ListItems.Add(x, "FILE" & x)
        f.Text = .Directory & "." & .Name
        f.Icon = FIcon(x)
        f.SmallIcon = f.Icon
        f.SubItems(1) = HexN(.Load, 6)
        f.SubItems(2) = HexN(.Exec, 6)
        f.SubItems(3) = HexN(.Length, 6)
        f.SubItems(4) = HexN(.StartSector, 3)
        Set f = Nothing
    End With
End Sub

Private Sub cmdDelete_Click()
    ' Button: Delete selected files
    If MsgBox("Are you sure?", _
            vbExclamation Or vbYesNo) = vbYes Then
        SelFilesAction FileAct_Delete
    End If
    ListViewFiles.SetFocus
End Sub

Private Sub cmdSelectAll_Click()
    ' Button: Select all files
    LVSelectAll ListViewFiles
    ListViewFiles_Click
    ListViewFiles.SetFocus
End Sub

Private Sub ListViewFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ' Files: Sort on column
    ListViewFiles.SortKey = ColumnHeader.Index - 1
End Sub

Private Sub ListViewFiles_OLECompleteDrag(Effect As Long)
    ' Files: Drag from ListViewFiles complete
End Sub

Private Sub ListViewFiles_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Files: Drop on to ListViewFiles
    Dim varFilename
    
    If Data.GetFormat(vbCFFiles) Then
        For Each varFilename In Data.Files
            If Not ImportDFSFile(mDisk, mDiskTable, CStr(varFilename)) Then
                Exit For
            End If
        Next
        OpenDisk "", 0
    End If
    ListViewFiles.Arrange = lvwAutoLeft ' Keep's changing!
End Sub

Private Sub ListViewFiles_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    ' Files: Dragging over ListViewFiles
    If Data.GetFormat(vbCFFiles) And mDisk.Unprotected Then
        Effect = vbDropEffectCopy And Effect
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub ListViewFiles_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    ' Files: Drag from ListViewFiles started
    ' Note:  Files will be left in the "temporary directory"
    Dim d(1 To 10) As Byte
    Dim srcf As Long
    Dim strFile As String
    Dim x As Integer
    Dim y As Integer
    
    mintSIK = CLng(Mid(ListViewFiles.SelectedItem.Key, 5))
    
    ' Save files from image to temporary directory
    With mDisk
        srcf = FreeFile
        Open .ImageFile For Binary Access Read As srcf
            For x = 1 To ListViewFiles.ListItems.Count
                If ListViewFiles.ListItems(x).Selected Then
                    y = CLng(Mid(ListViewFiles.ListItems(x).Key, 5))
                    strFile = ExtractDFSFile(srcf, y, mDisk)
                    If strFile = "" Then
                        Exit For
                    Else
                        Data.Files.Add strFile
                        Data.Files.Add strFile & ".inf"
                    End If
                End If
            Next
        Close srcf
    End With
    
    Data.SetData , vbCFFiles
    AllowedEffects = vbDropEffectCopy
End Sub

Private Sub SelFilesAction(intAct As Integer)
    ' Act on selected files
    ' i.e. delete, lock & unlock
    Dim x As Integer
    Dim y As Integer
    Dim l As Boolean
    Dim m As Boolean
   
    DiskCatalogue_Refresh mDisk
    
    l = intAct = FileAct_Lock
    m = False
    
    With ListViewFiles
        For x = 1 To .ListItems.Count
            If .ListItems(x).Selected Then
                With .ListItems(x)
                    Select Case intAct
                        Case FileAct_Delete
                            DeleteDFSFile mDisk, .Text
                            m = True
                        Case FileAct_Lock, FileAct_Unlock
                            y = CLng(Mid(.Key, 5))
                            If mDisk.Catalogue.Files(y).Locked <> l Then
                                mDisk.Catalogue.Files(y).Locked = l
                                m = True
                                .Icon = FIcon(y)
                                .SmallIcon = .Icon
                            End If
                    End Select
                End With
            End If
        Next
    End With
    
    If m Then
        If intAct = FileAct_Delete Then
            DiskCatalogue_Save mDisk, mDiskTable, False
            OpenDisk "", 0
        Else
            SaveCat False  ' No changes to info
        End If
    End If
End Sub

Private Sub OptionSide0_Click()
    ' Select Side 0
    If mDisk.Side <> 0 Then
        OpenDisk mDisk.ImageFile, 0, mDisk.DiskNo
    End If
End Sub

Private Sub OptionSide1_Click()
    ' Select Side 1
    If mDisk.Side <> 1 Then
        OpenDisk mDisk.ImageFile, 1, mDisk.DiskNo
    End If
End Sub

Private Sub ProtectionDisabled_Click()
    ' Enable/Disable Protection
    With mDisk
        .DisableProtection = Not .DisableProtection
        .Unprotected = .DisableProtection Or Not .Locked

        ProtectionDisabled.Checked = .DisableProtection
    
        ListViewFiles_Click
        DiskTitle.Locked = Not .Unprotected
        ComboOption.Locked = Not .Unprotected
        cmdCompact.Enabled = .Unprotected And mDisk.Open
    End With
End Sub

Private Sub UpdateDiskTableList()
    ' Refresh Disk Table list
    Dim x As Integer
    Dim n As String
    
    ComboDiskList.Clear
    
    For x = 0 To MaxDisks - 1
        With mDiskTable.Disk(x)
            If .Formatted Then
                n = x & ": " & .DiskTitle
                ComboDiskList.AddItem n
            End If
        End With
    Next
End Sub

Private Sub SetDiskTableListValue()
    ' Set DiskTableList to show current disk
    Dim n As String
    
    If mDisk.MMB Then
        If mDisk.DiskNo >= 0 Then
            n = mDisk.DiskNo & ": " & mDiskTable.Disk(mDisk.DiskNo).DiskTitle
            ComboDiskList = n
        End If
    Else
        ComboDiskList = ""
    End If
End Sub

Private Sub SplitDSD_Click()
    ' Convert DSD to two SSD's images
    ' Opens side 0 if successful
    Dim varFile As Variant
    Dim strFile As String
    
    varFile = glrCommonFileOpenSave( _
            Filter:=glrAddFilterItem("", _
            "DFS Images", "*.dsd"), _
            Hwnd:=Me.Hwnd, _
            DialogTitle:="Select Double Sided Disk")
    
    If Not IsNull(varFile) Then
        strFile = varFile
        OpenDisk ConvertDSD(strFile), 0
    End If
End Sub
