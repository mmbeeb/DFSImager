Attribute VB_Name = "DiskOps"
' General Disk Operations
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

' Sectors interleaved on .DSD disk
Private Const DSD_Interleave = 10

' Catalogue file entry
Public Type cataloguefile_Type
    Name As String
    Directory As String
    FullName As String ' Dir + . + Name
    Locked As Boolean
    Load As Long
    Exec As Long
    Length As Long
    StartSector As Integer
    SectorsUsed As Integer
End Type

' Disk catalogue
Public Type catalogue_type
    DiskTitle As String
    CycleNo As Byte
    Option As Byte
    Sectors As Integer
    SectorsUsed As Integer
    LastSector As Integer
    FileCount As Byte
    Files(1 To CatalogueMaxFiles) As cataloguefile_Type
    Modified As Boolean
End Type

' Disk
Public Type disk_type
    ImageFile As String
    Open As Boolean
    DoubleSided As Boolean
    DiskNo As Integer
    Side As Byte
    Catalogue As catalogue_type
    DSD As Boolean
    MMB As Boolean
    DisableProtection As Boolean
    Locked As Boolean
    Unprotected As Boolean ' = DisableProtection or Not Locked
End Type

Public Function Disk_Open(strFile As String, _
                intDiskNo As Integer, bytSide As Byte) As disk_type
    ' Open disk
    Dim Disk As disk_type
    Dim strExt As String
    
    Debug.Print "Disk_Open: "; strFile; " No="; intDiskNo
    
    strExt = UCase(Right(strFile, 4))
    
    With Disk
        .MMB = strExt = ".MMB"
        .DSD = strExt = ".DSD"
    
        .ImageFile = strFile
        .Open = strFile <> ""
        .DiskNo = intDiskNo
        .DoubleSided = .DSD
        .Side = bytSide
        
        .DisableProtection = False
        If .MMB Then
            .Locked = True
        Else
            .Locked = False
        End If
        .Unprotected = .DisableProtection Or Not .Locked
    End With
    
    If Not DiskCatalogue_Read(Disk) Then
        Disk_Close Disk
        Debug.Print , "disk closed"
    Else
        Debug.Print , "Open : "; Disk.Catalogue.DiskTitle
        Debug.Print , "Files: "; Disk.Catalogue.FileCount
    End If
    
    Disk_Open = Disk
End Function

Public Function Disk_Close(Disk As disk_type)
    ' Close disk
    Debug.Print "Disk_Close: "; Disk.ImageFile
    With Disk
        .ImageFile = ""
        .Open = False
    End With
End Function

Public Function Disk_New(DiskTable As disktable_type, _
                    bytSides As Byte, Hwnd As Long) As disk_type
    ' Create new disk (.SSD or .DSD)
    ' New Image
    Dim Disk As disk_type
    Dim varFile As Variant
    Dim strFile As String
    Dim strFilter As String
    Dim strDescription As String
    Dim x As Integer
    
    If bytSides = 1 Then
        strFilter = "*.ssd;*.img"
        strDescription = "Single Sided Disk"
    Else
        strFilter = "*.dsd"
        strDescription = "Double Sided Disk"
    End If
    
    varFile = glrCommonFileOpenSave(glrOFN_OVERWRITEPROMPT, _
                    OpenFile:=False, _
                    Filter:=glrAddFilterItem("", strDescription, strFilter), _
                    Hwnd:=Hwnd)
    
    If Not IsNull(varFile) Then
        strFile = varFile
        Debug.Print "Disk_New: "; strFile
        Debug.Print "   sides="; bytSides
        
        If Dir(strFile, vbNormal) <> "" Then ' User already prompted on dialogue
            Kill strFile
        End If

        With Disk
            .ImageFile = strFile
            .Open = True
            .DisableProtection = False
            .Locked = False
            .Unprotected = True
            
            .MMB = False
            .DSD = bytSides = 2
            .DoubleSided = .DSD
            
            With .Catalogue
                .DiskTitle = ""
                .Option = 0
                .FileCount = 0
                .Sectors = DiskSectors
                .SectorsUsed = 0
                .Modified = True
                .LastSector = 2
            End With
            
            For x = 0 To bytSides - 1
                .Side = x
                .Catalogue.CycleNo = 99
                DiskCatalogue_Save Disk, DiskTable, True
            Next
            
            .Side = 0
        End With
    End If
    
    Disk_New = Disk
End Function

Private Function SecPtr(Disk As disk_type, intSector As Integer) As Long
    ' Return file byte offset for sector
    Dim blk As Integer
    Dim sec As Integer
    
    If Disk.DSD Then
        ' .DSD files
        blk = intSector \ DSD_Interleave
        sec = intSector - (blk * DSD_Interleave)
        
        SecPtr = (blk * 2 + Disk.Side) * DSD_Interleave + sec
        'Debug.Print "Sec "; intSector, blk, sec, SecPtr
    ElseIf Disk.MMB Then
        ' .MMB files
        SecPtr = (MMB_DiskTableSize + Disk.DiskNo _
                         * MMB_DiskSize) + intSector
    Else
        ' .SSD or .IMG files
        SecPtr = intSector
    End If
    
    SecPtr = SecPtr * SecSize + 1
End Function

Public Function DiskCatalogue_Read(Disk As disk_type) As Boolean
    ' Read Disk Catalogue
On Error GoTo err_
    Dim cat(0 To 511) As Byte
    Dim f As Long
    Dim x As Integer
    Dim y As Integer
    Dim o As Integer
    Dim b As Byte
    Dim mixedbyte As Byte
    
    Debug.Print "DiskCatalogue_Read: "; Disk.ImageFile; " Side="; Disk.Side
    DiskCatalogue_Read = True
    
    If Disk.ImageFile <> "" Then
        ' Read disk catalogue
        f = FreeFile
        Open Disk.ImageFile For Binary Access Read As f
        Get f, SecPtr(Disk, 0), cat
        Close f
        
        With Disk.Catalogue
            .Modified = False
            
            ' Read disk title (Chr 0 terminated)
            .DiskTitle = ""
            x = 0
            Do
                If x > 7 Then b = cat(x + &HF8) Else b = cat(x)
                If b > 0 Then
                    .DiskTitle = .DiskTitle & Chr(b)
                End If
                x = x + 1
            Loop Until x = 11 Or b = 0
            
            .Option = (cat(&H106) And &HF0) \ &H10
            .Sectors = (cat(&H106) And &H3) * &H100 + cat(&H107)
            .SectorsUsed = 0
            .CycleNo = BCDtoBin(cat(&H104))
            .FileCount = cat(&H105) / 8
            
            For y = 1 To .FileCount
                With .Files(y)
                    o = y * 8
                    
                    ' Filename (padded with spaces)
                    .Name = ""
                    For x = 0 To 6
                        .Name = .Name & Chr(cat(o + x))
                    Next
                    .Name = RTrim(.Name)
                    .Directory = Chr(cat(o + 7) And &H7F)
                    .FullName = .Directory & "." & .Name
                    .Locked = cat(o + 7) >= &H80
            
                    o = o + &H100
                    mixedbyte = cat(o + 6)
                    
                    .Load = CLng(cat(o + 1)) * &H100 + CLng(cat(o))
                    If (mixedbyte And &HC) > 0 Then .Load = .Load + &HFF0000
     
                    .Exec = CLng(cat(o + 3)) * &H100 + CLng(cat(o + 2))
                    If (mixedbyte And &HC0) > 0 Then .Exec = .Exec + &HFF0000
    
                    .Length = ((mixedbyte And &H30) \ &H10) * &H10000 + _
                        CLng(cat(o + 5)) * &H100 + CLng(cat(o + 4))
    
                    .StartSector = (mixedbyte And &H3) * &H100 + CLng(cat(o + 7))
    
                    .SectorsUsed = ((mixedbyte And &H30) \ &H10) * &H100 + _
                        CLng(cat(o + 5))
                    If cat(o + 4) > 0 Then .SectorsUsed = .SectorsUsed + 1
                    
                    Disk.Catalogue.SectorsUsed = _
                            Disk.Catalogue.SectorsUsed + .SectorsUsed
                End With
            Next
        End With
    End If
    
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Reading Catalogue " & Disk.Side & " of '" & Disk.ImageFile & "'"
    DiskCatalogue_Read = False
    Resume exit_
End Function

Public Sub DiskCatalogue_Refresh(Disk As disk_type)
    ' Refresh Disk Catalogue
    DiskCatalogue_Read Disk
End Sub

Public Sub DiskCatalogue_Save(Disk As disk_type, _
                            DiskTable As disktable_type, _
                            boolTitleChanged As Boolean)
    ' Save Disk Catalogue
On Error GoTo err_
    Dim cat(0 To 511) As Byte
    Dim f As Long
    Dim x As Integer
    Dim y As Integer
    Dim o As Integer
    Dim b As Byte
    Dim mixedbyte As Byte
    
    If Disk.ImageFile = "" Then Exit Sub
    
    Debug.Print "DiskCatalogue_Save: "; Disk.ImageFile
    With Disk.Catalogue
        ' Increment cycle no.
        .CycleNo = (.CycleNo + 1) Mod 100
        
        ' Check & fix file list
        'For y = 1 To .FileCount
        '    Debug.Print "sc "; y, .Files(y).FullName
        'Next
        
        ' Disk title (Chr 0 terminated)
        For x = 0 To 10
            If x + 1 > Len(.DiskTitle) Then
                b = 0
            Else
                b = Asc(Mid(.DiskTitle, x + 1, 1))
            End If
            
            If x > 7 Then cat(x + &HF8) = b Else cat(x) = b
        Next
        
        cat(&H106) = (.Option * &H10) Or (.Sectors \ &H100)
        cat(&H107) = .Sectors And &HFF
        cat(&H104) = BintoBCD(.CycleNo)
        cat(&H105) = .FileCount * 8
        
        For y = 1 To .FileCount
            With .Files(y)
                o = y * 8
                
                ' Filename (padded with spaces)
                For x = 0 To 6
                    cat(o + x) = Asc(Mid(.Name & String(7, " "), x + 1, 1))
                Next
                cat(o + 7) = Asc(.Directory) Or IIf(.Locked, &H80, 0)
     
                o = o + &H100
                mixedbyte = 0
                
                cat(o) = .Load And &HFF
                cat(o + 1) = (.Load \ &H100) And &HFF
                If .Load >= &H10000 Then mixedbyte = &HC
 
                cat(o + 2) = .Exec And &HFF
                cat(o + 3) = (.Exec \ &H100) And &HFF
                If .Exec >= &H10000 Then mixedbyte = mixedbyte Or &HC0

                cat(o + 4) = .Length And &HFF
                cat(o + 5) = (.Length \ &H100) And &HFF
                mixedbyte = mixedbyte Or (.Length \ &H1000 And &H30)
                
                cat(o + 7) = .StartSector And &HFF
                mixedbyte = mixedbyte Or (.StartSector \ &H100 And 3)
                
                cat(o + 6) = mixedbyte
            End With
        Next
    End With
    
    ' Write disk catalogue
    f = FreeFile
    Open Disk.ImageFile For Binary Access Write As f
    Put f, SecPtr(Disk, 0), cat
    Close f
    
    If Disk.MMB And boolTitleChanged Then
        ' Update disk table
        UpdateDiskTable DiskTable, Disk.DiskNo, Disk.Catalogue.DiskTitle
    End If
        
    Disk.Catalogue.Modified = False
exit_:
On Error Resume Next
    Close f
    Exit Sub
err_:
    eBox "Writing Catalogue of '" & Disk.ImageFile & "'"
    Resume exit_
End Sub

Private Function GetFileIndex(ByRef cat As catalogue_type, _
                strFileName As String) As Byte
    ' Return file index
    Dim y As Integer
    GetFileIndex = 0
    With cat
        For y = 1 To .FileCount
            If .Files(y).FullName = strFileName Then
                GetFileIndex = y
                Exit Function
            End If
        Next
    End With
End Function

Public Function ExtractDFSFile(lngSrcFileHandle As Long, _
        intFileNo As Integer, Disk As disk_type) As String
    ' Saves dfs file (and .inf) from image to temporary folder
    ' Returns pathname, or "" if error occurs
On Error GoTo err_
    Dim f As Long
    Dim strName As String
    Dim strPath As String
    Dim strInf As String
    Dim bytData() As Byte
    
    ExtractDFSFile = ""
    With Disk.Catalogue.Files(intFileNo)
        strName = .Directory & "." & .Name
        strPath = TempFolder(strName)
        ReDim bytData(1 To .Length)
        
        RW_FileData lngSrcFileHandle, Disk, .StartSector, .Length, bytData
        
        f = FreeFile
        Open strPath For Binary Access Write As f
        Put f, , bytData
        Close f
    End With
    
    If WriteInf(strPath, Disk.Catalogue.Files(intFileNo), bytData) Then
        ExtractDFSFile = strPath
    End If
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Extract file"
    Resume exit_
End Function

Public Function ImportDFSFile(Disk As disk_type, DiskTable As disktable_type, _
                  strDosFilename As String) As Boolean
    ' Import file in to disk image
On Error GoTo err_
    Dim strExt As String
    Dim f As Long
    Dim file As cataloguefile_Type
    Dim l As Long
    
    ImportDFSFile = True
    l = FileLen(strDosFilename)
    If l > 0 And l <= 200 * KB Then
        strExt = LCase(Right(strDosFilename, 4))
        If strExt <> ".inf" Then ' Ignore .inf files
            ' Read .inf file (must exist)
            If ReadInf(strDosFilename, file) Then
                ' read file
                Dim bytData() As Byte
                ReDim bytData(1 To file.SectorsUsed * SecSize)
                f = FreeFile
                Open strDosFilename For Binary Access Read As f
                Get f, , bytData
                Close f
            
                ImportDFSFile = AddFile(Disk, DiskTable, file, bytData)
            End If
        End If
    End If
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    ImportDFSFile = False
    eBox "Import file"
    Resume exit_
End Function

Private Function WriteInf(strDosFilename As String, _
                file As cataloguefile_Type, bytData() As Byte) As Boolean
    ' Create .inf file
On Error GoTo err_
    Dim strInf As String
    Dim f As Long
    
    WriteInf = False
    With file
        strInf = Left(.FullName & String(8, " "), 10) _
                & HexN(.Load, 6, " ") & " " _
                & HexN(.Exec, 6, " ")
        If .Locked Then strInf = strInf & " Locked"
        strInf = strInf & " CRC= " & _
                Hex(CalcCRC(bytData, .Length))
    End With
    
    f = FreeFile
    Open strDosFilename & ".inf" For Binary Access Write As f
    Put f, , strInf
    Close f
    WriteInf = True
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Write .inf"
    Resume exit_
End Function

Private Function ReadInf(strDosFilename As String, _
                file As cataloguefile_Type) As Boolean
    ' Read .inf file
    ' Returns True if successful
On Error GoTo err_
    Dim f As Long
    Dim strInf As String
    
    ReadInf = False
    With file
        .Length = FileLen(strDosFilename)
        If .Length > 0 Then
            Debug.Print "DFSFile: "; strDosFilename, .Length
                
            If Dir(strDosFilename & ".inf", vbNormal) <> "" Then
                f = FreeFile
                Open strDosFilename & ".inf" For Input As f
                Input #f, strInf
                Close f
                Debug.Print "Inf: "; strInf
                
                .FullName = Parse(strInf)
                .Directory = Left(.FullName, 1)
                .Name = Mid(.FullName, 3)
                
                .Load = CLng("&H" & Parse(strInf))
                .Exec = CLng("&H" & Parse(strInf))
                .Locked = InStr(UCase(Mid(strInf, 25)), "LOCKED") > 0
                    
                .SectorsUsed = (.Length \ &H100)
                If (.Length And &HFF) > 0 Then .SectorsUsed = .SectorsUsed + 1
                ReadInf = .Name <> "" And .Directory <> ""
            End If
        End If
    End With
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Read .inf"
    Resume exit_
End Function

Private Function AddFile(Disk As disk_type, DiskTable As disktable_type, _
                    file As cataloguefile_Type, bytData() As Byte) As Boolean
    ' Add file to DFS disk
    ' Returns true if successful
    Dim f As Long
    Dim y As Integer
    Dim p As Integer
    Dim o As Boolean
    
    AddFile = False
    
    DiskCatalogue_Refresh Disk
    
    ' Catalogue full?
    y = GetFileIndex(Disk.Catalogue, file.FullName)
    If y = 0 And Disk.Catalogue.FileCount = 31 Then
        MsgBox "Catalogue full!", vbExclamation Or vbOKOnly
        Exit Function
    End If
    
    ' Overwrite existing file?
    If y > 0 Then
        o = MsgBox("Overwrite " & file.FullName, _
                    vbYesNo Or vbExclamation) = vbYes
        If o Then
            o = DeleteDFSFile(Disk, file.FullName)
        End If
    Else
        o = True
    End If
    
    If o Then
        With file
            ' Room on disk?
            If .SectorsUsed > (Disk.Catalogue.Sectors - _
                                    Disk.Catalogue.SectorsUsed) Then
                MsgBox "No room on disk!", vbExclamation Or vbOKOnly
            Else
                .StartSector = GetDiskBlock(Disk.Catalogue, .SectorsUsed)
                If .StartSector = 0 Then
                    ' There is enough room if the disk is compacted
                    If MsgBox("No block large enough!" & vbNewLine & _
                            "Compact disk?", vbExclamation Or vbYesNo) = vbYes Then
                        CompactDisk Disk, DiskTable
                        '.StartSector = cat.LastSector
                    End If
                End If
                
                If .StartSector > 0 Then
                    'Debug.Print "FN   "; .FullName
                    'Debug.Print "Dir  "; .Directory
                    'Debug.Print "N    "; .Name
                    'Debug.Print "Load "; Hex$(.Load)
                    'Debug.Print "Exec "; Hex$(.Exec)
                    'Debug.Print "lock "; .Locked
                    'Debug.Print "SU   "; .SectorsUsed
                    'Debug.Print "strs "; .StartSector
                    
                    ' write file
                    f = FreeFile
                    Open Disk.ImageFile For Binary Access Write As f
                    RW_FileData f, Disk, .StartSector, .Length, bytData, _
                                     blnRead:=False
                    Close f
                    
                    ' add to catalogue (must be in 'start sector' desc. order)
                    With Disk.Catalogue
                        p = 1
                        For y = 1 To .FileCount
                            If .Files(y).StartSector < file.StartSector Then
                                p = y
                                Exit For
                            Else
                                p = y + 1
                            End If
                        Next
                        Debug.Print "Cat ptr="; p
                        If p <= .FileCount Then ' Insert gap
                            For y = .FileCount To p Step -1
                                .Files(y + 1) = .Files(y)
                            Next
                        End If
                        .Files(p) = file
                        .FileCount = .FileCount + 1
                    End With
                    DiskCatalogue_Save Disk, DiskTable, False
                    AddFile = True
                End If
            End If
        End With
    End If
End Function

Public Function GetDiskBlock(cat As catalogue_type, intSize As Integer) As Integer
    ' Return start sector of smallest free block of minimum size of intSize
    ' If zero returned, no block found
    Dim x As Integer
    Dim s As Integer
    Dim b As Integer
    Dim z As Integer
    
    GetDiskBlock = 0
    With cat
        If .FileCount = 0 Then
            GetDiskBlock = 2
        Else
            s = 2
            z = 0
            For x = .FileCount To 0 Step -1
                If x = 0 Then
                    b = .Sectors - s
                    ''Debug.Print x, "END OF DISK", Hex$(s), Hex$(.Sectors), b
                Else
                    b = .Files(x).StartSector - s
                    ''Debug.Print x, .Files(x).FullName, Hex$(s), Hex$(.Files(x).StartSector), b
                End If
                    
                If b >= intSize And (b < z Or z = 0) Then
                    GetDiskBlock = s
                    z = b
                End If
                    
                If x > 0 Then
                    s = .Files(x).StartSector + .Files(x).SectorsUsed
                End If
            Next
            
        End If
    End With
    ''Debug.Print "GDB : "; GetDiskBlock
End Function

Public Function DeleteDFSFile(Disk As disk_type, strFileName As String) As Boolean
    ' Delete DFS file
    ' Assumes catalogue is current and does not write to disk
    ' Returns true if file deleted
On Error GoTo err_
    Dim y As Integer
    Dim x As Integer
   
    Debug.Print "Delete file: "; strFileName
    DeleteDFSFile = False
    
    With Disk.Catalogue
        y = GetFileIndex(Disk.Catalogue, strFileName)
        If y > 0 Then
            If Disk.Catalogue.Files(y).Locked And Not Disk.DisableProtection Then
                xBox "File '" & strFileName & "' is locked"
            Else
                If y < .FileCount Then
                    For x = y To .FileCount - 1
                        .Files(x) = .Files(x + 1)
                    Next
                End If
                .FileCount = .FileCount - 1
                DeleteDFSFile = True
            End If
        Else
            xBox "File '" & strFileName & "' not found"
        End If
    End With
    
exit_:
    Debug.Print "Delete = "; DeleteDFSFile
    Exit Function
err_:
    eBox "DeleteDFSFile"
    Resume exit_
End Function

Public Function CompactDisk(Disk As disk_type, _
                    DiskTable As disktable_type) As Boolean
    ' Compact disk
    ' Assumes catalogue is current
    ' Returns true if any files moved
On Error GoTo err_
    Dim y As Integer
    Dim s As Integer
    Dim z As Integer
    Dim bytData() As Byte
    Dim f As Long
    Dim boolUpdateCat As Boolean
    Dim ok As Boolean
    
    Debug.Print "CompactDisk"
    CompactDisk = False
    f = FreeFile
    boolUpdateCat = False
    Open Disk.ImageFile For Binary Access Read Write As f
    s = 2
    With Disk.Catalogue
        For y = .FileCount To 1 Step -1
            With .Files(y)
                z = .StartSector - s
                ''Debug.Print y, Hex$(s), Hex$(.StartSector), z
                If z > 0 Then
                    ' Move file
                    ReDim bytData(1 To .Length)
                    ok = RW_FileData(f, Disk, .StartSector, .Length, bytData)
                    If ok Then
                        ok = RW_FileData(f, Disk, s, .Length, bytData, _
                                blnRead:=False)
                    End If
                    Erase bytData
                    .StartSector = s
                    boolUpdateCat = True
                End If
                s = s + .SectorsUsed
            End With
            If Not ok Then Exit For
        Next
    End With
    Close f
    If boolUpdateCat And ok Then
        DiskCatalogue_Save Disk, DiskTable, False
    End If
    CompactDisk = boolUpdateCat
exit_:
On Error Resume Next
    Close f
    Exit Function
err_:
    eBox "Compact Disk"
    Resume exit_
End Function

Public Sub DumpDFSFile(strDFSFilename As String, Disk As disk_type)
    ' Open 'frmViewFile' and dump contents of DFS file
On Error GoTo err_
    Dim f As Long
    Dim strName As String
    Dim bytData() As Byte
    Dim intFileNo As Integer
    
    intFileNo = GetFileIndex(Disk.Catalogue, strDFSFilename)
    If intFileNo > 0 Then
        With Disk.Catalogue.Files(intFileNo)
            strName = .Directory & "." & .Name
            
            If .Length > 0 Then
                'Debug.Print "Dump "; strName, Hex$(.Length)
                ReDim bytData(1 To .Length)
                
                f = FreeFile
                Open Disk.ImageFile For Binary Access Read As f
                If RW_FileData(f, Disk, .StartSector, .Length, bytData) Then
                    frmViewFile.Visible = True
                    frmViewFile.Dump bytData, strName
                End If
            Else
                xBox "File '" & strName & "' empty!"
            End If
        End With
    End If
exit_:
On Error Resume Next
    Close f
    Exit Sub
err_:
    eBox "Dump file"
    Resume exit_
End Sub

Private Function RW_FileData(lngFilehandle As Long, Disk As disk_type, _
                                intStartSec As Integer, lngLength As Long, _
                                ByRef bytData() As Byte, _
                                Optional blnRead As Boolean = True) As Boolean
    ' Read or Write DFS file data
    ' (.DSD files - possible fragmented due to interleaving,
    ' so read one sector at a time)
    ' Note: bytData lower bound = 1
On Error GoTo err_
    Dim bytSec(0 To SecSize - 1) As Byte
    Dim intSec As Integer
    Dim lngPtr As Long
    Dim i As Long
    Dim x As Integer
    
    Debug.Print "RW_FileData "; blnRead, Hex$(intStartSec), Hex$(lngLength)
    
    RW_FileData = True
    intSec = intStartSec
    lngPtr = SecPtr(Disk, intStartSec)
    
    If Disk.DSD Then
        ' .DSD files
        i = 1
        Do While i <= lngLength
            'Debug.Print , Hex$(i), intSec, lngPtr
            If blnRead Then
                ' Read sector
                Get lngFilehandle, lngPtr, bytSec
                For x = 0 To SecSize - 1
                    If i <= lngLength Then
                        bytData(i) = bytSec(x)
                        i = i + 1
                    End If
                Next
            Else
                ' Write Sector
                For x = 0 To SecSize - 1
                    If i <= lngLength Then
                        bytSec(x) = bytData(i)
                        i = i + 1
                    End If
                Next
                Put lngFilehandle, lngPtr, bytSec
            End If
            
            ' Increment counter/pointer
            intSec = intSec + 1
            If intSec Mod DSD_Interleave = 0 Then
                lngPtr = lngPtr + (DSD_Interleave + 1) * SecSize
            Else
                lngPtr = lngPtr + SecSize
            End If
        Loop
    Else
        ' .SSD, .IMG or .MMB files
        If blnRead Then
            Get lngFilehandle, lngPtr, bytData
        Else
            Put lngFilehandle, lngPtr, bytData
        End If
    End If
exit_:
    Exit Function
err_:
    eBox "RW_FileData"
    RW_FileData = False
    Resume exit_
End Function

Public Function ConvertDSD(strFileName As String) As String
    ' Convert .dsd to two .ssd's
On Error GoTo err_
    Const BlkSize As Integer = 10 * SecSize ' interleaved sectors
    Dim blk(1 To BlkSize) As Byte
    Dim blks As Long
    Dim src As Long
    Dim obj(1 To 2) As Long
    Dim s As Integer
    Dim fn(1 To 2) As String
    Dim ok As Boolean
    Dim l As Long
    Dim x As Long
    
    fn(1) = Left(strFileName, Len(strFileName) - 4) ' strip .dsd
    fn(2) = fn(1) & "_side1.ssd"
    fn(1) = fn(1) & "_side0.ssd"
    
    Debug.Print "Convert "; strFileName
    Debug.Print " > "; fn(1)
    Debug.Print " > "; fn(2)
    
    l = FileLen(strFileName)
    blks = l \ BlkSize
    If blks * BlkSize < l Then blks = blks + 1
    
    ' Overwrite existing files?
    ok = True
    For x = 1 To 2
        If Dir(fn(x), vbNormal) <> "" Then
            ok = MsgBox("Overwrite " & fn(x) & "?", vbExclamation Or vbYesNo) = vbYes
            If ok Then
                Kill fn(x)
            Else
                Exit For
            End If
        End If
    Next
    
    If ok Then
        obj(1) = FreeFile
        Open fn(1) For Binary Access Write As obj(1)
        obj(2) = FreeFile
        Open fn(2) For Binary Access Write As obj(2)
        src = FreeFile
        Open strFileName For Binary Access Read As src
        s = 1
        For x = 1 To blks
            Get src, , blk
            Put obj(s), , blk
            If s = 1 Then s = 2 Else s = 1
            blks = blks - 1
        Next
        ConvertDSD = fn(1)
    End If
    
exit_:
On Error Resume Next
    If ok Then
        Close obj(1)
        Close obj(2)
        Close src
    End If
    Exit Function
err_:
    eBox "Converting " & strFileName
    Resume exit_
End Function

Public Function CombineSSDs(strFiles() As String, _
                        strTarget As String) As String
    ' Combine two SSD's in to a DSD
On Error GoTo err_
    Dim ok As Boolean
    Dim intBlk As Integer
    Dim lngBlkSize As Long
    Dim bytData() As Byte
    Dim lngSrc(1 To 2) As Long
    Dim lngLen(1 To 2) As Long
    Dim intSectors(1 To 2) As Integer
    Dim lngTgt As Long
    Dim bytSide As Byte
    Dim x As Integer
    
    Debug.Print "CombineSSDs:"
    
    ok = True
    
    For x = 1 To 2
        lngLen(x) = FileLen(strFiles(x))
        intSectors(x) = lngLen(x) \ SecSize
        
        ok = ok And intSectors(x) * SecSize = lngLen(x) And _
                    intSectors(x) >= 2 And _
                    intSectors(x) <= DiskSectors
                    
        'Debug.Print , x, strFiles(x), lngLen(x), intSectors(x), ok
    Next
    
    If ok Then
        ' Open source files
        For x = 1 To 2
            lngSrc(x) = FreeFile
            Open strFiles(x) For Binary Access Read As lngSrc(x)
        Next
        
        ' Open target
        If Dir(strTarget, vbNormal) <> "" Then
            Kill strTarget
        End If
        
        lngTgt = FreeFile
        Open strTarget For Binary Access Write As lngTgt
        
        lngBlkSize = 10 * SecSize
        ReDim bytData(1 To lngBlkSize)
        
        intBlk = 0
        Do While lngLen(1) > 0 Or lngLen(2) > 0
            bytSide = (intBlk And 1) + 1
            
            'Debug.Print intBlk, bytSide, lngLen(1), lngLen(2)
            Get lngSrc(bytSide), , bytData
            lngLen(bytSide) = lngLen(bytSide) - lngBlkSize
            
            Put lngTgt, , bytData
        
            intBlk = intBlk + 1
        Loop
    End If
    
    CombineSSDs = strTarget
    
exit_:
On Error Resume Next
    If ok Then
        Close lngSrc(1)
        Close lngSrc(2)
        Close lngTgt
    End If
    Exit Function
    
err_:
    eBox "Converting " & strFiles(1) & " and " & strFiles(2)
    Resume exit_
End Function

