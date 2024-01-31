Attribute VB_Name = "Globals"
' General variables, constants & functions
' Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit

Public Const ProgVersion = "14.02.06"

Public Const SecSize As Long = 256 ' bytes
Public Const KB As Long = 1024 ' bytes
Public Const MB As Long = 1024 * KB

' Sizes in Sectors
Public Const DiskCatalogueSize = 2 * SecSize ' DFS
Public Const DiskSectors As Integer = 800 ' Only size supported
Public Const DiskSize As Long = DiskSectors * SecSize

Public Const CatalogueMaxFiles = 31

' MMB variables
' Disk Table - Sizes in Sectors
Public Const Disk1Offset As Long = 32 * SecSize
Public Const DiskTableSize As Long = 32 * SecSize
Public Const MaxDisks = 511

Public Const MMB_DiskSize As Long = 800  ' sectors
Public Const MMB_DiskTableSize As Long = 32 ' sectors

' Disk table values
Public Const DiskReadOnly = 0
Public Const DiskReadWrite = &HF
Public Const DiskUnformatted = &HF0
Public Const DiskInvalid = &HFF

' Show integer with unit
Public Function ShowInt(varNumber As Variant, strUnit As String, _
        Optional strPlural As String = "") As String
    If varNumber = 1 Then
        ShowInt = varNumber & strUnit
    Else
        If strPlural = "" Then
            ShowInt = varNumber & strUnit & "s"
        Else
            ShowInt = varNumber & strPlural
        End If
    End If
End Function

' Return size in bytes, KB or MB dependent on size
Public Function ShowSize(lngBytes As Long) As String
    Dim x As Integer
    Dim z As Double
    Dim u As String
    
    If lngBytes < 0 Then
        ShowSize = "#Error#"
        Exit Function
    ElseIf lngBytes < KB Then
        z = lngBytes
        If lngBytes = 1 Then u = "byte" Else u = "bytes"
    ElseIf lngBytes < MB Then
        z = lngBytes / KB
        u = "KB"
    Else
        z = lngBytes / MB
        u = "MB"
    End If
    
    ShowSize = Format(z, "0.00")
    x = InStr(ShowSize, ".")
    If x > 0 Then
        Do While Right(ShowSize, 1) = "0"
            ShowSize = Left(ShowSize, Len(ShowSize) - 1)
        Loop
    End If
    If Right(ShowSize, 1) = "." Then
        ShowSize = Left(ShowSize, Len(ShowSize) - 1)
    End If
    ShowSize = ShowSize & " " & u
End Function

' Return Hex no. h padded to length l
Public Function HexN(h As Variant, l As Integer, _
            Optional strPad As String = "0") As String
    HexN = Right(String(l, strPad) & Hex$(h), l)
End Function

' Return pathname in temporary folder (in Window's temp folder)
Public Function TempFolder(Optional strFile As String = "") As String
On Error GoTo err_
    Dim f As String
    f = fReturnTempDir & "\mmbeeb.tmp"
    If Dir(f, vbDirectory) = "" Then
        MkDir f
    End If
    TempFolder = f & "\" & strFile
exit_:
    Exit Function
err_:
    eBox "TempFolder"
    Resume exit_
End Function

' Calculate CRC (Cyclic Redundancy Check)
' Based on code in AUG page 349
Public Function CalcCRC(dat() As Byte, datlen As Long) As Long
    Dim x As Long
    Dim y As Long
    Dim crc As Long
    
    crc = 0
    For y = 1 To datlen
        crc = crc Xor (CLng(dat(y)) * &H100)
        For x = 0 To 7
            crc = crc * 2
            If crc >= &H10000 Then
                crc = (crc - &H10000 + 1) Xor &H1020
            End If
        Next
    Next
    
    CalcCRC = crc
End Function

' Return Boot Option String
Public Function BootOpt(bytOption As Byte) As String
    Select Case bytOption
        Case 0: BootOpt = "None"
        Case 1: BootOpt = "*LOAD $.!BOOT"
        Case 2: BootOpt = "*RUN $.!BOOT"
        Case 3: BootOpt = "*EXEC $.!BOOT"
    End Select
End Function

' Return Boot Option No
Public Function BootOptNo(strOption As String) As Byte
    BootOptNo = 0
    Select Case strOption
        Case "*LOAD $.!BOOT": BootOptNo = 1
        Case "*RUN $.!BOOT":  BootOptNo = 2
        Case "*EXEC $.!BOOT": BootOptNo = 3
    End Select
End Function

' Convert BCD byte to binary
Public Function BCDtoBin(bcd As Byte) As Byte
    BCDtoBin = CByte(Hex$(bcd))
End Function

' Convert binary to BCD byte
Public Function BintoBCD(bin As Byte) As Byte
    BintoBCD = CByte("&H" & bin)
End Function

' Parse string
Public Function Parse(ByRef s As String) As String
    Dim i As Integer
    Parse = ""
    s = Trim(s)
    If s <> "" Then
        i = InStr(s, " ")
        If i > 0 Then
            Parse = Left(s, i - 1)
            s = Mid(s, i + 1)
        Else
            Parse = s
            s = ""
        End If
    End If
    Debug.Print "Parse '"; s; "' -- '"; Parse; "'"
End Function

' Report error
Public Sub eBox(strCaption As String)
    MsgBox strCaption & vbNewLine & _
            "Error: " & Err.Description & vbNewLine & Err.Number, _
            vbExclamation Or vbOKOnly
    Debug.Print "Error - "; strCaption; ": "; Err.Description; " ("; Err.Number; ")"
End Sub

' Message
Public Sub xBox(strMessage As String)
    MsgBox strMessage, vbExclamation Or vbOKOnly
End Sub

' A BASIC file?
Public Function IsBASIC(lngExec As Long) As Boolean
    Dim x As Long
    
    x = lngExec And 65535
    IsBASIC = x = 32799 Or x = 32803 Or x = 32811
End Function
