VERSION 5.00
Begin VB.Form frmViewFile 
   Caption         =   "ViewFile"
   ClientHeight    =   5925
   ClientLeft      =   5865
   ClientTop       =   1830
   ClientWidth     =   4815
   Icon            =   "frmViewFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   4815
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5910
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmViewFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' View File Form
' Created/Written by Martin Mather 2005
' http://mmbeeb.mysite.wanadoo-members.co.uk/

Option Explicit
    
Private Const BytesPerLine = 8

Public Sub Dump(bytData() As Byte, strName As String)
On Error GoTo err_
    ' Data to dump
    Dim i As Long
    Dim j As Long
    Dim s As String
    Dim c As String
    Dim z As Long
    Dim p As Long
    Dim ss As String
    
    Me.MousePointer = vbHourglass
    z = UBound(bytData)
    
    List1.Clear
    
    i = 1
    Do While i <= z
        c = ""
        s = hx(i - 1, 5) & " - "
        For j = 1 To BytesPerLine
            If i <= z Then
                s = s & hx(bytData(i), 2) & " "
                If bytData(i) < 32 Then
                    c = c & "."
                Else
                    c = c & Chr(bytData(i))
                End If
            Else
                s = s & "   "
                c = c & " "
            End If
            i = i + 1
        Next
        s = s & "- " & c
        
        List1.AddItem s
    Loop
    
    Me.Caption = "Dump of '" & strName & "'"
exit_:
    Me.MousePointer = vbDefault
    Exit Sub
err_:
    eBox "Dump file"
    Resume exit_
End Sub

Private Function hx(l As Variant, n As Long)
    hx = Right("00000000" & Hex$(l), n)
End Function

