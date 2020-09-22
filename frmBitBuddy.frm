VERSION 5.00
Begin VB.Form frmBitBuddy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bit Buddy V1.0"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   4050
      TabIndex        =   7
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   375
      Left            =   5175
      TabIndex        =   6
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6300
      TabIndex        =   5
      Top             =   900
      Width           =   1095
   End
   Begin VB.CommandButton cmdMask 
      Caption         =   "Mask Bit"
      Default         =   -1  'True
      Height          =   375
      Left            =   2925
      TabIndex        =   4
      Top             =   900
      Width           =   1095
   End
   Begin VB.TextBox txtHex 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1215
      MaxLength       =   10
      TabIndex        =   2
      Top             =   900
      Width           =   1590
   End
   Begin VB.CheckBox CheckBit 
      Height          =   285
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   495
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "Hex Value:"
      Height          =   285
      Left            =   225
      TabIndex        =   3
      Top             =   945
      Width           =   960
   End
   Begin VB.Label lBit 
      Alignment       =   2  'Center
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   1
      Top             =   225
      Width           =   195
   End
End
Attribute VB_Name = "frmBitBuddy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim curValue As Variant
Dim bEmpty As Boolean

Private Function LeftFill(HexValue, Size As Long, Optional FillChar As String = "0") As String
    LeftFill = "" & HexValue
    While Len(LeftFill) < Size
        LeftFill = FillChar & LeftFill
    Wend
End Function

' convert a number to base 16
Function Dec2Hex(ByVal number As Variant) As String
    Const digits As String = "0123456789ABCDEF" 'valid digits
    Dim digitValue As Variant
     
    ' convert to base 16
    Do While number
        digitValue = number - Int(number / CDec(16)) * CDec(16)
        number = Int(number / CDec(16))
        Dec2Hex = Mid$(digits, digitValue + 1, 1) & Dec2Hex
    Loop
    'Dec2Hex = "&H" & Dec2Hex
End Function

Public Function Hex2Dec(ByVal HexStr As String) As Double
    Dim mult As Double
    Dim DecNum As Double
    Dim ch As String
    mult = 1
    DecNum = 0

    Dim i As Integer
    For i = Len(HexStr) To 1 Step -1
        ch = Mid(HexStr, i, 1)
        If (ch >= "0") And (ch <= "9") Then
            DecNum = DecNum + (Val(ch) * mult)
        Else
            If (ch >= "A") And (ch <= "F") Then
                DecNum = DecNum + ((Asc(ch) - Asc("A") + 10) * mult)
            Else
                If (ch >= "a") And (ch <= "f") Then
                    DecNum = DecNum + ((Asc(ch) - Asc("a") + 10) * mult)
                Else
                    Hex2Dec = 0
                    Exit Function
                End If
            End If
        End If
        mult = mult * 16
    Next i
    Hex2Dec = DecNum
End Function


Private Sub CheckBit_Click(Index As Integer)
    Dim vValue As Variant
    Dim tempPos, tempChar
    Dim tempStr As String
    
    vValue = CDec(0)
    
    On Error GoTo err_handler
    
    vValue = (2 ^ Index) '* CheckBit(Index).Value
            
    If CheckBit(Index).Value = 0 Then
        curValue = curValue - vValue
    Else
        curValue = curValue + vValue
        bEmpty = False
    End If
   
    txtHex.Text = "0x" & LeftFill(Dec2Hex(curValue), 8)
err_handler:
    If Err = 6 Then CheckBit(Index).Value = False
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub cmdClear_Click()
    Dim i As Integer
    For i = 0 To 31
        CheckBit(i).Value = False
    Next i
    txtHex.Text = "0x" & Format(Hex(curValue), "00000000")
End Sub

Private Sub CmdExit_Click()
    Unload Me
    End
End Sub

Private Sub cmdMask_Click()
    Dim tempStr As String
    Dim tempPos
    Dim decValue, decMod
    Dim bit_mask
    Dim i As Integer
    
    tempPos = InStr(1, txtHex.Text, "x", vbTextCompare)
    
    tempStr = Mid(txtHex.Text, tempPos + 1, Len(txtHex.Text) - tempPos + 1)
    decValue = CDec(Hex2Dec(tempStr))
    
    For i = 0 To 31
        CheckBit(i).Value = Unchecked
    Next i
    
    bEmpty = True
    i = 0
    On Error GoTo err_handler
    Do While decValue <> 0
        decMod = decValue - Int(decValue / CDec(2)) * CDec(2)
        If decMod = 1 Then   'This method Blows!!!!!! Causes Overflow
            CheckBit(i).Value = Checked
        End If
        decValue = Int(decValue / CDec(2))
        i = i + 1
    Loop

err_handler:
    If Err = 340 Then
        MsgBox "Value is larger than 32 bits. Only the lower 32 bits is masked", , "value out of range"
        If bEmpty Then txtHex.Text = "0x" & Format(Hex(curValue), "00000000")
    End If

End Sub

Private Sub Form_Load()
    ' add 31 check box n label
    Dim i As Integer
    For i = 1 To 31
        Load CheckBit(i)
        Load lBit(i)
        CheckBit(i).Visible = True
        lBit(i).Visible = True
    Next i
    
    ' align all controls
    For i = 31 To 0 Step -1
        lBit(i).Caption = i
        CheckBit(i).Value = False
        lBit(i).Left = 225 + (31 - i) * 225
        lBit(i).Top = 225
        CheckBit(i).Left = 225 + (31 - i) * 225
        CheckBit(i).Top = 495
    Next i
    
    Me.Width = lBit(0).Left + lBit(0).Width + 225
    curValue = 0
    txtHex.Text = "0x" & Format(Hex(curValue), "00000000")
    
End Sub

Private Sub txtHex_KeyPress(KeyAscii As Integer)
    Dim key
    key = Chr$(KeyAscii)
    ' check for valid hex code
    If (KeyAscii = 8) Or _
        (txtHex.SelLength <> 0) Then GoTo check_key
    
    If Left(txtHex.Text, 2) = "0x" Then
        If Len(txtHex.Text) = 10 Then
            KeyAscii = 0
            Exit Sub
        End If
    Else    ' no prefix "0x", only allows 8 char
        If Len(txtHex.Text) = 8 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
                
check_key:
    If (key < "0" Or key > "9") And _
        (KeyAscii <> 8) And _
        (key < "a" Or key > "f") And _
        (key < "A" Or key > "F") Then
        KeyAscii = 0
    End If
End Sub
