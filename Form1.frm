VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "yArch"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8040
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboBuddy 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   0
      Width           =   3015
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8493
      _Version        =   393217
      ScrollBars      =   2
      OLEDropMode     =   1
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The other day I remembered I never did try to figure out how to decode the
'timestamps. I figured it would have already been mentioned somewhere else and
'if it is, I couldn't find it! I did come across some software (only one!) that
'offered decoding timestamps.. and it's only free for about 20 uses?! Well, this
'might "spring to life" some more.

'Oh, and this isn't meant to be "pretty" or "efficient". It's meant to demonstrate
'how to do what it's supposed to do. I'm not about to play "dress up" with a form
'to satisfy some people. :)

'Credit would be nice for a change.

'-coozzzzz


' Developed by SumoRicky

Option Explicit

Private Sub EnumUsers()
Dim sPath As String, sCur As String
    sPath = "C:\Program Files\Yahoo!\Messenger\Profiles\"
    sCur = Dir(sPath, vbDirectory)
    cboUser.Clear
    Do Until sCur = ""
        If sCur <> "." And sCur <> ".." And sCur <> "Archive" Then
            If IsDir(sPath & sCur & "\Archive") Then
                cboUser.AddItem sCur
            End If
        End If
        sCur = Dir
    Loop
End Sub

Private Sub EnumBuddies(sUser As String)
Dim sPath As String, sCur As String
    sPath = "C:\Program Files\Yahoo!\Messenger\Profiles\" & sUser & "\Archive\Messages\"
    sCur = Dir(sPath, vbDirectory)
    cboBuddy.Clear
    Do Until sCur = ""
        If sCur <> "." And sCur <> ".." And sCur <> "Archive" Then
            cboBuddy.AddItem sCur
        End If
        sCur = Dir
    Loop
End Sub

Private Sub EnumLogs(sUser As String, sBuddy As String)
Dim sPath As String, sCur As String
    sPath = "C:\Program Files\Yahoo!\Messenger\Profiles\" & sUser & "\Archive\Messages\" & sBuddy & "\"
    sCur = Dir(sPath, vbDirectory)
    cboDate.Clear
    Do Until sCur = ""
        If sCur <> "." And sCur <> ".." And sCur <> "Archive" Then
            If FileLen(sPath & sCur) > 50 Then cboDate.AddItem sCur
        End If
        sCur = Dir
    Loop
End Sub

Private Function GetYahooLog(sFile As String, sUser As String) As String
Dim strBuff As String, strArr() As String, lngLoop As Long
Dim strTime As String, strMsg As String, blnFrom As Boolean
    'REPLACE FILE NAME BELOW
    'REPLACE FILE NAME BELOW
    If LoadFile(sFile, strBuff) = True Then
        Me.Caption = "yArch - Loading " & sFile
        strArr = Split(strBuff, String(3, vbNullChar))
        For lngLoop = 0 To (UBound(strArr) - 1) Step 4
            'the below will avoid any errors but will also ignore a lot of messages
            'because of it. some archives have additional null-characters (if I
            'recall correctly) which this demonstration doesn't take into account.
            'This is just as a "head's up" in case you notice this.
            If (lngLoop + 3 <= UBound(strArr)) Then
                If Len(strArr(lngLoop + 2)) > 0 And Len(strArr(lngLoop + 3)) > 0 And Len(strArr(lngLoop)) = 6 Then
                    'read DecodeTime()!!!! time may be different based on your
                    'Time Zone!!
                    strTime = DecodeTime(Mid(strArr(lngLoop), 2, 4))
                    blnFrom = (Asc(strArr(lngLoop + 2)) = 0)
                    strMsg = DecodeMessage(strArr(lngLoop + 3), cboUser)
                    'have all of the information we need.. just add it to RTB...
                    With Rtb
                        .SelColor = IIf(blnFrom = True, vbBlue, vbRed)
                        .SelBold = True
                        .SelText = strTime & vbCrLf
                        .SelBold = False
                        .SelColor = IIf(blnFrom = True, vbBlue, vbRed)
                        .SelText = strMsg & vbCrLf & vbCrLf
                        .SelStart = Len(.Text)
                    End With
                End If
            End If
        Next lngLoop
        
        Me.Caption = "yArch - Loaded " & sFile
    End If
End Function

Private Function IsDir(sPath As String) As Boolean
On Error GoTo Hell
    If (GetAttr(sPath) And vbDirectory) = vbDirectory Then IsDir = True
Exit Function
Hell:
    IsDir = False
End Function

Private Sub CboBuddy_Click()
    EnumLogs cboUser.Text, cboBuddy.Text
End Sub

Private Sub cboDate_Click()
    Rtb = ""
    Call GetYahooLog("C:\Program Files\Yahoo!\Messenger\Profiles\" & cboUser & "\Archive\Messages\" & cboBuddy & "\" & cboDate, cboUser)
End Sub

Private Sub CboUser_Click()
    EnumBuddies cboUser.Text
End Sub

Private Sub Form_Load()
    Call EnumUsers
End Sub
Private Function LoadFile(ByVal strFile As String, ByRef strBuff As String) As Boolean
    'This will just buffer the contents of the file.
    Dim intFF As Integer
    If Dir$(strFile, vbNormal) <> vbNullString Then
        intFF = FreeFile
        Open strFile For Binary As intFF
            strBuff = Space$(LOF(intFF))
            Get #intFF, 1, strBuff
            strBuff = Mid(strBuff, 20)
        Close intFF
        LoadFile = True
    End If
End Function
Private Function DecodeMessage(ByVal strDecode, ByVal strUser As String) As String
    'The basic XOR method was mentioned in my previous submission of
    'viewing archives. It's not required here.
    Dim intLen1 As Integer, intLen2 As Integer
    Dim intLoop As Integer, intCnt As Integer
    intLen1 = Len(strDecode)
    intLen2 = Len(strUser)
    If (intLen1 > 0) And (intLen2 > 0) Then
        intCnt = 1
        For intLoop = 1 To intLen1
            DecodeMessage = DecodeMessage & Chr(Asc(Mid(strDecode, intLoop, 1)) Xor Asc(Mid(strUser, intCnt, 1)))
            intCnt = intCnt + 1
            If intCnt > intLen2 Then
                intCnt = 1
            End If
        Next intLoop
    End If
End Function
Private Function DecodeTime(ByRef strTime As String) As Date
    Dim lngDateBase As Date
    Dim lngDateDiff As Long
    Dim lngDateTimeZone As Long
    Dim lngSeconds As Double
    'This is how the date in the archive is compared. It's the number of
    'seconds from the time of that message to 01/01/1970. I was dreading
    'that it may be the number of nano-seconds since 1601 which is an
    'extremely large number to deal with in VB! (64-bit?)
    lngDateBase = DateSerial(1970, 1, 1)
    'I was getting overflows using DateAdd() for some reason so I'll subtract
    'the seconds from 1970 to 1990 and add them seperately to the final date.
    'The below is the amount of seconds (should be 100% accurate!!)
    lngDateDiff = 631152000
    'The dates are stored within GMT (+0) and I'm glad for that. I was wondering
    'where calculations went wrong by 5 hours!!! So, since I'm in EST (which is
    '-5 GMT), I'll use -5. Change this to whatever.
    lngDateTimeZone = -5
    'Alright.. The time-stamps (which include the dates) are stored in series
    'of 4 bytes. From the left-to-right, the decimals of those bytes represent
    'the number of seconds. Since the maximum of each byte is 255, they range
    'from 0->255. For example, if the first decimal is "127" that means 127
    'seconds have elapsed. Once it exceeds 255, the next character's
    'decimal will increase. So, if the first 2 characters are 127 and 116, they
    'would be equal to ((127) + (116 * 256)) seconds. Now, since there are for
    'characters, you'll end up having to "256*256" and so forth because of the
    'way the seconds increment.
    lngSeconds = Asc(Mid$(strTime, 1, 1))
    lngSeconds = lngSeconds + (Asc(Mid$(strTime, 2, 1)) * 256#)
    lngSeconds = lngSeconds + (Asc(Mid$(strTime, 3, 1)) * (256# ^ 2))
    lngSeconds = lngSeconds + (Asc(Mid$(strTime, 4, 1)) * (256# ^ 3))
    lngSeconds = lngSeconds - lngDateDiff
    DecodeTime = DateAdd("s", lngSeconds, lngDateBase)
    DecodeTime = DateAdd("s", lngDateDiff, DecodeTime)
    DecodeTime = DateAdd("h", lngDateTimeZone, DecodeTime)
End Function

Private Sub Rtb_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Hell
Dim sUser As String
    sUser = Mid(Data.Files(1), InStrRev(Data.Files(1), "\") + 1)
    sUser = Mid(sUser, InStr(sUser, "-") + 1)
    sUser = Mid(sUser, 1, Len(sUser) - 4)
    GetYahooLog Data.Files(1), sUser
Exit Sub
Hell:
End Sub
