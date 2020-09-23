VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Disco Lights"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   118
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   591
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "save"
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   6480
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.FileListBox flbFiles 
      ForeColor       =   &H80000007&
      Height          =   1260
      Left            =   6480
      Pattern         =   "*.txt"
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox picEdit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   4
      Top             =   600
      Width           =   6255
   End
   Begin VB.HScrollBar scrPreview 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   6255
   End
   Begin ComctlLib.Slider sldInterval 
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   327682
      LargeChange     =   1
      Min             =   1
      Max             =   40
      SelStart        =   4
      Value           =   4
   End
   Begin VB.Timer tmrChange 
      Interval        =   500
      Left            =   6000
      Top             =   120
   End
   Begin VB.Label lblSpeed 
      Caption         =   "0.2 sec"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Interval"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************
'* Author: Josh Duck   *
'* Date  : NOV 2002    *
'***********************
Dim num_init As Boolean, cap_init As Boolean, scr_init As Boolean
Dim seqStep
Dim seqString
Dim selRow

Dim changeInterval
Public Sub LoadPattern(file As String)
   On Error GoTo ERR_HANDLER:
   Open file For Input As #1
      Input #1, seqString
  Close #1
  picEdit.Cls
  DrawSelector
  DrawPattern

Exit Sub
ERR_HANDLER:
    MsgBox "Error Opening file" & file
    Exit Sub
End Sub
Public Sub SavePattern(file As String)
   On Error GoTo ERR_HANDLER:
   Open file For Output As #1
      Write #1, seqString;
  Close #1

Exit Sub
ERR_HANDLER:
    MsgBox "Error writing to file" & file
    Exit Sub
End Sub
Public Sub SetLights(num_on As Boolean, cap_on As Boolean, scr_on As Boolean)
   If KeyState(num) <> num_on Then Call keynum
   If KeyState(cap) <> cap_on Then Call keycap
   If KeyState(scr) <> scr_on Then Call keyscr
End Sub
Public Sub DrawPattern()
   Dim i As Integer
   For i = 1 To Len(seqString)
      DrawLight (i)
   Next i
   picEdit.Refresh
End Sub
Public Sub DrawRow(start As Integer)
   DrawLight (start * 3 + 1)
   DrawLight (start * 3 + 2)
   DrawLight (start * 3 + 3)
End Sub

Public Sub DrawLight(light As Integer)
   Dim x As Integer, y As Integer
   x = ((light - 1) \ 3) * 15 + 7 - (scrPreview.Value * 15)
   y = ((light - 1) Mod 3) * 15 + 7
      
   If (x > -5) And x < picEdit.Width + 5 Then
      picEdit.DrawWidth = 10
      If Mid(seqString, light, 1) = "1" Then
         picEdit.PSet (x, y), vbGreen
      Else
         picEdit.PSet (x, y), vbWhite
      End If
      picEdit.DrawWidth = 1
      picEdit.Circle (x, y), 5, vbBlack, False
   End If
End Sub
Public Sub HideSelector()
   If selRow <> -1 Then
      picEdit.DrawWidth = 1
      picEdit.Line (0, 45)-(picEdit.Width, 50), vbWhite, BF         'Cover the whole bottom bit, just to be sure
   End If
End Sub
Public Sub DrawSelector()
   If selRow <> -1 Then
      picEdit.DrawWidth = 1
      picEdit.Line ((selRow - scrPreview.Value) * 15, 45)-((selRow + 1 - scrPreview.Value) * 15, 50), RGB(0, 0, 100), BF
   End If
End Sub
Public Sub DeleteRow()
   If selRow <> -1 And Len(seqString) > (selRow * 3) Then                    'Do we have a row selected at all?
      Dim tempString As String
      tempString = Left(seqString, selRow * 3) & Right(seqString, Len(seqString) - (selRow * 3 + 3))
      seqString = tempString
   
      If Len(seqString) <= selRow * 3 Then
         selRow = Len(seqString) \ 3 - 1
      End If
      picEdit.Cls
      DrawPattern
      DrawSelector
      picEdit.Refresh
      
      ReDoScrollBar
   End If
End Sub
Public Sub InsertRow()
    If Len(seqString) <= selRow * 3 Or selRow = -1 Then     'We selected a row that doesnt exists or there is no selection
      seqString = seqString & "000"                         'Just tack it on to the end
      DrawRow (Len(seqString) \ 3 - 1)
  Else
      Dim tempString As String
      tempString = Left(seqString, selRow * 3) & "000" & Right(seqString, Len(seqString) - selRow * 3)
      seqString = tempString
      picEdit.Cls
      DrawPattern
      DrawSelector
      picEdit.Refresh
   End If
   ReDoScrollBar
End Sub
Public Sub ReDoScrollBar()
   Dim scrollamount
   scrollamount = Len(seqString) \ 3 - (picEdit.Width \ 15)
   If scrollamount < 0 Then scrollamount = 0
   scrPreview.Max = scrollamount
End Sub
Public Sub ChangeLight(x As Integer, y As Integer)
   Dim light As Integer
   HideSelector
   selRow = x \ 15 + scrPreview.Value
   DrawSelector
   If y < 45 Then                          'We clicked on a note
       light = (x \ 15 + scrPreview.Value) * 3 + (y \ 15) + 1
   
      If light <= Len(seqString) Then       'Clicked past the end
          If (Mid(seqString, light, 1) = "1") Then
             Mid(seqString, light, 1) = "0"
          Else
            Mid(seqString, light, 1) = "1"
            End If
         DrawLight (light)
     End If
       
  End If
  picEdit.Refresh
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub cmdSave_Click()
   If txtName.Text <> "" And seqString <> "" Then
      SavePattern File_Path & Replace(Replace(Replace(txtName.Text, "\", "_"), "/", "_"), ".", "_") & ".txt"
   End If
      flbFiles.Refresh
      ' flbFiles.Path = File_Path

End Sub

Private Sub flbFiles_Click()
   LoadPattern (File_Path & flbFiles.FileName)
   txtName.Text = Left(flbFiles.FileName, Len(flbFiles.FileName) - 4)
   ReDoScrollBar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDelete Then
      DeleteRow
   ElseIf KeyCode = vbKeyInsert Then
      InsertRow
   End If
End Sub
Private Sub Form_Load()
   seqString = ""
   seqStep = 0
   changeInterval = 20
   selRow = -1
   DrawPattern
   ReDoScrollBar
   If Right(App.Path, 1) = "/" Or Right(App.Path, 1) = "\" Then     'Will occur in root dir
      File_Path = App.Path & "patterns\"
   Else
      File_Path = App.Path & "\patterns\"
   End If
   flbFiles.Path = File_Path
   LoadPattern (File_Path & "default.txt")

   
   num_init = KeyState(num)
   cap_init = KeyState(cap)
   scr_init = KeyState(scr)
   SetLights False, False, False
End Sub
Private Sub Form_Unload(Cancel As Integer)
   SetLights num_init, cap_init, scr_init
End Sub
Private Sub picEdit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   ChangeLight Int(x), Int(y)
End Sub

Private Sub scrPreview_Change()
   picEdit.Cls
   DrawSelector
   DrawPattern
End Sub

Private Sub sldInterval_Change()
   lblSpeed = sldInterval / 20 & " sec"
   changeInterval = sldInterval * 50
   tmrChange.Interval = changeInterval
End Sub
Private Sub tmrChange_Timer()
   Dim on1 As Boolean, on2 As Boolean, on3 As Boolean
   on1 = Mid(seqString, seqStep * 3 + 1, 1) = "1"
   on2 = Mid(seqString, seqStep * 3 + 2, 1) = "1"
   on3 = Mid(seqString, seqStep * 3 + 3, 1) = "1"
   SetLights on1, on2, on3
   
   seqStep = seqStep + 1
   If (seqStep + 1) * 3 > Len(seqString) Then
      seqStep = 0
   End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(LCase(Chr(KeyAscii)))
End Sub
