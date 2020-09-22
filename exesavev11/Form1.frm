VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Shimoon EXESave 1.1"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "New Username"
      Height          =   3015
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   4095
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Enter"
         Default         =   -1  'True
         Height          =   855
         Left            =   1200
         TabIndex        =   4
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Product key:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Real name:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label lblKey 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Product Key:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblName 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Real name:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblUser 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MainFile As String
Dim AlterFile As String
Dim TempStr As Byte
Dim strUser As String
Dim strName As String
Dim strKey As String
Dim OffsetStep As Integer
Dim Record As Integer

Private Sub Command1_Click()
Dim TempByte As Byte
Dim UserData(20) As Byte
OffsetStep = 40

MainFile = App.Path & "\" & App.EXEName & ".exe"

AlterFile = App.Path & "\" & App.EXEName & "1.exe"

Open MainFile For Binary Access Read As #1
Open AlterFile For Binary Access Write As #2

For i = 1 To LOF(1)
   Get #1, i, TempByte
   Put #2, i, TempByte
Next i

Close #1

For i = 1 To Len(Text1.Text)
    UserData(i) = Asc(Mid(Text1.Text, i, 1))
Next i

If Len(Text1.Text) < 20 Then
    For i = Len(Text1.Text) + 1 To 20
        UserData(i) = Asc(" ")
    Next i
End If

For i = 1 To 20
    OffsetStep = OffsetStep - 1
    Record = LOF(2) - OffsetStep
    Put #2, Record, UserData(i)
Next i

'NOW FOR REAL NAME
OffsetStep = 60
For i = 1 To Len(Text2.Text)
    UserData(i) = Asc(Mid(Text2.Text, i, 1))
Next i

If Len(Text2.Text) < 20 Then
    For i = Len(Text2.Text) + 1 To 20
        UserData(i) = Asc(" ")
    Next i
End If

For i = 1 To 20
    OffsetStep = OffsetStep - 1
    Record = LOF(2) - OffsetStep
    Put #2, Record, UserData(i)
Next i


'NOW FOR KEY

OffsetStep = 80
For i = 1 To Len(Text3.Text)
    UserData(i) = Asc(Mid(Text3.Text, i, 1))
Next i

If Len(Text3.Text) < 20 Then
    For i = Len(Text3.Text) + 1 To 20
        UserData(i) = Asc(" ")
    Next i
End If

For i = 1 To 20
    OffsetStep = OffsetStep - 1
    Record = LOF(2) - OffsetStep
    Put #2, Record, UserData(i)
Next i

Close #2

Dim strBat As String

strBat = "del " & App.EXEName & ".exe" & vbCrLf & "ren " & App.EXEName & "1.exe" & " " & App.EXEName & ".exe" & vbCrLf & App.EXEName & ".exe"



Open App.Path & "\runbat.bat" For Output As #1
Print #1, strBat
Close #1

Shell App.Path & "\runbat.bat", vbHide
End

End Sub

Private Sub Form_Load()
MainFile = App.Path & "\" & App.EXEName & ".exe"

AlterFile = App.Path & "\" & App.EXEName & "1.exe"

On Error GoTo 10

Kill App.Path & "\runbat.bat"




Open MainFile For Binary Access Read As #1
OffsetStep = 40

'get username
For i = 1 To 20
OffsetStep = OffsetStep - 1
Record = LOF(1) - OffsetStep
Get #1, Record, TempStr
strUser = strUser & Chr(TempStr)
Next i

'get real name
OffsetStep = 60
For i = 1 To 20
OffsetStep = OffsetStep - 1
Record = LOF(1) - OffsetStep
Get #1, Record, TempStr
strName = strName & Chr(TempStr)
Next i

'get key
OffsetStep = 80
For i = 1 To 20
OffsetStep = OffsetStep - 1
Record = LOF(1) - OffsetStep
Get #1, Record, TempStr
strKey = strKey & Chr(TempStr)
Next i
Close #1

'display data
lblUser.Caption = strUser
lblName.Caption = strName
lblKey.Caption = strKey
Exit Sub


10:
Open MainFile For Binary Access Read As #1
OffsetStep = 40

'get username
For i = 1 To 20
OffsetStep = OffsetStep - 1
Record = LOF(1) - OffsetStep
Get #1, Record, TempStr
strUser = strUser & Chr(TempStr)
Next i

'get real name
OffsetStep = 60
For i = 1 To 20
OffsetStep = OffsetStep - 1
Record = LOF(1) - OffsetStep
Get #1, Record, TempStr
strName = strName & Chr(TempStr)
Next i

'get key
OffsetStep = 80
For i = 1 To 20
OffsetStep = OffsetStep - 1
Record = LOF(1) - OffsetStep
Get #1, Record, TempStr
strKey = strKey & Chr(TempStr)
Next i

Close #1

'display data
lblUser.Caption = strUser
lblName.Caption = strName
lblKey.Caption = strKey

End Sub

