VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InStr vs. InBArr (by Merri)"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmTest.frx":0000
      Left            =   120
      List            =   "frmTest.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   120
      Width           =   4455
   End
   Begin VB.Frame Frame2 
      Caption         =   "InBArr"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "InStr"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         Caption         =   "InBArr"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label3 
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "InBArrRev"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "InBArrRev"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         Caption         =   "InStrRev"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TestFile As String, Iterations As Long, KeyWord As String
Private Sub Combo1_Click()
    Label1 = vbNullString
    Label2 = vbNullString
    Label3 = vbNullString
    Label4 = vbNullString
End Sub
Private Sub Command1_Click()
    Dim Buffer() As Byte, FileNumber As Byte
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    Open TestFile For Binary Access Read As #FileNumber
        'resize buffer
        ReDim Preserve Buffer(LOF(FileNumber) - 1)
        'read data
        Get #FileNumber, , Buffer
    Close #FileNumber
    Start
    For A = 1 To Iterations
        Result = InBArrRev(Buffer, KeyWord, , Compare)
    Next A
    Label1 = Finish & " ms : " & Result
End Sub
Private Sub Command2_Click()
    Dim Buffer As String, FileNumber As Byte
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    Open TestFile For Input As #FileNumber
        'read data
        Buffer = Input(FileLen(TestFile), FileNumber)
    Close #FileNumber
    Start
    For A = 1 To Iterations
        Result = InStrRev(Buffer, KeyWord, , Compare)
    Next A
    Label2 = Finish & " ms : " & Result
End Sub
Private Sub Command3_Click()
    Dim Buffer() As Byte, FileNumber As Byte, FileLength As Long
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    Open TestFile For Binary Access Read As #FileNumber
        FileLength = LOF(FileNumber)
        'resize buffer
        ReDim Preserve Buffer(FileLength - 1)
        'read data
        Get #FileNumber, , Buffer
    Close #FileNumber
    Start
    For A = 1 To Iterations
        Result = InBArr(Buffer, KeyWord, , Compare)
    Next A
    Label3 = Finish & " ms : " & Result
End Sub
Private Sub Command4_Click()
    Dim Buffer As String, FileNumber As Byte, FileLength As Long
    Dim A As Long, Result As Long, Compare As Byte
    Compare = Combo1.ListIndex
    'open file for input
    FileNumber = FreeFile
    FileLength = FileLen(TestFile)
    Open TestFile For Input As #FileNumber
        'read data
        Buffer = Input(FileLength, FileNumber)
    Close #FileNumber
    Start
    For A = 1 To Iterations
        Result = InStr(1, Buffer, KeyWord, Compare)
    Next A
    Label4 = Finish & " ms : " & Result
End Sub
Private Sub Form_Load()
    Combo1.ListIndex = 0
    TestFile = "c:\autoexec.bat"
    KeyWord = "select"
    Iterations = 100000
End Sub
