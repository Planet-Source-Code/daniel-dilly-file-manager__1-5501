VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Manager"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdir 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   3855
   End
   Begin VB.ListBox filelist 
      Height          =   5520
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3855
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   6000
      TabIndex        =   1
      Top             =   360
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   6000
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblfile 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function FillFileList()
filelist.Clear
If Dir1.Path = "C:\" Then
 For x = 0 To Dir1.ListCount - 1
 newdir = "<" + Right(Dir1.List(x), Len(Dir1.List(x)) - Len(Dir1.Path)) + ">"
 filelist.AddItem newdir
 Next x
Else
 filelist.AddItem "<..>"
 For x = 0 To Dir1.ListCount - 1
 newdir = "<" + Right(Dir1.List(x), Len(Dir1.List(x)) - Len(Dir1.Path) - 1) + ">"
 filelist.AddItem newdir
 Next x
End If
For x = 1 To File1.ListCount - 1
filelist.AddItem File1.List(x)
Next x
txtdir.Text = Dir1.Path
End Function


Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub filelist_Click()
If filelist.Text = "<..>" Then
 lblfile.Caption = ""
 Exit Sub
End If
x = InStr(1, filelist.Text, "<")
If x = 1 Then
 y = InStr(1, filelist.Text, ">")
 z = Mid(filelist.Text, x + 1, y - 2)
 If Dir1.Path = "C:\" Then
  lblfile.Caption = Dir1.Path + z + "\"
 Else
  lblfile.Caption = Dir1.Path + "\" + z + "\"
 End If
Else
 If Dir1.Path = "C:\" Then
  lblfile.Caption = Dir1.Path + filelist.Text
 Else
  lblfile.Caption = Dir1.Path + "\" + filelist.Text
 End If
End If
End Sub

Private Sub filelist_DblClick()
x = InStr(1, filelist.Text, "<")
If x = 1 Then
 y = InStr(1, filelist.Text, ">")
 z = Mid(filelist.Text, x + 1, y - 2)
 If Dir1.Path = "C:\" Then
  Dir1.Path = Dir1.Path + z
 Else
  Dir1.Path = Dir1.Path + "\" + z
 End If
 FillFileList
End If
End Sub

Private Sub Form_Load()
Dir1.Path = "C:\"
FillFileList
End Sub

