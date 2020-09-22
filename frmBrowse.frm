VERSION 5.00
Begin VB.Form frmBrowse1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Source Excel File"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   ControlBox      =   0   'False
   Icon            =   "frmBrowse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   415
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.FileListBox FilFile 
      Height          =   2040
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
   End
   Begin VB.DirListBox DirDirectory 
      Height          =   1890
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   5280
      Width           =   1095
   End
End
Attribute VB_Name = "frmBrowse1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

'This hides the Browser form and shows
'the Settings form.
frmBrowse1.Hide
frmConvert.Show

End Sub

Private Sub cmdOK_Click()

' Store selected path and file name in the SelectedFile variable
' and adds a backslash to end of path if one does not exist and
'is needed.
If FilFile.FileName = "" Then
   msg = "Please choose a file"
   title = "File Error!"
   style = vbOKOnly
   Response = MsgBox(msg, style, title)
   Exit Sub
Else
    If Right(FilFile.Path, 1) = "\" Then
        selectedfile = FilFile.Path + FilFile.FileName
    Else
        selectedfile = FilFile.Path + "\" + FilFile.FileName
    End If
End If

frmConvert.txtExcel = selectedfile

frmBrowse1.Hide
frmConvert.Show

End Sub

Private Sub DirDirectory_Change()

'This changes the File list to match
'the current directory.
FilFile.FileName = DirDirectory.Path
FilFile.Refresh

End Sub

Private Sub drvDrive_Change()

'Changes the directory window to match
'the current drive.
DirDirectory.Path = drvDrive.Drive
DirDirectory.Refresh
FilFile.Refresh

End Sub

Private Sub Form_Load()

'Sets the default drive to C: and the default
'location of starting directory.  Also sets the
'default file pattern to *.XLS.
FilePattern = "*.xls"
FilFile.Pattern = FilePattern + ";*XLS"
drvDrive.Drive = "C:"
DirDirectory = "C:\"

End Sub
