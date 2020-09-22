VERSION 5.00
Begin VB.Form frmBrowse2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose Destination Directory"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   Icon            =   "frmBrowse2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   3840
      Width           =   1095
   End
   Begin VB.DirListBox DirDirectory 
      Height          =   2790
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "frmBrowse2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()

'This hides the Browser form and shows
'the Settings form.
frmBrowse2.Hide
frmConvert.Show

End Sub

Private Sub cmdOK_Click()

'This section hides the Browse form and
'and shows the Setings form
frmConvert.txtAccess = DirDirectory.Path
frmBrowse2.Hide
frmConvert.Show

End Sub

Private Sub drvDrive_Change()

'Changes the directory window to match
'the current drive.
DirDirectory.Path = drvDrive.Drive
DirDirectory.Refresh


End Sub

Private Sub Form_Load()

'Sets the default drive to C: and the
'default location of starting directory.
drvDrive.Drive = "C:"
DirDirectory = "C:\"

End Sub
