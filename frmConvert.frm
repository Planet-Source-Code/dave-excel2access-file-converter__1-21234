VERSION 5.00
Begin VB.Form frmConvert 
   Caption         =   "Convert Excel Workbooks to Access"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "frmConvert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Program"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import Workbook"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton cmdBrowse2 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtAccess 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   3735
   End
   Begin VB.TextBox txtDBName 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtExcel 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label lblDBLocal 
      Caption         =   "Directory to Create Access Database:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblDBName 
      Caption         =   "Access File Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblImport 
      Caption         =   "Choose Excel File to Import:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()

'This hides the frmConvert form and
'shows the Excel file browse form frmBrowse1.
frmConvert.Hide
frmBrowse1.Show

End Sub

Private Sub cmdBrowse2_Click()

'This hides the frmConvert form and shows
'the Access destination directory browse
'form frmBrowse2.
frmConvert.Hide
frmBrowse2.Show

End Sub

Private Sub cmdExit_Click()

'Unloads all forms and ends the program.
Unload frmBrowse1
Unload frmBrowse2
Unload Me
End

End Sub

Private Sub cmdImport_Click()

'This section performs the actual import of the
'Excel Workbook into an Access Database and creates
'the Access file if one does not exist.

On Error Resume Next

'This section makes sure all text boxes are
'filled out properly.
If txtExcel.Text = "" Then
    MsgBox "Please select an Excel file to import!", vbOKOnly, "Warning!"
Exit Sub
End If

If txtDBName.Text = "" Then
    MsgBox "Please type in a name for the NEW Access database.", vbOKOnly, "Warning!"
Exit Sub
End If

If txtAccess.Text = "" Then
    MsgBox "Please select a destination directory for your Access database file.", vbOKOnly, "Warning!"
Exit Sub
End If

frmConvert.MousePointer = vbHourglass

'This section creates the database file so
'the import can take place and looks for
'duplicate filenames.
Dim db As Database
Dim newDB As String
Dim xl As Excel.Application
Dim sht As Excel.Worksheet
Dim timex As Integer

If Right(txtAccess.Text, 1) = "\" Then
   newDB = txtAccess.Text + txtDBName + ".mdb"
Else
   newDB = txtAccess.Text + "\" + txtDBName + ".mdb"
End If

If Dir(newDB) <> "" Then
    Kill newDB
End If
Set db = dao.CreateDatabase(newDB, dbLangGeneral)
db.Close
Set db = Nothing

For x = 1 To 30000
 x = x + 1
Next x

'Special thanks goes out to Jennifer Campion for this
'bit of code.  I could not have completed this without her help.
'This code actually does the importing of the data from
'Excel into Access.
Dim ac As Access.Application
Set ac = New Access.Application
ac.OpenCurrentDatabase newDB, False

Set xl = New Excel.Application
xl.Workbooks.Open txtExcel.Text

For Each sht In ActiveWorkbook.Worksheets
    If sht.Range("A1").CurrentRegion.Count <> 1 Then
       ac.DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel97, "tbl" & sht.Name, txtExcel.Text, True, sht.Name & "!" & sht.Range("A1").CurrentRegion.Address(False, False)
    End If
Next sht
    
'This closes all active Access and Excel applications.
xl.ActiveWorkbook.Close
xl.Quit
Set xl = Nothing
ac.Quit
Set ac = Nothing
Call clear

frmConvert.MousePointer = vbDefault
MsgBox "Import is complete!", vbOKOnly, "Notice:"

End Sub

