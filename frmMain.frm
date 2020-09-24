VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "MP3 Information"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8550
   ScaleWidth      =   9855
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSubDirs 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   600
      Width           =   9615
   End
   Begin VB.TextBox txtFilePath 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "d:\"
      Top             =   120
      Width           =   7935
   End
   Begin VB.TextBox txtFiles 
      Height          =   5655
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   2760
      Width           =   9615
   End
   Begin VB.CommandButton cmdListFiles 
      Caption         =   "List MP3s"
      Height          =   375
      Left            =   8160
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdListFiles_Click()
    On Error GoTo ErrHand
    
    Dim TempString As String
    Dim X As Long
    Dim TempSubDirs As Variant
    Dim TempFiles As Variant
    
    'create instance of clsMP3
    Dim ObjMP3 As clsMP3
    Set ObjMP3 = New clsMP3
    
    If ObjMP3.SearchDir(txtFilePath.Text, "*.mp3") Then
        'make array of subdirs
        TempSubDirs = Split(ObjMP3.Subdirs, "|")
        
        'loop through TempSubDirs and print the subdirs
        txtSubDirs.Text = "The following Sub dirs were found:" & vbCrLf
        For X = 0 To UBound(TempSubDirs) - 1
            txtSubDirs.Text = txtSubDirs.Text & TempSubDirs(X) & vbCrLf
        Next
        
        'make array of files
        TempFiles = Split(ObjMP3.Files, "|")
        
        'loop through TempArray and read each mp3
        txtFiles.Text = "MP3 data:" & vbCrLf
        For X = 0 To UBound(TempFiles) - 1
        
            If ObjMP3.ReadMP3(TempFiles(X)) Then
                DoEvents
                'concatenate the info in tempstring
                TempString = ObjMP3.Album & "-" & _
                    ObjMP3.Artist & "-" & _
                    ObjMP3.BitRate & "-" & _
                    ObjMP3.Comment & "-" & _
                    ObjMP3.Duration & "-" & _
                    ObjMP3.Frequency & "-" & _
                    ObjMP3.Genre & "-" & _
                    ObjMP3.Mode & "-" & _
                    ObjMP3.MpegLayer & "-" & _
                    ObjMP3.MpegVersion & "-" & _
                    ObjMP3.Songname & "-" & _
                    ObjMP3.Track & "-" & _
                    ObjMP3.Year & "-"
                    
                    'print it to txtFiles
                    txtFiles.Text = txtFiles.Text & TempString & vbCrLf
            End If
        Next
    End If
    
    MsgBox "Done!", vbInformation
    Exit Sub
    
ErrHand:
    MsgBox Err.Number & " " & Err.Description, vbCritical
    Resume Next
End Sub
