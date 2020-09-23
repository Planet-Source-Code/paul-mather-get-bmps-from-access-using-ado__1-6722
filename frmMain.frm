VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPhoto 
      AutoSize        =   -1  'True
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   5115
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim dbConnectionString As String
Dim dbPath As String
Dim tempImageFile As String
Dim rsObj As New ADODB.Recordset
Dim returnedImage

    dbPath = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "Images.mdb"
    dbConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & dbPath & ";DefaultDir=" & App.Path & ";"
    tempImageFile = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & App.Title & ".bmp"
    
    Call rsObj.Open("SELECT * FROM [Images];", dbConnectionString)
    If Not rsObj.BOF Then
        Do While Not rsObj.EOF
            Debug.Print DisplayBitmap(rsObj.Fields("Photo"), tempImageFile)
            Set picPhoto.Picture = LoadPicture(tempImageFile)
            rsObj.MoveNext
        Loop
    End If
    On Error Resume Next
    Kill tempImageFile
End Sub

Private Sub picPhoto_Resize()
    Me.Width = picPhoto.Width + picPhoto.Left + 200
    Me.Height = picPhoto.Height + picPhoto.Top + 500
End Sub
