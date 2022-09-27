VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FRename 
   Caption         =   "Rename File Extension in Bulk"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox FileList 
      Height          =   1065
      Left            =   150
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton CReplace 
      Caption         =   "Find And Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2265
      Width           =   2775
   End
   Begin VB.TextBox TReplace 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1545
      Width           =   4695
   End
   Begin VB.TextBox TFind 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   1065
      Width           =   4695
   End
   Begin RichTextLib.RichTextBox RTF 
      Height          =   1740
      Left            =   585
      TabIndex        =   5
      Top             =   2835
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   3069
      _Version        =   393217
      TextRTF         =   $"FRename.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1065
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Replace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1545
      Width           =   1215
   End
End
Attribute VB_Name = "FRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CReplace_Click()
    FileList.Path = App.Path
    FileList.Refresh 'REFRESHES FILE LIST
    i = 0
    While i < FileList.ListCount
        If ((InStr(1, FileList.List(i), TFind.Text) > 0)) Then
            RTF.LoadFile App.Path & "\" & FileList.List(i), vbCFText
            RTF.SaveFile (App.Path & "\" & FileList.List(i) & TReplace.Text), rtfText
            
        End If
        i = i + 1
    Wend
    MsgBox "Replaced All !", vbInformation
End Sub
