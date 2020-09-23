VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bit Mapper"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Choose a New Color"
      Height          =   495
      Left            =   5760
      TabIndex        =   5
      Top             =   3840
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3015
      Left            =   120
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   541
      TabIndex        =   3
      Top             =   720
      Width           =   8175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Code"
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Bitmap"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   3120
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "You can set the transparent color here:"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "This program takes a bitmap of your choice and saves the code for generating it at runtime in a text file."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Picture1.Picture = LoadPicture(CommonDialog1.FileName)
Picture1.Refresh

End Sub

Private Sub Command2_Click()
    Open App.Path & "\code.txt" For Output As #2
        Dim i As Integer
        Dim j As Integer
        Print #2, "' This is the code for Psetting your bitmap."
        Print #2, "' All you have to do is replace the word 'canvas' with the name"
        Print #2, "' of the control onto which you are drawing."
        Do Until i = Picture1.ScaleWidth
            j = 0
            Do Until j = Picture1.ScaleHeight
                'This is where we get the picture
                Dim CurrColor As Long
                CurrColor = GetPixel(Picture1.hdc, i, j)
                If CurrColor <> Shape1.FillColor Then
                    Dim CurrLine As String
                    CurrLine = "Canvas.pSet(" & i & "," & j & "), " & CurrColor
                    Print #2, CurrLine
                End If
                j = j + 1
            Loop
            i = i + 1
        Loop
    Close #2
    MsgBox "Code has been outputted to " & App.Path & "\Code.txt"
    
End Sub

Private Sub Command3_Click()
    CommonDialog1.ShowColor
    Shape1.FillColor = CommonDialog1.Color
End Sub
