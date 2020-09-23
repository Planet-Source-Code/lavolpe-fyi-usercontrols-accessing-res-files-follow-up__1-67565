VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000016&
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   435
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl12 
      Height          =   1965
      Left            =   3690
      TabIndex        =   2
      Top             =   2505
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   3466
      ResFile         =   "C:\Program Files\Microsoft Visual Studio\VB98\Projects\Alpha Button RAC\test.RES"
      ResID           =   "103LAVOLPE"
      ResSect         =   "CUSTOM"
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   2970
      Left            =   480
      TabIndex        =   0
      Top             =   1875
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   5239
      ResFile         =   "C:\Program Files\Microsoft Visual Studio\VB98\Projects\Alpha Button RAC\test.RES"
      ResID           =   "109LAVOLPE"
      ResSect         =   "CUSTOM"
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      Height          =   1725
      Left            =   255
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Me.BackColor = CLng(Rnd * vbWhite)
End Sub

