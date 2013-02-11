VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Configure"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSound 
      Caption         =   "Emulate Sound"
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   1020
      Width           =   2055
   End
   Begin VB.CheckBox chkEGBC 
      Caption         =   "Emulate GameBoy Color"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
