VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MIDI Example"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   2010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   510
      Left            =   990
      TabIndex        =   1
      Top             =   0
      Width           =   1005
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   510
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPlay_Click()
PlayMID "music" 'pretty obvious really
End Sub

Private Sub cmdStop_Click()
StopMID "music" 'pretty obvious really
End Sub

Private Sub Form_Load()
'Open a .mid file ready to play...
Dim MidiFile As String
MidiFile = "C:\Music.mid"
OpenMID MidiFile, "music"
'The alias is the reference used when attempting to stop, close it
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Close the .mid
CloseMID "music"
End Sub



