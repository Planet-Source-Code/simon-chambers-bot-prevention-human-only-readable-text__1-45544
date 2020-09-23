VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generate Code"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLength 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "5"
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtCode 
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton CmdGen 
      Caption         =   "&Generate"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2655
   End
   Begin VB.PictureBox picCode 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      ScaleHeight     =   585
      ScaleWidth      =   2625
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Length of Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Code as picture:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Random Code:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Purpose:       To show how to create a picture that would confusing any automatic bot's
'Written by:    Simon Chambers (sic_uk) - Email: sic_uk@yahoo.co.uk
'Comments:      I'm unsure if there any patents on this technique, so if anyone has a
'               deffinate answer to this then email straight away @: sic_uk@yahoo.co.uk
'
'               Legal: You may use this however you wish, along's it does not affect me.
'               You will take full responsiblity for the actions of this program. This header
'               MUST ALWAYS stay intact if you are to use this source code.
'
'               Feel free to use it any of your program's, also i would most greatly
'               appriciate if you put my name somewhere in the documentation that you
'               used my code! :-) (its just being courteous anyway!)

Option Explicit

' Change these at will to experiment!!
Const DefaultSpacing As Integer = 15
Const DefaultStartPos As Integer = 100

'Purpose:       To Simply generate a Code
'Written by:    Simon Chambers (sic_uk) - Email: sic_uk@yahoo.co.uk
'Comments:      I used call on my own subroutines just to make it clearer!

Private Sub CmdGen_Click()
        
    'Clear the picCode PictureBox
    picCode.Cls
    
    'Change the colour to White Text on Blue Background (easier to read!)
    picCode.BackColor = RGB(0, 0, 255)
    picCode.ForeColor = RGB(255, 255, 255)
    
    'Create a code (with its length obtained from txtLength converted into a Double)
    'and put it into txtCode.text
    txtCode.Text = GenerateCode(Val(txtLength.Text))
    
    'Generate the Picture into picCode, using the code from txtCode and the default
    'spacings and starting posistions
    Call GeneratePicture(picCode, txtCode.Text, DefaultSpacing, DefaultStartPos)
    
    'Draw some lines to confuse any OCR program.
    Call DrawSomeLines(picCode)
End Sub
