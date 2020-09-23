VERSION 5.00
Object = "*\Aaqua_button.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin button_aqua.UserControl1 UserControl12 
      Height          =   405
      Left            =   1260
      TabIndex        =   1
      Top             =   1410
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin button_aqua.UserControl1 UserControl11 
      Height          =   405
      Left            =   1380
      TabIndex        =   0
      Top             =   570
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
      Caption         =   "CIdea - Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl11_Click()
    MsgBox "this works"
End Sub

Private Sub UserControl11_DblClick()
    MsgBox "so does this"
End Sub
