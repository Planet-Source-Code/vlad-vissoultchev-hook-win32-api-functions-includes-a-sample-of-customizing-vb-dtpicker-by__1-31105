VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2544
   ClientLeft      =   48
   ClientTop       =   288
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2544
   ScaleWidth      =   3744
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   348
      Left            =   840
      TabIndex        =   0
      Top             =   1092
      Width           =   2028
      _ExtentX        =   3577
      _ExtentY        =   614
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.8
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   -2147483635
      CalendarTitleForeColor=   -2147483634
      Format          =   22740993
      CurrentDate     =   37277
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--- can be used to unhook GetLocaleInfoA but not implemented
Private m_pOrigGetLocaleInfoA As Long

Private Sub Form_Load()
    HookImportedFunctionByName _
            GetModuleHandle("mscomct2.ocx"), _
            "KERNEL32.DLL", _
            "GetLocaleInfoA", _
            AddressOf MyGetLocaleInfo, _
            m_pOrigGetLocaleInfoA
End Sub
