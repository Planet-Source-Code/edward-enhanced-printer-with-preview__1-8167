VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enhanced Print with Preview Test"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   5640
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtTest 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1815
      Index           =   2
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   3255
   End
   Begin VB.TextBox txtTest 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1215
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   6615
   End
   Begin VB.TextBox txtTest 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Default         =   -1  'True
      Height          =   315
      Left            =   4440
      TabIndex        =   0
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox txtTest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   3480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2760
      Width           =   3255
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ==============================================================
' Project:      Enhance Print object Test Project
' Author:       edward moth
' Copyright:    © 2000 qbd software ltd
'
' ==============================================================
' Module:       frmTest
' Purpose:      Test qcPrinter Class
' ==============================================================

Option Explicit

Dim qPrint As New qcPrinter

Private Sub cmdPreview_Click()

Dim iCount As Integer

With qPrint
  .ResetItems
  .ScaleMode = eCentimetre
  .MarginBottom = 2
  .MarginTop = 1
  .MarginLeft = 1
  .MarginRight = 1
End With
For iCount = 0 To 3
  With txtTest(iCount)
    qPrint.AddText .Text, .FontName, .FontSize, .FontBold, .FontItalic, .FontUnderline, .ForeColor
  End With
Next
With qPrint
' Change the Alignment values
  .TextItem(1).Alignment = eCentre
  .TextItem(2).Alignment = eJustify
  .TextItem(3).Alignment = eRight
  .TextItem(4).Alignment = eJustify
' Change the Indent values
  .TextItem(4).IndentLeft = 1
  .TextItem(4).IndentRight = 13
' Show print preview
  .Preview
End With

End Sub


Private Sub cmdClose_Click()

Set qPrint = Nothing
Unload Me
End

End Sub

Private Sub Form_Load()
txtTest(0).Text = "edward moth and" & vbCrLf & "qbd software present" & vbCrLf & vbCrLf

txtTest(1).Text = "An enhanced printer control that includes a Preview function.  In this example the text in" _
   & " each of these text boxes will be printed with the same attributes as the box.  Although this text" _
   & " box cannot display justified text, it will be justified when printed." & vbCrLf & "Each TextItem added to the qcPrinter class can" _
   & " be altered or removed and can contain a number of paragraphs.  The TextItem is used to hold text" _
   & " that has common attributes." & vbCrLf & vbCrLf
txtTest(2).Text = "Right justified" & vbCrLf & "handy for" & vbCrLf & "addresses" & vbCrLf & vbCrLf
txtTest(3).Text = "Hey edward, do that funky disco biscuit thing ... hehehe ... gin soaked boys rule" & vbCrLf & vbCrLf _
   & " This text will have a left indent of 1cm and a right indent of 13cm added to its attributes."

End Sub

