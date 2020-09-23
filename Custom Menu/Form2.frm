VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   1005
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   1770
   LinkTopic       =   "Form2"
   ScaleHeight     =   1005
   ScaleWidth      =   1770
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuFile 
      Caption         =   "File"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'======================================================================'
'| Created By: §e7eN                                                  |'
'| Description: This will allow you to make custom menus              |'
'|              useing Images.                                        |'
'|                                                                    |'
'|                                                                    |'
'| Contact: hate_114@hotmail.com                                      |'
'|                                                                    |'
'| *If you wish to use this in one of your Programs please E-mail me* |'
'======================================================================

Private Sub Form_Load()
Load Form1
    
With Form1
    .CreateMenuButton "New"
    .CreateMenuButton "Open..."
    .CreateMenuButton "Save"
    .CreateMenuButton "Save As..."
    .CreateMenuButton "-"
    .CreateMenuButton "Page Setup"
    .CreateMenuButton "Print"
    .CreateMenuButton "-"
    .CreateMenuButton "Exit"
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Form1.ShowMenu
End If
End Sub

Private Sub MnuFile_Click()
Form1.ShowMenu
End Sub
