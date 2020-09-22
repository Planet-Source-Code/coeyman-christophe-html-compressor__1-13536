VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decoment 1.0"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2040
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4470
   ScaleWidth      =   2040
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check5 
      Caption         =   "delete right Tab"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CheckBox Check4 
      Caption         =   "delete left Tab"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "delete right blank"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CheckBox Check2 
      Caption         =   "delete left blank"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "delete white line"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   2025
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "Form1.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim commentaire As String
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' Définit les constantes (à partir de WIN32API.TXT).
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40

Private Sub Form_Load()

SetWindowPos hwnd, conHwndTopmost, 0, 0, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, _
            conSwpNoActivate Or conSwpShowWindow
commentaire = "Decoment coded by Christophe COEYMAN                   "
Label1 = commentaire
End Sub

Private Sub Image1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer
Dim param1 As String
Dim param2 As String
Dim string1 As String
Dim string2 As String
Dim pos As Integer
With Data
    For i = 1 To .Files.Count
        Open .Files.Item(i) For Input As #1
        Open .Files.Item(i) + ".cnv" For Output As #2
            Do While (EOF(1) <> True)
                Line Input #1, string1
                pos = 1
                    Do While ((pos <> Len(string1)))
                        If (InStr(pos, UCase(string1), "//") <> 0) Then
                               If ((InStr(pos, UCase(string1), "//") <> (InStr(pos, UCase(string1), "HTTP://") + 5)) Or (InStr(pos, UCase(string1), "HTTP://") = 0)) Then
                                    string2 = Mid$(string1, 1, InStr(pos, UCase(string1), "//") - 1)
                                    pos = Len(string1)
                               
                               Else
                                    pos = (InStr(pos, UCase(string1), "HTTP://") + 7)
                                    string2 = ""
                               End If
                               
                        Else
                            
                            If (InStr(pos, UCase(string1), "/*") <> 0) Then
                                Do While (InStr(pos, UCase(string1), "*/") = 0)
                                    Line Input #1, string1
                                        
                                Loop
                                pos = Len(string1)
                            Else
                                string2 = string1
                                pos = Len(string1)
                                
                            End If
                        End If
                    Loop
                 
                 If (Len(string1) = 1) Then
                        string2 = string1
                 End If
                If (Check3.Value = 1) Then
                    string2 = RTrim$(string2)
                End If
                If (Check2.Value = 1) Then
                    string2 = LTrim$(string2)
                End If
                
                If (Check4.Value = 1) Then
                    pos = 0
                    Do
                        pos = pos + 1
                    Loop Until (Mid$(string2, pos, 1) <> Chr$(9))
                    
                    string2 = Mid$(string2, pos, Len(string2) - pos + 1)
                
                End If
                
                If (Check5.Value = 1) Then
                    pos = Len(string2) + 1
                    Do
                        pos = pos - 1
                        If (pos = 0) Then
                            Exit Do
                        End If
                    Loop Until ((Mid$(string2, pos, 1) <> Chr$(9)))
                    
                    string2 = Mid$(string2, 1, Len(string2) - (Len(string2) - pos))
                
                End If
                
                
                If (Check1.Value = 0) Then
                        Print #2, string2
                Else
                   If (Len(string2) <> 0) Then
                        Print #2, string2
                 End If
                End If
                string1 = ""
                string2 = ""
            
            
            Loop
        Close #2
        Close #1
        
    Next i
End With
End Sub

Private Sub Timer1_Timer()
commentaire = Mid$(commentaire, 2, Len(commentaire) - 1) + Mid$(commentaire, 1, 1)
Label1 = commentaire
End Sub
