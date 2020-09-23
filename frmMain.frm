VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sort Array"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   2925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSort 
      Caption         =   "Sort Array"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Frame frmeSep 
      Height          =   135
      Left            =   0
      TabIndex        =   12
      Top             =   1720
      Width           =   2895
   End
   Begin VB.TextBox txtMyArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   1560
      TabIndex        =   5
      Text            =   "joseph_huntley@email.com"
      Top             =   1430
      Width           =   1335
   End
   Begin VB.TextBox txtMyArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Text            =   "Huntley"
      Top             =   1140
      Width           =   1335
   End
   Begin VB.TextBox txtMyArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   1560
      TabIndex        =   3
      Text            =   "Joseph"
      Top             =   850
      Width           =   1335
   End
   Begin VB.TextBox txtMyArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   1560
      TabIndex        =   2
      Text            =   "by"
      Top             =   570
      Width           =   1335
   End
   Begin VB.TextBox txtMyArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Text            =   "Array"
      Top             =   290
      Width           =   1335
   End
   Begin VB.TextBox txtMyArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Text            =   "Sort"
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblMyArray 
      Caption         =   "MyArray(5):"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   1430
      Width           =   1335
   End
   Begin VB.Label lblMyArray 
      Caption         =   "MyArray(4):"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   1140
      Width           =   1335
   End
   Begin VB.Label lblMyArray 
      Caption         =   "MyArray(3):"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblMyArray 
      Caption         =   "MyArray(2):"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   570
      Width           =   1335
   End
   Begin VB.Label lblMyArray 
      Caption         =   "MyArray(1):"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   290
      Width           =   1335
   End
   Begin VB.Label lblMyArray 
      Caption         =   "MyArray(0):"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************
'*             Sort Array by Joseph Huntley               *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'*                                                        *
'*  Made:  September 29, 1999                             *
'*  Level: Intermediate                                   *
'**********************************************************
'*   The forms here are only used to demonstrate how to   *
'* use the functions 'SortArray' and                      *
'* FirstInAlphabeticalOrder'. You may copy the functions  *
'* into your project for use. If you need any help,       *
'* please e-mail me.                                      *
'**********************************************************
'* Notes: This could be used to sort a listbox instead of *
'*        using the Sorted property.                      *
'**********************************************************

Sub SortArray(strArray() As String)

'**********************************************************
'*             Sort Array by Joseph Huntley               *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'**********************************************************
'*   You may use this code freely as long as credit is    *
'* given to the author, and the header remains intact.    *
'**********************************************************


'--------------------- The Arguments ----------------------
'strArray    - The string array to sort.
'----------------------------------------------------------




  Dim intOut As Integer, intIn As Integer
  Dim strTemp As String
  
    'loop through array
    For intOut% = LBound(strArray()) To UBound(strArray())
        For intIn% = intOut% + 1 To UBound(strArray())
            'check if the inner loop's current dimension is
            'higher precendence, then the outer. If so, swap
            'them.
            If FirstInAlphabeticalOrder(strArray(intOut%), strArray(intIn%)) = 2 Then
               strTemp$ = strArray(intIn%)
               strArray(intIn%) = strArray(intOut%)
               strArray(intOut%) = strTemp$
            End If
        Next intIn%
    Next intOut%
    
End Sub

Function FirstInAlphabeticalOrder(strOne As String, strTwo As String) As Long
   
'**********************************************************
'*     First in Alphabetical Order by Joseph Huntley      *
'*               joseph_huntley@email.com                 *
'*                http://joseph.vr9.com                   *
'**********************************************************
'*   You may use this code freely as long as credit is    *
'* given to the author, and the header remains intact.    *
'**********************************************************


'--------------------- The Arguments ----------------------
'strOne - The first string to compare.
'strTwo - The second string to compare.
'----------------------------------------------------------
'Returns: 0 if strOne and strTwo are at the same level.
'         1 if strOne is at a higher level.
'         2 if strTwo is at a higher level.
   
'Description: Checks to see which of two strings is on a
'             higher level. Alphabetically-wise.
   
   
   Dim intChar As Integer, intLen As Integer
   Dim strChar1 As String, strChar2 As String
   
      'Check to see which string has more length
      'assign intLen% the length of that string.
      If Len(strOne$) > Len(strTwo$) Then
         intLen% = Len(strOne$)
      ElseIf Len(strTwo$) > Len(strOne$) Then
         intLen% = Len(strTwo$)
      Else
         intLen% = Len(strOne$)
      End If
        
   
      For intChar% = 1 To intLen%
        strChar1$ = UCase$(Mid$(strOne$, intChar%, 1))
        strChar2$ = UCase$(Mid$(strTwo$, intChar%, 1))
           
            'if no more character's are left on a string
            'then that string automatically takes precedence.
            'So exit the function.
            If Len(strChar1$) = 0 Then
               FirstInAlphabeticalOrder = 1
               Exit Function
            ElseIf Len(strChar2$) = 0 Then
               FirstInAlphabeticalOrder = 2
               Exit Function
            End If
            
            'if character ascii value is between the ascii
            'value of 'A' and 'Z', and the other character's
            'ascii value is not. Precednce goes to the first
            'string. If that and vice-versa is false. Check
            'which ascii value is lower than the other ascii value.
            'If one if it is lower that string takes precedence. If
            'their equal, continue to the next character.
            
            If Asc(strChar1$) >= Asc("A") And Asc(strChar1$) <= Asc("Z") And Asc(strChar2$) <= Asc("A") And Asc(strChar2$) >= Asc("Z") Then
               FirstInAlphabeticalOrder = 1
               Exit Function
            ElseIf Asc(strChar2$) >= Asc("A") And Asc(strChar2$) <= Asc("Z") And Asc(strChar1$) <= Asc("A") And Asc(strChar1$) >= Asc("Z") Then
               FirstInAlphabeticalOrder = 2
               Exit Function
            ElseIf Asc(strChar1$) < Asc(strChar2$) Then
               FirstInAlphabeticalOrder = 1
               Exit Function
            ElseIf Asc(strChar2$) < Asc(strChar1$) Then
               FirstInAlphabeticalOrder = 2
               Exit Function
            End If
               
      Next intChar%
   
End Function

Private Sub cmdSort_Click()
   
   Dim MyArray(0 To 5) As String 'declare array that we're gona sort
   
   Dim intBuffer As Integer
   
     For intBuffer% = 0 To 5
        MyArray(intBuffer%) = txtMyArray(intBuffer%).Text
     Next intBuffer%
     
   Call SortArray(MyArray())
   
     For intBuffer% = 0 To 5
        txtMyArray(intBuffer%).Text = MyArray(intBuffer%)
     Next intBuffer%
     
   
   
   
   
End Sub


Private Sub Form_Load()

End Sub
