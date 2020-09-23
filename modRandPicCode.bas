Attribute VB_Name = "modGen"
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
Private Const FontToUse1 As String = "Arial"
Private Const FontToUse2 As String = "Verdana"
Private Const FontToUse3 As String = "Times New Roman"
Private Const FontToUse4 As String = "Abadi MT Condensed"

'Purpose:   Generates a Random string
'Writen By: Simon Chambers (sic_uk) - Email: sic_uk@yahoo.co.uk
'Comment:   I couldn't find any examples that did this - so i made my own! 8-)
'           ASCII range going to be used:
'           48-57 Numbers               (Range=9)
'           65-90 Uppercase Letters     (Range=25)
'           97-122 Lowercase Letters    (Range=25)

Function GenerateCode(Length As Integer) As String
Dim work As Integer
Dim i As Integer

On Error Resume Next    ' Ignore any errors to prevent the program from exiting prematurely

    'Check to see if the length is actually specified
    If Length <= 0 Then Exit Function
    
    'Clear the string variable
    GenerateCode = ""
    
    'Loop round and round until required length is achieved
    For i = 1 To Length
        
        'Ensure the seed is random
        Randomize Timer * Int(42 * Rnd)

        'Generate a Random Number to determine number, uppercase, lowercase letter
        Select Case Int((3 * Rnd) + 1)
            Case 1  'A number is going to be generated
                work = Int((9 * Rnd) + 48)              'Generate a number between 48-57
                GenerateCode = GenerateCode + Chr(work) 'Concatante the ASCII value of the number on the end
            Case 2  ' A uppercase letter is going to be generated
                work = Int((25 * Rnd) + 65)             'Generate a number between 65-90
                GenerateCode = GenerateCode + Chr(work) 'Concatante the ASCII value of the number on the end
            Case 3  ' A lowercase letter is going to be generated
                work = Int((25 * Rnd) + 97)             'Generate a number between 97-122
                GenerateCode = GenerateCode + Chr(work)
        End Select
        
        'Clean up variables
        work = Empty
    Next i

    'Finished so lets getta outta here
End Function

'Purpose:       To generate a Picture from the code generator
'Written By:    Simon Chambers (sic_uk) - Email: sic_uk@yahoo.co.uk
'Comment:       This uses the fonts specifed by the variables FontToUse? - where ? is a number

Sub GeneratePicture(PicBox As PictureBox, Code As String, CharacterSpacing As Integer, StartPos As Integer)
Dim work As String
Dim i As Long

On Error Resume Next    ' Ignore any errors to prevent the program from exiting prematurely
    
    'Set the starting position on the picturebox
    PicBox.CurrentY = StartPos
    PicBox.CurrentX = StartPos
    
    'Loop through each character
    For i = 1 To Len(Code)

        'Workout the new character position
        PicBox.CurrentX = PicBox.CurrentX + CharacterSpacing
        
        'Ensure a random seed is chosen
        Randomize Timer * Int(23 * Rnd)
        
        'Change the font randomly out of one of the four constants
        Select Case Int((4 * Rnd) + 1)
            Case 1
                PicBox.FontName = FontToUse1
            Case 2
                PicBox.FontName = FontToUse2
            Case 3
                PicBox.FontName = FontToUse3
            Case 4
                PicBox.FontName = FontToUse4
        End Select
        
        'Create a random size font - set between 12 and 16
        PicBox.FontSize = Int((4 * Rnd) + 12)
        
        'Determine if itallicaised or not
        If Int((2 * Rnd) + 1) = 2 Then PicBox.FontItalic = True
        
        'Determine if bold or not
        If Int((2 * Rnd) + 1) = 2 Then PicBox.FontBold = True
        
        'Stick the new character into the picturebox control using an accient command!
        PicBox.Print Mid(Code, i, 1);

    Next i

End Sub

'Purpose:       To further confuse the OCR translator, we are going to draw random lines.
'Written By:    Simon Chambers (sic_uk) - Email: sic_uk@yahoo.co.uk
'Comment:       - The maximum number of lines is only the maximum!
'               - The colour lines are red, simply because the eye can't recieve red as well
'                 as other colours. Any other colour can saterate the eye and make the text
'                 hard to read by the eye. However a computer can see all colours equally
'                 and hence would be affected by any colour!
Sub DrawSomeLines(PicBox As PictureBox, Optional MaxNumberofLines As Integer)
Dim i As Integer
Dim xa, xb As Integer
Dim ya, yb As Integer

    'Ensure a random seed is chosen
    Randomize Timer * Int(18 * Rnd)
    
    If MaxNumberofLines <= 0 Then MaxNumberofLines = 10
    
    ' Start a loop - the maximum is random
    For i = 0 To Int((MaxNumberofLines * Rnd) + 1)
        'Ensure a random seed is chosen
        Randomize Timer * Int(18 * Rnd)
        
        'Create random points to draw
        xa = Int((700 * Rnd) + 1)
        ya = Int((700 * Rnd) + 1)
        
        xb = Int((700 * Rnd) + 100)
        yb = Int((700 * Rnd) + 100)
        
        'Draw a Line using the points specified - the colour is going to be red because
        'the human eye can see less red than any other colour!
        PicBox.Line (xa, ya)-(xb, yb), RGB(255, 0, 0)
    
        'loop round to the next loop
    Next i

End Sub
