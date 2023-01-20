VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   Icon            =   "DoScores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Sub OutputtoDbgview Lib "kernel32" _
  Alias "OutputDebugStringA" (ByVal lpOutputString As String)

Private Sub Form_Load()
   Form1.Visible = False
 
    ' Arrow Bridge Game Scores
    OutputtoDbgview ("Starting Arrow Bridge Scores")
    xSourceFile = "C:\sbbs\xtrn\Abridge\TOPCHARS.ABR"
    xDestFile = "C:\temp\TOPCHARS.ABR"
    FileCopy xSourceFile, xDestFile
    AnalyzeArrowB (xDestFile)
    Kill (xDestFile)
   
   ' NY2008 Games Scores
    OutputtoDbgview ("Starting NY2008 Scores")
    xSourceFile = "C:\sbbs\xtrn\ny2008\NYSCORES.ASC"
    xDestFile = "C:\temp\NYSCORES.ASC"
    FileCopy xSourceFile, xDestFile
    AnalyzeNY2008 (xDestFile)
    Kill (xDestFile)
  
    ' BRE Games Scores
    OutputtoDbgview ("Starting BRE Scores")
    xSourceFile = "C:\sbbs\xtrn\bre\bulletin\scores.txt"
    xDestFile = "C:\temp\scores.txt"
    FileCopy xSourceFile, xDestFile
    AnalyzeBRE (xDestFile)
    Kill (xDestFile)
    
    ' Clans Games Scores
    OutputtoDbgview ("Starting The Clans Scores")
    xSourceFile = "C:\sbbs\xtrn\clans\scores.asc"
    xDestFile = "C:\temp\scores.asc"
    FileCopy xSourceFile, xDestFile
    AnalyzeClans (xDestFile)
    Kill (xDestFile)
   
    ' LOD Games Scores
    OutputtoDbgview ("Starting LOD Scores")
    xSourceFile = "C:\sbbs\xtrn\LOD\LODRANK.ASC"
    xDestFile = "C:\temp\LODRANK.ASC"
    FileCopy xSourceFile, xDestFile
    AnalyzeLOD (xDestFile)
    Kill (xDestFile)
   
    ' LORD Games Scores
    OutputtoDbgview ("Starting LORD Scores")
    xSourceFile = "C:\sbbs\xtrn\LORD\HISCORE.TXT"
    xDestFile = "C:\temp\HISCORE.TXT"
    FileCopy xSourceFile, xDestFile
    AnalyzeLORD (xDestFile)
    Kill (xDestFile)
    
    '
    '
    '**************************************************************
    '
    '
    
    ' League 10 LORD Games Scores
    OutputtoDbgview ("Starting LORD League 10 Scores")
    xSourceFile = "C:\sbbs\xtrn\LORD10\HISCORE.TXT"
    xDestFile = "C:\temp\HISCORE10.TXT"
    FileCopy xSourceFile, xDestFile
    AnalyzeLORD10 (xDestFile)
    Kill (xDestFile)
    
    ' League 10 BRE Games Scores
    OutputtoDbgview ("Starting BRE League 10 Scores")
    xSourceFile = "C:\sbbs\xtrn\league10\bre10\bulletin\plyscore.txt"
    xDestFile = "C:\temp\bre10.txt"
    FileCopy xSourceFile, xDestFile
    AnalyzeBRE10 (xDestFile)
    Kill (xDestFile)
    
    ' League 10 Clans Games Scores
    OutputtoDbgview ("Starting The Clans League 10 Scores")
    xSourceFile = "C:\sbbs\xtrn\league10\clans10\scores.asc"
    xDestFile = "C:\temp\clans10.asc"
    FileCopy xSourceFile, xDestFile
    AnalyzeClans10 (xDestFile)
    Kill (xDestFile)
    '
    '
    '**************************************************************
    '
    '
    '****************************************************************
    '
    '
    '***************    Trade Wars 2002 ****************************
    '
    ' TW2002 Games Scores
    OutputtoDbgview ("Starting TW2002 Game Scores")
    xSourceFile = "C:\SBBS\XTRN\TW2002\scores\TWTrader.txt"
    xDestFile = "C:\temp\twtrader.txt"
    FileCopy xSourceFile, xDestFile
    AnalyzeTW2002 (xDestFile)
    Kill (xDestFile)
    
    ' TWGS Games Scores
    OutputtoDbgview ("Starting TWGS Game Scores")
    xSourceFile = "D:\Program Files\EIS\TWGS\Game\TW1\OUTPUT\TWTRADER.TXT"
    xDestFile = "C:\temp\TWTRADER.TXT"
    FileCopy xSourceFile, xDestFile
    AnalyzeTWGS (xDestFile)
    Kill (xDestFile)
    
    xSourceFile = "D:\Program Files\EIS\TWGS\Game\TW1\OUTPUT\TWCORP.TXT"
    xDestFile = "C:\temp\TWCORP.TXT"
    FileCopy xSourceFile, xDestFile
    AnalyzeTWGS (xDestFile)
    Kill (xDestFile)
    OutputtoDbgview ("Done with ALL Game Scores")
    End
   
End Sub
Function NormalizeSpaces(s As String) As String

    Do While InStr(s, String(2, " ")) > 0
        s = Replace(s, String(2, " "), " ")
    Loop
    NormalizeSpaces = s

End Function

Public Sub CleanUpLines(DirtyLine As String)
Dim CleanLine As String

Debug.Print "This is the dirty line: " & DirtyLine
CleanLine = Replace(DirtyLine, " ", ".")
' Debug.Print "Clean: " & CleanLine
CleanLine = Replace(CleanLine, "", " ")
Debug.Print "This is a Clean Line: " & CleanLine
' CleanLine = NormalizeSpaces(CleanLine)
' Debug.Print "SQL:    " & CleanLine


End Sub



Public Sub AnalyzeArrowB(fileName As String)
    OutputtoDbgview ("Begin Analyze Arror Bridge")
    Dim intEmpFileNbr As Integer
    
    Lines = ""
    intEmpFileNbr = FreeFile

    Open fileName For Input As #intEmpFileNbr
    Do Until EOF(intEmpFileNbr)
        
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line2
        Input #intEmpFileNbr, Line3
        Input #intEmpFileNbr, Line4
        Lines = Line1 & " " & Line2 & " " & Line3 & " " & Line4
    '    Debug.Print Lines
        
    strArray = Split(Lines, " ")
    sCharName = strArray(0)
    If InStr(LCase(sCharName), "[none]") <> 0 Then
        Close #intEmpFileNbr
        Exit Sub
    End If
    
    sLevel = strArray(1)
    sType = strArray(2)
    sLastPlayed = strArray(3)
    
    'Here is where we save to the mySQL database
   
    Call WriteArrowBtoMySQL
    
    sCharName = ""
    sLevel = ""
    sType = ""
    sLastPlayed = ""
        
    Loop
    
    Close #intEmpFileNbr
    
End Sub
Public Sub AnalyzeNY2008(fileName As String)
    
    OutputtoDbgview ("Begin Analyze NY2008")
    Lines = ""
    intEmpFileNbr = FreeFile

    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    Do Until EOF(intEmpFileNbr)
        
        Input #intEmpFileNbr, Lines
        Lines = NormalizeSpaces(Lines)
        ' Lines = Replace(Lines, " ", ".")
        ' Debug.Print Lines
        
    strArray = Split(Lines, " ")
    
    sRank = strArray(0)
    sLevel = strArray(1)
    sCharName = strArray(2)
    sPoints = strArray(3)
    
    If Not IsNumeric(sPoints) Then
        sCharName = strArray(2) & " " & strArray(3)
        sPoints = strArray(4)
        sSex = strArray(5)
        If sSex = "M" Then
            sSex = "Male"
        Else
            sSex = "Female"
        End If
    
        sType = strArray(6)
        If InStr(sType, "CRACK") <> 0 Then
            sType = "Crack Addict"
        Else
            sType = StrConv(sType, vbProperCase)
        End If
    Else
        sSex = strArray(4)
        If sSex = "M" Then
            sSex = "Male"
        Else
            sSex = "Female"
        End If
        
        sType = strArray(5)
        If InStr(sType, "CRACK") <> 0 Then
            sType = "Crack Addict"
        Else
            sType = StrConv(sType, vbProperCase)
        End If
    End If
        
    
    'Here is where we save to the mySQL database
    Call WriteNY2008toMySQL
    
    sRank = ""
    sLevel = ""
    sCharName = ""
    sPoints = ""
    sSex = ""
    sType = ""
    
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeBRE(fileName As String)
    
    OutputtoDbgview ("Begin Analyze BRE")
    Lines = ""
    intEmpFileNbr = FreeFile

    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    Do Until EOF(intEmpFileNbr)
        
        Input #intEmpFileNbr, Lines
        If InStr(Lines, "ÄÄÄÄÄ") <> 0 Then
            Close #intEmpFileNbr
            Exit Sub
        End If
        
            
        
        Lines = NormalizeSpaces(Lines)
        ' Lines = Replace(Lines, " ", ".")
         Debug.Print Lines
        
        strArray = Split(Lines, " ")
    
        sRank = strArray(0)
        sCharName = strArray(1)
        sType = strArray(2)
        If Not IsNumeric(sType) Then
            sCharName = strArray(1) & " " & strArray(2)
            sType = strArray(3)
            sPoints = strArray(4)
            sLevel = strArray(5)
        Else
            sPoints = strArray(3)
            sLevel = strArray(4)
        End If
            
        
        'Here is where we save to the mySQL database
        Call WriteBREtoMySQL
        
        sRank = ""
        sCharName = ""
        sType = ""
        sPoints = ""
        sLevel = ""
    
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeClans(fileName As String)
    
    OutputtoDbgview ("Begin Analyze The Clans")
    Lines = ""
    intEmpFileNbr = FreeFile

    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    Do Until EOF(intEmpFileNbr)
        Line Input #intEmpFileNbr, Lines
        OutputtoDbgview ("Line to Process: " & Lines)
        sCharName = Mid(Lines, 1, InStr(Lines, "  "))
        sCharName = Replace(sCharName, "'", ".")
        Lines = Mid(Lines, Len(sCharName), Len(Lines))
        Lines = NormalizeSpaces(Lines)
        Lines = LTrim(Lines)
        Lines = Replace(Lines, ", ", ",")
        strArray = Split(Lines, " ")
    
        sType = strArray(0)
        If IsNumeric(sType) Then ' a number, no symbol
            sType = " "
            sPoints = strArray(0)
            sRank = strArray(1)
            If UBound(strArray) >= 2 Then
                sRank = strArray(1) & "  " & strArray(2)
            Else
                ' nada
            End If
        Else
            sPoints = strArray(1)
            sRank = strArray(2)
            If UBound(strArray) >= 3 Then
                sRank = strArray(2) & "  " & strArray(3)
            Else
                ' nada
            End If
        End If
        
        
        'Here is where we save to the mySQL database
         'Debug.Print sCharName & "  " & sType & "  " & sPoints & "  " & sRank
        
         Call WriteClasntoMySQL
        
        sCharName = ""
        sType = ""
        sPoints = ""
        sRank = ""
        
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeLOD(fileName As String)
    
    OutputtoDbgview ("Begin Analyze LOD")
    Lines = ""
    intEmpFileNbr = FreeFile

    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    Do Until EOF(intEmpFileNbr)
        Input #intEmpFileNbr, LineA
        Input #intEmpFileNbr, LineB
        Lines = LineA & "," & LineB
        If InStr(Lines, "Team Name") <> 0 Then
            Close #intEmpFileNbr
            Exit Sub
        End If
        
        sCharName = Mid(Lines, 1, InStr(Lines, "   "))
        
        Lines = Mid(Lines, Len(sCharName), Len(Lines))
        Lines = NormalizeSpaces(Lines)
        Lines = LTrim(Lines)
      '   Lines = Replace(Lines, "  ", ".")
        strArray = Split(Lines, " ")
    
        sPoints = strArray(0)
        sLevel = strArray(1)
        sType = strArray(2)
        sSex = strArray(3)
        sRank = strArray(4)
        
        'Here is where we save to the mySQL database
         'Debug.Print sCharName & "  " & sPoints & "   " & sLevel & "    " & sType & "  " & sSex & "  " & sRank
        
         Call WriteLODtoMySQL
        
        sCharName = ""
        sType = ""
        sPoints = ""
        sRank = ""
        
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeLORD(fileName As String)
    Dim SexClass As String
    Dim tClass  As String
    
    OutputtoDbgview ("Begin Analyze LORD")
    Lines = ""
    intEmpFileNbr = FreeFile
    
    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    
    Do Until EOF(intEmpFileNbr)
        Line Input #intEmpFileNbr, Lines
        
        SexClass = Mid(Lines, 1, 4)
        sSex = Mid(SexClass, 1, 1)
        If sSex = "F" Then
            sSex = "Female"
        Else
            sSex = "Male"
        End If
        
        tClass = Mid(SexClass, 3, 1)
        
        Select Case tClass
            Case "M"
                sClass = "Magician"
            Case "T"
                sClass = "Thief"
            Case "D"
                sClass = "Dark Knight"
            Case Else
                sClass = "None"
        End Select
        
        Lines = Mid(Lines, 5)
        sCharName = Mid(Lines, 1, InStr(Lines, "  "))
        Lines = Mid(Lines, Len(sCharName), Len(Lines))
        Lines = NormalizeSpaces(Lines)
        Lines = LTrim(Lines)
        strArray = Split(Lines, " ")
        sPoints = strArray(0)
        sLevel = strArray(1)
        Lines = Mid(Lines, (Len(sPoints) + Len(sLevel) + 2), Len(Lines))
        If InStr(Lines, "Dead") <> 0 Then
            sRank = "Dead"
            sMastered = Mid(Lines, 1, Len(Lines) - 4)
        Else
            sRank = "Alive"
            sMastered = Mid(Lines, 1, Len(Lines) - 5)
        End If
        
        'Here is where we save to the mySQL database
        Call WriteLORDtoMySQL
        
        'Debug.Print sSex & "  " & sClass & "   " & sCharName & "  " & sPoints & "  " & sLevel & "  " & sMastered & "   " & sRank
        
        sSex = ""
        sClass = ""
        sCharName = ""
        sPoints = ""
        sLevel = ""
        sMastered = ""
        sRank = ""
        
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeLORD10(fileName As String)
    Dim SexClass As String
    Dim tClass  As String
    
    OutputtoDbgview ("Begin Analyze LORD League 10")
    Lines = ""
    intEmpFileNbr = FreeFile

    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    
    Do Until EOF(intEmpFileNbr)
        Line Input #intEmpFileNbr, Lines
        
        SexClass = Mid(Lines, 1, 4)
        sSex = Mid(SexClass, 1, 1)
        If sSex = "F" Then
            sSex = "Female"
        Else
            sSex = "Male"
        End If
        
        tClass = Mid(SexClass, 3, 1)
        
        Select Case tClass
            Case "M"
                sClass = "Magician"
            Case "T"
                sClass = "Thief"
            Case "D"
                sClass = "Dark Knight"
            Case Else
                sClass = "None"
        End Select
        
        Lines = Mid(Lines, 5)
        sCharName = Mid(Lines, 1, InStr(Lines, "  "))
        Lines = Mid(Lines, Len(sCharName), Len(Lines))
        Lines = NormalizeSpaces(Lines)
        Lines = LTrim(Lines)
        strArray = Split(Lines, " ")
        sPoints = strArray(0)
        sLevel = strArray(1)
        Lines = Mid(Lines, (Len(sPoints) + Len(sLevel) + 2), Len(Lines))
        If InStr(Lines, "Dead") <> 0 Then
            sRank = "Dead"
            sMastered = Mid(Lines, 1, Len(Lines) - 4)
        Else
            sRank = "Alive"
            sMastered = Mid(Lines, 1, Len(Lines) - 5)
        End If
        
        'Here is where we save to the mySQL database
        Call WriteLORD10toMySQL
        
        'Debug.Print sSex & "  " & sClass & "   " & sCharName & "  " & sPoints & "  " & sLevel & "  " & sMastered & "   " & sRank
        
        sSex = ""
        sClass = ""
        sCharName = ""
        sPoints = ""
        sLevel = ""
        sMastered = ""
        sRank = ""
        
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeBRE10(fileName As String)
    
    OutputtoDbgview ("Begin Analyze BRE League 10")
    Lines = ""
    intEmpFileNbr = FreeFile
    
    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    OutputtoDbgview ("Before Do Untill to read scores  - BRE League 10")
    Do Until EOF(intEmpFileNbr)
    OutputtoDbgview ("Start of Do Until, to read scores  - BRE League 10")
        
        Input #intEmpFileNbr, Lines
        If InStr(Lines, "       ") = 0 Then
            Close #intEmpFileNbr
            Exit Sub
        End If
                 
        Debug.Print Lines
        
        sRank = Mid(Lines, 1, InStr(Lines, ")") + 1)
        sRank = LTrim(sRank)
        sRank = RTrim(sRank)
        
        Lines = Mid(Lines, InStr(Lines, ")") + 1)
        
        sCharName = Mid(Lines, 1, InStr(Lines, "  "))
        
        Lines = Mid(Lines, Len(sCharName), Len(Lines))
        Lines = NormalizeSpaces(Lines)
        Lines = LTrim(Lines)
        
        sPoints = Mid(Lines, 1, InStr(Lines, " "))
        sLevel = Mid(Lines, InStr(Lines, " "))
        
        
        'Here is where we save to the mySQL database
        Call WriteBRE10toMySQL
        
        sRank = ""
        sCharName = ""
        sPoints = ""
        sLevel = ""
    
    Loop
    
    Close #intEmpFileNbr
    
     On Error GoTo 0
   Exit Sub
MyErrorHandler:
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") at line " & Erl
   
    
End Sub

Public Sub AnalyzeClans10(fileName As String)
    
    OutputtoDbgview ("Begin Analyze The Clans League 10")
    Lines = ""
    intEmpFileNbr = FreeFile

    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    Do Until EOF(intEmpFileNbr)
        Input #intEmpFileNbr, Lines
        Debug.Print Lines
        If InStr(Lines, "No one has played") <> 0 Then
            Close #intEmpFileNbr
            Exit Sub
        End If
        
        sCharName = Mid(Lines, 1, InStr(Lines, "  "))
        sCharName = Replace(sCharName, "'", ".")
        
        Lines = Mid(Lines, Len(sCharName), Len(Lines))
        Lines = NormalizeSpaces(Lines)
        Lines = LTrim(Lines)
      '   Lines = Replace(Lines, "  ", ".")
        strArray = Split(Lines, " ")
    
        sType = strArray(0)
        If IsNumeric(sType) Then ' a number, no symbol
            sType = " "
            sPoints = strArray(0)
            sRank = strArray(1)
            If UBound(strArray) >= 2 Then
                sRank = strArray(1) & "  " & strArray(2)
            Else
                ' nada
            End If
        Else
            sPoints = strArray(1)
            sRank = strArray(2)
            If UBound(strArray) >= 3 Then
                sRank = strArray(2) & "  " & strArray(3)
            Else
                ' nada
            End If
        End If
        
        
        'Here is where we save to the mySQL database
         'Debug.Print sCharName & "  " & sType & "  " & sPoints & "  " & sRank
        
         Call WriteClasn10toMySQL
        
        sCharName = ""
        sType = ""
        sPoints = ""
        sRank = ""
        
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeTW2002(fileName As String)
    
    OutputtoDbgview ("Begin Analyze TW2002")
    ' The format for the twtrader.txt file has changed ;
    ' #       Rank Title      Corp        Trader Name                Ship Type
    
    Lines = ""
    intEmpFileNbr = FreeFile
    Open fileName For Input As #intEmpFileNbr
    Input #intEmpFileNbr, LineA
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    Input #intEmpFileNbr, Line1
    
    Do Until EOF(intEmpFileNbr)
        Line Input #intEmpFileNbr, Lines
        Lines = LTrim(Lines)
        OutputtoDbgview ("   " & Lines)
        If InStr(Lines, "                                ") <> 0 Then
            Close #intEmpFileNbr
            Exit Sub
        End If
        
        sDate = Mid(LineA, InStr(LineA, ":") + 1)
        'OutputtoDbgview ("TW2002 Date    " & sDate)
        
        ' Line number in file, NOT NEEDED
        sLevel = Mid(Lines, 1, InStr(Lines, " "))
        Lines = Mid(Lines, Len(sLevel), Len(Lines))
        sLevel = RTrim(sLevel)
        sLevel = LTrim(sLevel)
        'OutputtoDbgview ("TW2002 Line#   " & sLevel)
        Lines = LTrim(Lines)
        'OutputtoDbgview ("   " & Lines)
        
        'sRank = Rank Title
        sRank = Mid(Lines, 1, InStr(Lines, "  "))
        Lines = Mid(Lines, Len(sRank), Len(Lines))
        sRank = RTrim(sRank)
        sRank = LTrim(sRank)
        ' OutputtoDbgview ("TW2002 Rank   " & sRank)
        Lines = LTrim(Lines)
        
        ' sType = Corporation
        sType = Mid(Lines, 1, InStr(Lines, " "))
        Lines = Mid(Lines, Len(sType), Len(Lines))
        sType = RTrim(sType)
        sType = LTrim(sType)
        ' OutputtoDbgview ("TW2002 Type   " & sType)
        Lines = LTrim(Lines)
        
        'sPoints = Character Name sCharName
        sPoints = Mid(Lines, 1, InStr(Lines, " "))
        Lines = Mid(Lines, Len(sPoints), Len(Lines))
        sPoints = RTrim(sPoints)
        sPoints = LTrim(sPoints)
        sCharName = sPoints
       ' OutputtoDbgview ("TW2002 Points   " & sPoints)
        Lines = LTrim(Lines)
              
       ' sCharName = Mid(Lines, 1, InStr(Lines, "  "))
       ' Lines = Mid(Lines, Len(sCharName), Len(Lines))
       ' sCharName = RTrim(sCharName)
       ' sCharName = LTrim(sCharName)
       ' OutputtoDbgview ("TW2002 CharName   " & sCharName)
        Lines = LTrim(Lines)
       
        Lines = LTrim(Lines)
        sClass = LTrim(Lines)
        sClass = RTrim(Lines)
        OutputtoDbgview ("TW2002 Class   " & sClass)
        
       ' Here is where we save to the mySQL database
       ' NOT USING OutputtoDbgview ("Line Number" & "     " & sLevel)
       ' OutputtoDbgview ("Rank       " & "     " & sRank)
       ' OutputtoDbgview ("Corp       " & "     " & sType)
       ' NOT USEING OutputtoDbgview ("Corp       " & "     " & sPoints)
       ' OutputtoDbgview ("Trader Name" & "     " & sCharName)
       ' OutputtoDbgview ("Ship Type  " & "     " & sClass)
                
        Call WriteTW2002toMySQL
        
        sLevel = ""
        sRank = ""
        sPoints = ""
        sType = ""
        sCharName = ""
        sClass = ""
        
    Loop
    
    Close #intEmpFileNbr
    
End Sub

Public Sub AnalyzeTWGS(fileName As String)
    
    OutputtoDbgview ("Begin Analyze TWGS")
    Lines = ""
    intEmpFileNbr = FreeFile
    If InStr(LCase(fileName), "twtrader") <> 0 Then
        sDestFile = "twtrader.txt"
        Open fileName For Input As #intEmpFileNbr
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line1
        
        Do Until EOF(intEmpFileNbr)
            Line Input #intEmpFileNbr, Lines
            Debug.Print Lines
            
            Lines = LTrim(Lines)
            If InStr(Lines, "                                ") <> 0 Then
                Close #intEmpFileNbr
                Exit Sub
            End If
            
            sLevel = Mid(Lines, 1, InStr(Lines, "  "))
            sLevel = LTrim(sLevel)
            sLevel = RTrim(sLevel)
            
            Lines = Mid(Lines, Len(sLevel) + 1, Len(Lines))
            Lines = LTrim(Lines)
            
            sRank = Mid(Lines, 1, InStr(Lines, "  "))
            sRank = LTrim(sRank)
            sRank = RTrim(sRank)
            
            Lines = Mid(Lines, Len(sRank) + 1, Len(Lines))
            Lines = LTrim(Lines)
            
            sPoints = Mid(Lines, 1, InStr(Lines, " "))
            sPoints = LTrim(sPoints)
            sPoints = RTrim(sPoints)
            
            Lines = Mid(Lines, Len(sPoints) + 1, Len(Lines))
            Lines = LTrim(Lines)
            
            sType = Mid(Lines, 1, InStr(Lines, " "))
            sType = LTrim(sType)
            sType = RTrim(sType)
            
            Lines = Mid(Lines, Len(sType) + 1, Len(Lines))
            Lines = LTrim(Lines)
            
            sCharName = Mid(Lines, 1, InStr(Lines, "  "))
            
            Lines = Mid(Lines, Len(sCharName), Len(Lines))
            Lines = LTrim(Lines)
            
            sClass = LTrim(Lines)
            sClass = RTrim(Lines)
            
            'Here is where we save to the mySQL database
             'Debug.Print sLevel & "    " & sRank & "     " & sPoints & "    " & sType & "     " & sCharName & "  " & sClass
            Call WriteTWGStoMySQL
            
            sLevel = ""
            sRank = ""
            sPoints = ""
            sType = ""
            sCharName = ""
            sClass = ""
    

        Loop
    Else
        sDestFile = "twcorp.txt"
        Open fileName For Input As #intEmpFileNbr
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line1
        Input #intEmpFileNbr, Line1
            
        Do Until EOF(intEmpFileNbr)
            Line Input #intEmpFileNbr, LineA
                
            If Len(LineA) < 1 Then
                Close #intEmpFileNbr
                Exit Sub
            End If
            Lines = LTrim(LineA)
            Line Input #intEmpFileNbr, LineB
            xGarbage = Mid(Lines, 1, InStr(Lines, " "))
             
            Lines = Mid(Lines, Len(xGarbage) + 1, Len(Lines))
            Lines = LTrim(Lines)
            
            sLevel = Mid(Lines, 1, InStr(Lines, " "))
            sLevel = RTrim(sLevel)
            
            Lines = Mid(Lines, Len(sLevel) + 1, Len(Lines))
            
            sCorpName = Mid(Lines, 1, InStr(Lines, "  "))
            
            sCEO = Mid(Lines, Len(sCorpName), Len(Lines))
            sCEO = LTrim(sCEO)
            sCEO = RTrim(sCEO)
            
            Lines = LineB
            sCorpExp = Mid(Lines, InStr(Lines, "e:") + 2, InStr(Lines, "Co"))
            sCorpExp = LTrim(sCorpExp)
            sCorpExp = RTrim(sCorpExp)
            
            sCorpAlign = Mid(Lines, InStr(Lines, "t:") + 2, Len(Lines))
            sCorpAlign = LTrim(sCorpAlign)
            sCorpAlign = RTrim(sCorpAlign)
            
                
            'Here is where we save to the mySQL database
                       
            Call WriteTWGStoMySQL
            sLevel = ""
            sCorpName = ""
            sCEO = ""
            sCorpExp = ""
            sCorpAlign = ""
            
    
        Loop
    End If
            
    Close #intEmpFileNbr
    
End Sub


Public Sub WriteArrowBtoMySQL()

OutputtoDbgview ("Begin Write to SQL Arrow Bridge")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from arrowb WHERE abchar =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from arrowb WHERE abchar =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!abchar = sCharName
    Rs!Level = sLevel
    Rs!Type = sType
    Rs!lastplayed = sLastPlayed
              
    Rs.Update
    Rs.Close
        
Else
        
    If Rs!abchar = sCharName Then
        'no update
    Else
        Rs!abchar = sCharName
    End If
        
    If Rs!Level = sLevel Then
        ' no update
    Else
        Rs!Level = sLevel
    End If
        
    If Rs!Type = sType Then
        ' no update
    Else
        Rs!Type = sType
    End If
    
    If Rs!lastplayed = sLastPlayed Then
        'no update
    Else
        Rs!lasatplayed = sLastPlayed
    End If
    
        
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing
    
End Sub

Public Sub WriteNY2008toMySQL()

OutputtoDbgview ("Begin Write to SQL NY2008")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from ny2008 WHERE charname =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from ny2008 WHERE charname =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!Rank = sRank
    Rs!Level = sLevel
    Rs!charname = sCharName
    Rs!points = sPoints
    Rs!sex = sSex
    Rs!hoodlum = sType
              
    Rs.Update
    Rs.Close
        
Else
        
    If Rs!Rank = sRank Then
        'no update
    Else
        Rs!Rank = sRank
    End If
        
    If Rs!Level = sLevel Then
        'no update
    Else
        Rs!Level = sLevel
    End If
        
    If Rs!charname = sCharName Then
        ' no update
    Else
        Rs!charname = sCharName
    End If
    
    If Rs!points = sPoints Then
        ' no update
    Else
        Rs!points = sPoints
    End If
        
    If Rs!sex = sSex Then
        'no update
    Else
        Rs!sex = sSex
    End If
    
    If Rs!hoodlum = sType Then
        ' no update
    Else
        Rs!hoodlum = sType
    End If
        
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub

Public Sub WriteBREtoMySQL()

OutputtoDbgview ("Begin Write to SQL BRE")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from bre WHERE empirename =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from bre WHERE empirename =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!breid = sRank
    Rs!empirename = sCharName
    Rs!territory = sType
    Rs!score = sPoints
    Rs!networth = sLevel
              
    Rs.Update
    Rs.Close
        
Else
        
    If Rs!breid = sRank Then
        'no update
    Else
        Rs!breid = sRank
    End If
        
    If Rs!empirename = sCharName Then
        'no update
    Else
        Rs!empirename = sCharName
    End If
        
    If Rs!territory = sType Then
        ' no update
    Else
        Rs!territory = sType
    End If
    
    If Rs!score = sPoints Then
        ' no update
    Else
        Rs!score = sPoints
    End If
        
    If Rs!networth = sLevel Then
        'no update
    Else
        Rs!networth = sLevel
    End If
    
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub

Public Sub WriteClasntoMySQL()

OutputtoDbgview ("Begin write to SLQ The Clans")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from clans WHERE clanname =" & "'" & sCharName & "' and symbol =" & "'" & sType & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from clans WHERE clanname =" & "'" & sCharName & "' and symbol =" & "'" & sType & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!clanname = sCharName
    Rs!symbol = sType
    Rs!score = sPoints
    Rs!Status = sRank
              
              
    Rs.Update
    Rs.Close
        
Else
                
    If Rs!clanname = sCharName Then
        'no update
    Else
        Rs!clanname = sCharName
    End If
    
    If Rs!Status = sRank Then
        'no update
    Else
        Rs!Status = sRank
    End If
        
    If Rs!symbol = sType Then
        ' no update
    Else
        Rs!symbol = sType
    End If
    
    If Rs!score = sPoints Then
        ' no update
    Else
        Rs!score = sPoints
    End If
        
    
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub

Public Sub WriteLODtoMySQL()

OutputtoDbgview ("Begin write to SQL LOD")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from lod WHERE name =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from lod WHERE name  =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!Name = sCharName
    Rs!worth = sPoints
    Rs!Level = sLevel
    Rs!quest = sType
    Rs!attr = sSex
    Rs!comp = sRank
              
              
    Rs.Update
    Rs.Close
        
Else
                
    If Rs!Name = sCharName Then
        'no update
    Else
        Rs!Name = sCharName
    End If
    
    If Rs!worth = sPoints Then
        ' no update
    Else
        Rs!worth = sPoints
    End If
    
    If Rs!Level = sLevel Then
        'no update
    Else
        Rs!Level = sLevel
    End If
    
    If Rs!quest = sType Then
        ' no update
    Else
        Rs!quest = sType
    End If
    
    If Rs!attr = sSex Then
        'no update
    Else
        Rs!attr = sSex
    End If
    
    If Rs!comp = sRank Then
        'no update
    Else
        Rs!comp = sRank
    End If
        
    
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub


Public Sub WriteLORDtoMySQL()
OutputtoDbgview ("Begin write to SQL LORD")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from lord WHERE cname =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from lord WHERE cname  =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!sex = sSex
    Rs!ctype = sClass
    Rs!cname = sCharName
    Rs!Experience = sPoints
    Rs!Level = sLevel
    Rs!mastered = sMastered
    Rs!cstatus = sRank
              
    Rs.Update
    Rs.Close
        
Else
                
    If Rs!sex = sSex Then
        'no update
    Else
        Rs!sex = sSex
    End If
    
    If Rs!ctype = sClass Then
        ' no update
    Else
        Rs!ctype = sClass
    End If
    
    If Rs!cname = sCharName Then
        'no update
    Else
        Rs!cname = sCharName
    End If
    
    If Rs!Experience = sPoints Then
        ' no update
    Else
        Rs!Experience = sPoints
    End If
    
    If Rs!Level = sLevel Then
        'no update
    Else
        Rs!Level = sLevel
    End If
    
    If Rs!mastered = sMastered Then
        'no update
    Else
        Rs!mastered = sMastered
    End If
        
    If Rs!cstatus = sRank Then
        'no update
    Else
        Rs!cstatus = sRank
    End If
    
    
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub

Public Sub WriteLORD10toMySQL()
OutputtoDbgview ("Begin write to SQL Lord League 10")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from lord10 WHERE cname =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from lord10 WHERE cname  =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!sex = sSex
    Rs!ctype = sClass
    Rs!cname = sCharName
    Rs!Experience = sPoints
    Rs!Level = sLevel
    Rs!mastered = sMastered
    Rs!cstatus = sRank
              
    Rs.Update
    Rs.Close
        
Else
                
    If Rs!sex = sSex Then
        'no update
    Else
        Rs!sex = sSex
    End If
    
    If Rs!ctype = sClass Then
        ' no update
    Else
        Rs!ctype = sClass
    End If
    
    If Rs!cname = sCharName Then
        'no update
    Else
        Rs!cname = sCharName
    End If
    
    If Rs!Experience = sPoints Then
        ' no update
    Else
        Rs!Experience = sPoints
    End If
    
    If Rs!Level = sLevel Then
        'no update
    Else
        Rs!Level = sLevel
    End If
    
    If Rs!mastered = sMastered Then
        'no update
    Else
        Rs!mastered = sMastered
    End If
        
    If Rs!cstatus = sRank Then
        'no update
    Else
        Rs!cstatus = sRank
    End If
    
    
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub

Public Sub WriteBRE10toMySQL()
OutputtoDbgview ("Begin write to SQL BRE League 10")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from bre10 WHERE empirename =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from bre10 WHERE empirename =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!breid = sRank
    Rs!empirename = sCharName
    Rs!score = sPoints
    Rs!territory = sLevel

              
    Rs.Update
    Rs.Close
        
Else
        
    If Rs!breid = sRank Then
        'no update
    Else
        Rs!breid = sRank
    End If
        
    If Rs!empirename = sCharName Then
        'no update
    Else
        Rs!empirename = sCharName
    End If
        
    If Rs!territory = sLevel Then
        ' no update
    Else
        Rs!territory = sLevel
    End If
    
    If Rs!score = sPoints Then
        ' no update
    Else
        Rs!score = sPoints
    End If
            
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub

Public Sub WriteClasn10toMySQL()
OutputtoDbgview ("Begin write to SQL The Clans League 10")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from clans10 WHERE clanname =" & "'" & sCharName & "' and symbol =" & "'" & sType & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText

Rs.Open "SELECT * from clans10 WHERE clanname =" & "'" & sCharName & "' and symbol =" & "'" & sType & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!clanname = sCharName
    Rs!symbol = sType
    Rs!score = sPoints
    Rs!Status = sRank
              
              
    Rs.Update
    Rs.Close
        
Else
                
    If Rs!clanname = sCharName Then
        'no update
    Else
        Rs!clanname = sCharName
    End If
    
    If Rs!Status = sRank Then
        'no update
    Else
        Rs!Status = sRank
    End If
        
    If Rs!symbol = sType Then
        ' no update
    Else
        Rs!symbol = sType
    End If
    
    If Rs!score = sPoints Then
        ' no update
    Else
        Rs!score = sPoints
    End If
        
    
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub


Public Sub WriteTW2002toMySQL()
OutputtoDbgview ("Beging write to SQL TW2002")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

DBQuery.Open "SELECT Count(*) as RSCount from tw2002 WHERE twtname =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText
OutputtoDbgview ("Just Ran Select Count * ")

Rs.Open "SELECT * from tw2002 WHERE twtname =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText
OutputtoDbgview ("Just ran RS.Open Select * from tw2002 where")

'Executes the query-command and puts the result into Rs (recordset)
If DBQuery!RSCount = 0 Then
        
    Rs.AddNew
    Rs!twdate = sDate
    Rs!twrank = sRank
    Rs!twcorp = sType
    Rs!twtname = sCharName
    Rs!twshiptype = sClass
              
              
    Rs.Update
    Rs.Close
        
Else
    If Rs!twdate = sDate Then
        'no update
    Else
        Rs!twdate = sDate
    End If
    
    If Rs!twrank = sRank Then
        'no update
    Else
        Rs!twrank = sRank
    End If
    
    
    If Rs!twcorp = sType Then
        ' no update
    Else
        Rs!twcorp = sType
    End If
    
    If Rs!twtname = sCharName Then
        ' no update
    Else
        Rs!twtname = sCharName
    End If
    
    If Rs!twshiptype = sClass Then
        ' no update
    Else
        Rs!twshiptype = sClass
    End If
    
    
    Rs.Update
    Rs.Close
                
End If
    DBCon.Close
    Set Rs = Nothing
    Set DBCon = Nothing

End Sub


Public Sub WriteTWGStoMySQL()
OutputtoDbgview ("Begin write to SQL TWGS")
Call ConnecttoMYSQL

Set Rs = New ADODB.Recordset
Rs.CursorLocation = adUseServer
Set DBQuery = New ADODB.Recordset

If LCase(sDestFile) = "twtrader.txt" Then

    DBQuery.Open "SELECT Count(*) as RSCount from twgs WHERE twtname =" & "'" & sCharName & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText
    Rs.Open "SELECT * from twgs WHERE twtname =" & "'" & sCharName & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText
    
    If DBQuery!RSCount = 0 Then
        Rs.AddNew
        Rs!twnumber = sLevel
        Rs!twrank = sRank
        Rs!twalign = sPoints
        Rs!twcorpnum = sType
        Rs!twtname = sCharName
        Rs!twshiptype = sClass
        Rs.Update
        Rs.Close
    Else
        If Rs!twnumber = sLevel Then
            'no update
        Else
            Rs!twnumber = sLevel
        End If
        
        If Rs!twrank = sRank Then
            'no update
        Else
            Rs!twrank = sRank
        End If
        
        If Rs!twalign = sPoints Then
            'no update
        Else
            Rs!twalign = sPoints
        End If
        
        If Rs!twcorpnum = sType Then
            ' no update
        Else
            Rs!twcorpnum = sType
        End If
        
        If Rs!twtname = sCharName Then
            ' no update
        Else
            Rs!twtname = sCharName
        End If
        
        
        Rs.Update
        Rs.Close
                    
    End If
        DBCon.Close
        Set Rs = Nothing
        Set DBCon = Nothing
Else
    DBQuery.Open "SELECT Count(*) as RSCount from twgs WHERE twcorpnum =" & "'" & sLevel & "';", DBCon, adOpenStatic, adLockOptimistic, adCmdText
    Rs.Open "SELECT * from twgs WHERE twcorpnum =" & "'" & sLevel & "';", DBCon, adOpenDynamic, adLockOptimistic, adCmdText
    
    If DBQuery!RSCount = 0 Then
        ' No Matching Corporatoin for User
        Rs.Close
    Else
        If Rs!twceo = sCEO Then
            ' no update
        Else
            Rs!twceo = sCEO
        End If
        
        If Rs!twcorpname = sCorpName Then
            ' no update
        Else
            Rs!twcorpname = sCorpName
        End If
        
        If Rs!twcorpexp = sCorpExp Then
            'no update
        Else
            Rs!twcorpexp = sCorpExp
        End If
        
        If Rs!twcorpalign = sCorpAlign Then
            'no update
        Else
            Rs!twcorpalign = sCorpAlign
        End If
        
        Rs.Update
        Rs.Close
                    
    End If
        DBCon.Close
        Set Rs = Nothing
        Set DBCon = Nothing
End If

End Sub
