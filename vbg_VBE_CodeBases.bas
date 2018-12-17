Attribute VB_Name = "vbg_VBE_CodeBases"
Option Explicit
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
' Author: nahan@vba.guru
' Risk: As Is 100% Unsupported.
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='

Sub DealWithMultiDims(strModuleName As String)
    
    Dim lLineNum As Long
    Dim lNewLineNum As Long
    Dim aDims As Variant
    Dim lDimIndex As Long
    'On Error Resume Next
    Dim sThisLine As String
    Dim sNextLine As String
    
    With ThisWorkbook.VBProject
        With .VBComponents(strModuleName).CodeModule
            For lLineNum = 1 To .CountOfLines
                sThisLine = Trim(.Lines(lLineNum, 1)) ' 1 Line Block from the Index
                If (InStr(1, sThisLine, "Dim ") = 1) And (InStr(1, sThisLine, ",") >= 5) Then ' If Trimmed Line Starts with "Dim " and has a Comma that's Bad :(
                'sPreviousLine = .Lines(lLineNum - 1, 1) ' 1 Line Block from the Index Less 1
                    sThisLine = Replace(sThisLine, "Dim ", "")
                    Debug.Print sThisLine
                    .DeleteLines lLineNum, 1 ' Delete Dud Force Variant Multi Declarations.
                    aDims = Split(sThisLine, ",")
                    For lDimIndex = 0 To UBound(aDims)
                        sNextLine = "    Dim " & Trim(aDims(lDimIndex))
                        Debug.Print sNextLine
                        lNewLineNum = lLineNum + lDimIndex
                        .InsertLines lNewLineNum, sNextLine
                    Next lDimIndex
                'If Trim(sThisLine) = "" And Trim(sPreviousLine) = "" Then
                    
                End If
            Next lLineNum
        End With
    End With
End Sub

Sub DeleteDoubleBlankLinesInVBE(strModuleName As String)
    
    Dim lLineNum As Long
    On Error Resume Next
    Dim sThisLine As String
    Dim sPreviousLine As String
    
    With ThisWorkbook.VBProject
        With .VBComponents(strModuleName).CodeModule
            For lLineNum = .CountOfLines To 2 Step -1 ' Got's To Go Backwards when Live Deleting. Don't Care about Line 1 Should never be Blank Anyway.
                sThisLine = .Lines(lLineNum, 1) ' 1 Line Block from the Index
                sPreviousLine = .Lines(lLineNum - 1, 1) ' 1 Line Block from the Index Less 1
                Debug.Print sThisLine & " v " & sPreviousLine
                If Trim(sThisLine) = "" And Trim(sPreviousLine) = "" Then
                    .DeleteLines lLineNum, 1 ' Delete one of the Pair of Basically Blank Lines.
                End If
            Next lLineNum
        End With
    End With
End Sub

Sub TEST_DeleteDoubleBlankLines()

    'Call DeleteDoubleBlankLinesInVBE("BTool")
    Call DeleteDoubleBlankLinesInVBE("Builder")
    Call DeleteDoubleBlankLinesInVBE("Main")
    

End Sub

Sub TEST_MultiDims()

    'Call DealWithMultiDims("Module1")
    Call DealWithMultiDims("Builder")
    Call DealWithMultiDims("Main")

End Sub
