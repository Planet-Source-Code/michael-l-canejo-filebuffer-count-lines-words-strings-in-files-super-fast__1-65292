VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cBench As New cBenchMark

Private Sub Form_Initialize()
    cBench.Start
    
    Dim lngCountLines As Long
    lngCountLines = ReadFile_Buffer_Lines(App.Path & "\one_hundred_thousand_items.txt")
    'NOTE: Change the file path to the file you wish to read, i zipped up
    'two test files, one with 100,000 lines and one with 1,000,000 lines.
    
    cBench.Finish
    
    MsgBox "This file contains: " & lngCountLines & " lines. " & vbNewLine & "Total time elapsed: " & cBench.ElapsedTime
    End
End Sub


Public Function ReadFile_Buffer_Lines(sSource As String, _
    Optional iMemSizeMB As Double = 1.5) As Long


    
    Dim iRead As Integer
    Dim lFilePos As Long
    Dim lFileBuffer As Long
    Dim lFileLength As Long

    Dim strBuffer As String
    
    lFilePos = 1
    lFileBuffer = iMemSizeMB * 1048576 '1 MB buffering
    
    iRead = FreeFile
    
    On Error GoTo FileError
    
    Open sSource For Binary Access Read As #iRead
 
        lFileLength = LOF(iRead) 'get length of file
        
        Do While (lFilePos <= lFileLength)
        
            If (lFilePos + lFileBuffer) > lFileLength Then
             lFileBuffer = lFileLength 'if length of file is more than the 1MB buffer, reset to the file length instead
            End If
            
            strBuffer = Space(lFileBuffer) 'Allocate space in the string

            Get #iRead, lFilePos, strBuffer
    
            lFilePos = lFilePos + Len(strBuffer) 'position in file to read from
           
            ReadFile_Buffer_Lines = ReadFile_Buffer_Lines + Strings_Found(strBuffer, vbNewLine) 'count how many lines found in buffered data

            DoEvents 'uncomment this for faster performance, but it will hang as a result if the file is HUGEEE
        
        Loop
    
    Close #iRead
    
    ReadFile_Buffer_Lines = ReadFile_Buffer_Lines + 1 'add extra line because the first line is ignored

    Exit Function
    
FileError:
 
 MsgBox Err.Number & " - " & Err.Description
    
End Function


Public Function Strings_Found(strString As String, strFind As String) As Long
'Counts sub-strings found in a string

    Dim lngPos As Long
    
    On Error Resume Next
    
    Do
        lngPos = InStr(lngPos + Len(strFind), strString, strFind)
        If lngPos = 0 Then Exit Do
        Strings_Found = Strings_Found + 1
    Loop

End Function
