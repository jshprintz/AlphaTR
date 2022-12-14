VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TR_Checklist 
   Caption         =   "TR Checklist"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   OleObjectBlob   =   "TR_Checklist.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TR_Checklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private mCurrentRng, mSpaceRange As Range
Private Response As Variant
Private i, mPCount, mVarCntOne As Integer
Private mRWCount, mHeadlineStr As String
Private mWild As Boolean

Private Sub Cleanup_Click()
    mWild = False
    
    Call NoBorders
    Call OneSpace
    Call ReplaceQuotes
    Call RepWords
    Call Warning_Check
    
    ActiveDocument.CheckSpelling
    
    MsgBox ("Cleanup ran succesfully.")
    
    Unload Me
    

End Sub

Private Sub RepWords()
    
    'ALPHABETICAL LISTING OF COMMON TYPOS
    
        
    mRWCount = mRWCount + ReplaceWords(Chr(150), "-")    ' en dash
    mRWCount = mRWCount + ReplaceWords(Chr(151), "-")    ' em dash
    
    mRWCount = mRWCount + ReplaceWords("1 percent ", "1%")
    mRWCount = mRWCount + ReplaceWords("2 percent ", "2%")
    mRWCount = mRWCount + ReplaceWords("3 percent ", "3%")
    mRWCount = mRWCount + ReplaceWords("4 percent ", "4%")
    mRWCount = mRWCount + ReplaceWords("5 percent ", "5%")
    mRWCount = mRWCount + ReplaceWords("6 percent ", "6%")
    mRWCount = mRWCount + ReplaceWords("7 percent ", "7%")
    mRWCount = mRWCount + ReplaceWords("8 percent ", "8%")
    mRWCount = mRWCount + ReplaceWords("9 percent ", "9%")
    
    mRWCount = mRWCount + ReplaceWords("24-7", "24/7")
    
    mRWCount = mRWCount + ReplaceWords("alright", "all right")
    mRWCount = mRWCount + ReplaceWords("Alright", "All right")
    mRWCount = mRWCount + ReplaceWords("appdev", "AppDev")
    mRWCount = mRWCount + ReplaceWords("b2b", "B2B")
    mRWCount = mRWCount + ReplaceWords("b2b2c", "B2B2C")
    mRWCount = mRWCount + ReplaceWords("b2c", "B2C")
    mRWCount = mRWCount + ReplaceWords("Brownfield", "brownfield")
    mRWCount = mRWCount + ReplaceWords("c2b", "C2B")
    mRWCount = mRWCount + ReplaceWords("c2c", "C2C")
    mRWCount = mRWCount + ReplaceWords("CAAS", "CaaS")
    mRWCount = mRWCount + ReplaceWords("capex", "CapEx")
    mRWCount = mRWCount + ReplaceWords("codev", "co-dev")
    mRWCount = mRWCount + ReplaceWords("cross talk", "crosstalk")
    mRWCount = mRWCount + ReplaceWords("crossstalk", "crosstalk")
    mRWCount = mRWCount + ReplaceWords("devops", "DevOps")
    mRWCount = mRWCount + ReplaceWords("DRAAS", "DRaaS")
    mRWCount = mRWCount + ReplaceWords("ecommerce", "e-commerce")
    mRWCount = mRWCount + ReplaceWords("eCommerce", "e-commerce")
    mRWCount = mRWCount + ReplaceWords("eSignature", "e-signature")
    mRWCount = mRWCount + ReplaceWords("et cetera", "etc.")
    mRWCount = mRWCount + ReplaceWords("FinTech", "fintech")
    mRWCount = mRWCount + ReplaceWords("gonna", "going to")
    mRWCount = mRWCount + ReplaceWords("Greenfield", "greenfield")
    mRWCount = mRWCount + ReplaceWords("HIPPA", "HIPAA")
    mRWCount = mRWCount + ReplaceWords("InsurTech", "insurtech")
    mRWCount = mRWCount + ReplaceWords("IOT", "IoT")
    mRWCount = mRWCount + ReplaceWords("Ok", "Okay")
    mRWCount = mRWCount + ReplaceWords("opex", "OpEx")
    mRWCount = mRWCount + ReplaceWords("owner/operator", "owner-operator")
    mRWCount = mRWCount + ReplaceWords("PAAS", "PaaS")
    mRWCount = mRWCount + ReplaceWords("payor", "payer")
    mRWCount = mRWCount + ReplaceWords("SAAS", "SaaS")
    mRWCount = mRWCount + ReplaceWords("smart watch", "smartwatch")
    mRWCount = mRWCount + ReplaceWords("software as a service", "Software-as-a-Service")
    mRWCount = mRWCount + ReplaceWords("software-as-a-service", "Software-as-a-Service")
    mRWCount = mRWCount + ReplaceWords("three-D", "3D")
    mRWCount = mRWCount + ReplaceWords("two-D", "2D")
    mRWCount = mRWCount + ReplaceWords("wanna", "want to")
    mRWCount = mRWCount + ReplaceWords("zip code", "ZIP Code")
    mRWCount = mRWCount + ReplaceWords("Zip Code", "ZIP Code")
    mRWCount = mRWCount + ReplaceWords("Zip code", "ZIP Code")
    mRWCount = mRWCount + ReplaceWords("Zip", "ZIP")

    
            'MESSAGE BOX DISPLAYING LIST OF WORDS REPLACED IF ANY WORDS WERE REPLACED
            
            If mRWCount <> "" Then
                MsgBox ("The following words were replaced:" & vbCr & vbCr & mRWCount)
            End If

End Sub


Private Function ReplaceWords(TestWord, RepWord As String) As String

        Set mCurrentRng = ActiveDocument.Content
        
        mCurrentRng.Find.ClearFormatting

        ' This function will replace the first word with the second word
        
        With mCurrentRng.Find                        ' Replaces test word
            .Text = TestWord
            .Forward = True
            .MatchWildcards = False
            .MatchCase = True
            .MatchWholeWord = True
            .Wrap = 0
            
                Do While .Execute           ' counts the instances
                    ReplaceWords = TestWord + " with " + RepWord + vbCr
                    mCurrentRng.Find.Replacement.Text = RepWord
                   
                Loop
        
        End With
             ' replaces the word
        Set mCurrentRng = ActiveDocument.Content
        With mCurrentRng.Find
            .MatchWholeWord = True
            .MatchCase = True
            .Execute Findtext:=TestWord, replacewith:=RepWord, _
            Replace:=wdReplaceAll
        End With

End Function

Private Sub ReplaceQuotes()

    ' This is to account for graves and acute accents for those who use US International keyboards
        
        ActiveDocument.Content.Find.Execute Findtext:=Chr(96), MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:=Chr(39), Replace:=wdReplaceAll
        
        ActiveDocument.Content.Find.Execute Findtext:=Chr(180), MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:=Chr(39), Replace:=wdReplaceAll

    ' This is designed to replace all straight quotes with curly quotes without changing the users predefined settings.
    
    Dim blnQuotes As Boolean

        blnQuotes = Application.Options.AutoFormatAsYouTypeReplaceQuotes
    
    If Application.Options.AutoFormatAsYouTypeReplaceQuotes = True Then ' If the user already has curly quotes selected
    
        ActiveDocument.Content.Find.Execute Findtext:="'", MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:="'", Replace:=wdReplaceAll
        
    ElseIf Application.Options.AutoFormatAsYouTypeReplaceQuotes = False Then ' If the user does not have curly quotes selected
    
        Application.Options.AutoFormatAsYouTypeReplaceQuotes = True
        ActiveDocument.Content.Find.Execute Findtext:="'", MatchCase:=False, _
        MatchWholeWord:=False, MatchWildcards:=False, Wrap:=wdFindStop, _
        Format:=False, replacewith:="'", Replace:=wdReplaceAll
        Application.Options.AutoFormatAsYouTypeReplaceQuotes = False ' Changes it back to the user preferred setting
        
    End If

End Sub

Private Sub KeyStyleGuide_Click()

    ActiveDocument.FollowHyperlink ("https://alphasights.docsend.com/view/q2fy267i7t/d/yu5z3ki9ubfr9izd")

End Sub

Private Sub Leave_Click()

    Unload Me

End Sub

Private Sub TermDB_Click()

    ActiveDocument.FollowHyperlink ("https://docs.google.com/spreadsheets/d/1npD88Ud6XB4xIplKW8_OJjLkTd0O105e3WlfMPGzt5A/edit#gid=1097598184")
    
End Sub

Private Sub TR_Tag_Count_Click()
    
    Dim ICount, GCount, CCount, FCount, SCount, DCount, RCount As Integer
    Dim InDisp, GuessDisp, CrossDisp As String

    ICount = CountThis("[inaudible") 'Counts listed tags
    GCount = GCount + CountThis("? 0")
    GCount = GCount + CountThis("? 1")
    GCount = GCount + CountThis("? 2")
    GCount = GCount + CountThis("? 3")
    GCount = GCount + CountThis("? 4")
    GCount = GCount + CountThis("? 5")
    CCount = CountThis("[crosstalk")
    FCount = CountThis("[foreign")
    SCount = CountThis("[silence")
    DCount = CountThis("[call dropped")
    
    'Special loop used to search for research tags
    Set mCurrentRng = ActiveDocument.Content
    With mCurrentRng.Find
        .MatchWildcards = True
        .Text = "\[*\]"
        .Format = False
        .Wrap = 0
        .Forward = False
        
            Do While .Execute
                RCount = RCount + 1
            Loop
    End With
    
    RCount = RCount - FCount - CCount - GCount - ICount - SCount - DCount
    If RCount < 0 Then
        RCount = 0
    End If
    
    
    ' Display information to user
    
    
    If ICount <> 1 Then
        InDisp = "inaudibles"
    Else
        InDisp = "inaudible"
    End If
    
    If GCount <> 1 Then
        GuessDisp = "guesses"
    Else
        GuessDisp = "guess"
    End If
    
    If CCount <> 1 And CCount <> 0 Then
        CrossDisp = "crosstalks"
    Else
        CrossDisp = "crosstalk"
    End If

    ' Display final message
    
        Response = MsgBox("This document contains" & vbCr & "the following tags:" & vbCr & vbCr & _
                ICount & " " & InDisp & vbCr _
                & GCount & " " & GuessDisp & vbCr _
                & CCount & " " & CrossDisp & _
                vbCr & FCount & " foreign" & vbCr & _
                RCount & " research", vbOKOnly + vbInformation, "Tag Count")


End Sub


Private Sub Warning_Check()
    
    
    'For TGCount
    Dim TGCount As Integer
    
    'For Speaker IDs
    Dim InterviewerIDCount, FakeInterviewerIDCount, AdvisorIDCount, FakeAdvisorIDCount, SpeakerErrorCount As Integer
    
    'For Function Calls
    Dim TSCount As Integer
       
    FakeInterviewerIDCount = 0
    InterviewerIDCount = 0
    FakeAdvisorIDCount = 0
    AdvisorIDCount = 0
    SpeakerErrorCount = 0

    
    TSCount = 0
    mPCount = 0
    TGCount = 0
   
    mWild = False
    
    mPCount = mPCount + CheckComment(Chr(133), "Punctuation: Ellipses")
    mPCount = mPCount + CheckComment("...", "Punctuation: Ellipses")
    mPCount = mPCount + CheckComment("; ", "Semicolon")
    mPCount = mPCount + CheckComment("! ", "Exclamation Point")
    mPCount = mPCount + CheckComment("--", "Double Dash")
    mPCount = mPCount + CheckComment("~", "Tilde")
    mPCount = mPCount + CheckComment("<", "Wrong Bracket")
    mPCount = mPCount + CheckComment(">", "Wrong Bracket")
    mPCount = mPCount + CheckComment("{", "Wrong Bracket")
    mPCount = mPCount + CheckComment("}", "Wrong Bracket")
    
    mWild = True

    TSCount = TSCount + CheckComment("[0-9][0-9]:[0-9][0-9]:", "Style Guide: Timestamps should be in the following format: H:MM:SS.")
    TSCount = TSCount + CheckComment(":[0-9][0-9].[0-9]", "Style Guide: Timestamps should be in the following format: H:MM:SS.")
    TSCount = TSCount + CheckComment(":[0-9][0-9][0-9]", "Style Guide: Timestamps should be in the following format: H:MM:SS.")
    
    mWild = False

    TGCount = TGCount + CountThis("[inaudible")
    TGCount = TGCount + CountThis("? 0")
    TGCount = TGCount + CountThis("? 1")
    TGCount = TGCount + CountThis("? 2")
    TGCount = TGCount + CountThis("? 3")
    TGCount = TGCount + CountThis("? 4")
    TGCount = TGCount + CountThis("? 5")

        '------------------------------------------------------------------------------------------------------------------------------------
        
        '   DISPLAYS WARNING MESSAGES
        

        If mPCount = 1 Then
            Response = MsgBox("There is 1 punctuation error detected.", vbOKOnly + vbInformation, "Warning Check")
        ElseIf mPCount > 1 Then
            Response = MsgBox("There are " & mPCount & " punctuation errors detected.", vbOKOnly + vbInformation, "Warning Check")
        End If

        If TSCount = 1 Then
            Response = MsgBox("There is 1 timestamp error detected.", vbOKOnly + vbInformation, "Warning Check")
        ElseIf TSCount > 1 Then
            Response = MsgBox("There are " & TSCount & " timestamp errors detected.", vbOKOnly + vbInformation, "Warning Check")
        End If
        
        If TGCount > 5 Then
            Response = MsgBox("There are " & TGCount & " missing words in this split." & vbCr & "Consider clearing some missing words if possible.", vbOKOnly + vbInformation, "Warning Check")
        End If
    
End Sub

Private Function CheckComment(CheckFor, DispCom As String) As Integer
i = 0
                ' This function will search for the first item, select it, and populate a comment with a preselected display message
                    
                    Set mCurrentRng = ActiveDocument.Content
                    With mCurrentRng.Find
                        .MatchWildcards = mWild
                        .Text = CheckFor  'What to look for
                        .Format = False
                        .Wrap = 0
                        .Forward = False
                        
                            Do While .Execute
                                ActiveDocument.Comments.Add _
                                Range:=mCurrentRng, Text:=DispCom  'adds the comment if found
                                i = i + 1 'counts the number of instances
                            Loop
                    End With
                    
                    CheckComment = i


End Function

Private Sub NoBorders()
'
' NoBorders Macro
'
'
    Selection.WholeStory
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone

    
End Sub

Private Function CountThis(CntThis As String) As Integer
i = 0
                    ' Run a loop that will count whatever you want
                    
                    Set mCurrentRng = ActiveDocument.Content
                    With mCurrentRng.Find
                        .ClearFormatting
                        .MatchWildcards = False
                        .Text = CntThis
                        .Wrap = 0
                        .Forward = False

                            Do While .Execute
                                i = i + 1
                            Loop
                            
                    End With
                    
                    CountThis = i
                    
End Function

Private Sub UserForm_Initialize()

    #If Mac Then
    
        ResizeUserForm Me
    
    #End If
    
Call Marquee

Headline.Caption = mHeadlineStr

End Sub

Private Sub ResizeUserForm(frm As Object, Optional dResizeFactor As Double = 0#)
  
  Dim ctrl As Control
  Dim sColWidths As String
  Dim vColWidths As Variant
  Dim iCol As Long

  If dResizeFactor = 0 Then dResizeFactor = 1.333333
  With frm
    .Height = .Height * dResizeFactor
    .Width = .Width * dResizeFactor

    For Each ctrl In frm.Controls
      With ctrl
        .Height = .Height * dResizeFactor
        .Width = .Width * dResizeFactor
        .Left = .Left * dResizeFactor
        .Top = .Top * dResizeFactor
        On Error Resume Next
        .Font.Size = .Font.Size * dResizeFactor
        On Error GoTo 0

        ' multi column listboxes, comboboxes
        Select Case TypeName(ctrl)
          Case "ListBox", "ComboBox"
            If ctrl.ColumnCount > 1 Then
              sColWidths = ctrl.ColumnWidths
              vColWidths = Split(sColWidths, ";")
              For iCol = LBound(vColWidths) To UBound(vColWidths)
                vColWidths(iCol) = Val(vColWidths(iCol)) * dResizeFactor
              Next
              sColWidths = Join(vColWidths, ";")
              ctrl.ColumnWidths = sColWidths
            End If
        End Select
      End With
    Next
  End With
End Sub

Private Sub OneSpace()

        Set mCurrentRng = ActiveDocument.Content
        mCurrentRng.Find.ClearFormatting

        ' This sub will replace two spaces with one
        
        With mCurrentRng.Find                        ' Replaces test word
            .Text = "  "
            .Forward = True
            .MatchWildcards = False
            .MatchCase = True
            .MatchWholeWord = True
            .Wrap = 0
            
                Do While .Execute
                    mCurrentRng.Find.Replacement.Text = " "
                    .Execute Replace:=wdReplaceAll
                Loop
        End With
End Sub

Private Sub Marquee()

Dim RndNum As Integer
RndNum = Int((5 * Rnd) + 1)    ' Generate random value between 1 and 3.

If RndNum = 1 Then
    mHeadlineStr = "Focus: High Accuracy and Consistency in Spelling!"
ElseIf RndNum = 2 Then
    mHeadlineStr = "Close Listening is Key!"
ElseIf RndNum = 3 Then
    mHeadlineStr = "Always use Context and Ensure Sentences Make Sense!"
ElseIf RndNum = 4 Then
    mHeadlineStr = "Use Standard Transcript template unless AlphaView is indicated."
Else
    mHeadlineStr = "Don't Forget to Check the Interaction Details!"
    
End If


End Sub
