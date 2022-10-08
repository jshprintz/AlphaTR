Attribute VB_Name = "Alpha_TR"
Option Explicit
Private mVarCntOne As Integer
Private aVar As Variable

Sub Alpha_TR()

    ActiveDocument.TrackRevisions = False
    TR_Checklist.show

End Sub

