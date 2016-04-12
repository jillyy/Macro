Sub macro()
'
' macro macro
'
'
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 9
    Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Selection.Font.Name = "Times New Roman"
    Selection.TypeText Text:="Justify"
    ActiveDocument.SaveAs fileName:="/Users/Jill/Desktop/macro.docx", _
        FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False, HTMLDisplayOnlyOutput:=False, _
        MaintainCompat:=False
End Sub
