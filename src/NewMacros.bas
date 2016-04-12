Attribute VB_Name = "NewMacros"
Sub GrammarMacro()
' GrammarMacro Macro
With Selection.Find
' This part of the macro searches for common grammar mistakes and replaces them with the proper form
Set myRange = ActiveDocument.Content
myRange.Find.Execute FindText:=".)", ReplaceWith:=").", _
Replace:=wdReplaceAll
myRange.Find.Execute FindText:="!)", ReplaceWith:=")!", _
Replace:=wdReplaceAll
myRange.Find.Execute FindText:="?)", ReplaceWith:=")?", _
Replace:=wdReplaceAll
myRange.Find.Execute FindText:="non ", ReplaceWith:="non", _
Replace:=wdReplaceAll
myRange.Find.Execute FindText:="non-", ReplaceWith:="non", _
Replace:=wdReplaceAll
myRange.Find.Execute FindText:="datas", ReplaceWith:="data", _
Replace:=wdReplaceAll
myRange.Find.Execute FindText:="microns", ReplaceWith:="micrometer", _
Replace:=wdReplaceAll
myRange.Find.Execute FindText:="et. al", ReplaceWith:="et al", _
Replace:=wdReplaceAll
MsgBox "Common grammar mistakes check complete."

End With

End Sub

