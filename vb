


Function tt(regString As String)

Dim rgx As Object
Set rgx = CreateObject("VBScript.RegExp")
Dim allMatches As Object
Dim card_s As String
Dim loanac As String

card_s = "([0-9]{4}[\s]{1}[0-9]{4}[\s]{1}[0-9]{4}[\s]{1}[0-9]{4})|(?:^|\D)(\d{16})(?!\d)"
loanac = "(?:^|\D)(\d{10})(?!\d)|(?:^|\D)(\d{12})(?!\d)"

With rgx
.Pattern = card_s & "|" & loanac   ' & "|" &card_mini & "|" &  card & "|"&
.Global = True
.IgnoreCase = True
.MultiLine = True
End With

Set allMatches = rgx.Execute(regString)
'Loop to read all the matches found
For Each Item In allMatches
MsgBox Item.Value
Next
End Function
