Attribute VB_Name = "Module1"

Sub filebox_2_list(LB As ListBox, WhatToAdd As FileListBox)
Dim G
For G = 0 To WhatToAdd.ListCount - 1
WhatToAdd.Selected(G) = True
LB.AddItem WhatToAdd.FileName
Next G
End Sub
Sub filebox_2_list2(LB As ListBox, dir As DirListBox, WhatToAdd As FileListBox)
Dim G
For G = 0 To WhatToAdd.ListCount - 1
WhatToAdd.Selected(G) = True
If dir = "C:\" Then LB.AddItem "C:\" & WhatToAdd.FileName
If dir <> "C:\" Then LB.AddItem dir.Path & "\" & WhatToAdd.FileName
Next G
End Sub

Sub pause(Duration)
Dim starttime
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub

Sub crt(pth As String)
Dim f
f = FreeFile
Open App.Path + "\temp.html" For Output As #f
    Print #f, "<html><body><left><top><BODY BGCOLOR=000000><img src=""" & pth & """ BORDER=0 height=140 width=170></center></body></html>"
Close #f
DoEvents
Form2.WebBrowser1.Navigate App.Path & "\temp.html"
End Sub

