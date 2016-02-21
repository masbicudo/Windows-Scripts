Function GetSomeStuff
  Dim stuff
  Set stuff = CreateObject("Scripting.Dictionary")
    stuff.Add "A", "Anaconda"
    stuff.Add "B", "Boa"
    stuff.Add "C", "Cobra"

  Set GetSomeStuff = stuff
End Function

Set d = GetSomeStuff
Wscript.Echo d.Item("A")
Wscript.Echo d.Exists("B")
items = d.Items
For i = 0 To UBound(items)
  Wscript.Echo items(i)
Next
