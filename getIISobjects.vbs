'Get the IIS Server Object
Set oW3SVC = GetObject("IIS://LOCALHOST/W3SVC/1/ROOT")

For Each oVirtualDirectory In oW3SVC

    Set oFile = CreateObject("Scripting.FileSystemObject")
    Set oTextFile = oFile.OpenTextFile("C:\Users\Rumi\Desktop\New folder\Results.txt", 8, True)

    oTextFile.WriteLine(oVirtualDirectory.class + " -" + oVirtualDirectory.Name)
    oTextFile.Close

Next
Set oTextFile = Nothing
Set oFile = Nothing

Wscript.Echo "Done"
