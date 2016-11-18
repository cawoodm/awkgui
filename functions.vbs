Function FileExists(filename)
   If fso.FileExists(filename) then FileExists=true Else FileExists=false
End Function

Function SaveToFile(filename, data)
' Save data to filename
   Dim f
   Set f = fso.CreateTextFile(filename)
   f.write data
   f.close
   Set f = nothing
End Function

Function AppendToFile(filename, data)
' Append data to filename, will create file If it doesn't exist
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   Dim f
   set f = fso.OpenTextFile(filename, 8, true)
   f.write data
   f.close
   Set f = nothing
End Function

Function LoadFromFile(filename, ByRef data)
' Load data from filename
   Dim f
   Set f = fso.OpenTextFile(filename)
   data = f.ReadAll
   f.close
   Set f = nothing
End Function

Function LoadFromFileTop(filename, ByRef data, top)
' Load top n lines of data from filename
   Dim f
   Set f = fso.OpenTextFile(filename)
   Dim i
   i = 0
   Do While Not f.AtEndOfStream And i < top
      i = i + 1
      data = data & f.ReadLine & vbNewLine
   Loop
   f.close
   Set f = nothing
End Function

Function DeleteFile(filename)
' Load data from filename
   Dim f
   On Error Resume Next
   Set f = fso.GetFile(filename)
   If Err=0 then
      f.Delete()
      Set f = Nothing
      DeleteFile = true
   Else
      DeleteFile = false
   End If
   On Error Goto 0
   Set f = nothing
End Function