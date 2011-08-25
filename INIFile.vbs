' INIFile.vbs: A VBScript class for reading from and writing to INI files.
' Copyright (C) 2004, 2011 Christopher Parker <http://www.cparker15.com/>
' 
' INIFile.vbs is free software: you can redistribute it and/or modify
' it under the terms of the GNU Lesser General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' INIFile.vbs is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU Lesser General Public License for more details.
' 
' You should have received a copy of the GNU Lesser General Public License
' along with INIFile.vbs.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

Const ForReading = 1, ForWriting = 2, ForAppending = 8

Class INIFile
	Private FSO, FileName, FileContents, PosSection, PosEndSection
	
	' Separate one field between Start and End
	Private Function SeparateField(ByVal MyFrom, ByVal MyStart, ByVal MyEnd)
		Dim PosS: PosS = InStr(1, MyFrom, MyStart, 1)
		
		If PosS > 0 Then
			PosS = PosS + Len(MyStart)
			Dim PosE: PosE = InStr(PosS, MyFrom, MyEnd, 1)
			If PosE = 0 Then PosE = InStr(PosS, MyFrom, vbCrLf, 1)
			If PosE = 0 Then PosE = Len(MyFrom) + 1
			SeparateField = Mid(MyFrom, PosS, PosE - PosS)
		Else
			SeparateField = vbNullString
		End If
	End Function
	
	Private Function GetFileContents()
		If Not FileExists() Then
			GetFileContents = vbNullString
		ElseIf FileIsEmpty() Then
			GetFileContents = vbNullString
		Else
			GetFileContents = FSO.OpenTextFile(FileName, ForReading).ReadAll
		End If
	End Function
	
	' Write contents to file, overwriting previous file contents. If the file doesn't already exist, it is created.
	Private Sub WriteFileContents(ByVal MyContents)
		Dim FileStream: Set FileStream = FSO.OpenTextFile(FileName, ForWriting, True)
		FileStream.Write MyContents
		FileStream.Close()
	End Sub
	
	Public Default Function Init(MyFileName)
		Set FSO       = CreateObject("Scripting.FileSystemObject")
		FileName      = MyFileName
		PosSection    = 0
		PosEndSection = 0
		
		Load
		
		Set Init = Me
	End Function
	
	Public Function GetFileName()
		GetFileName = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))
	End Function
	
	Public Function GetFilePath()
		GetFilePath = Left(FileName, InStrRev(FileName, "\"))
	End Function
	
	Public Function FileExists()
		FileExists = FSO.FileExists(FileName)
	End Function
	
	Public Function FileIsEmpty()
		FileIsEmpty = FSO.OpenTextFile(FileName).AtEndOfStream
	End Function
	
	Public Function GetValue(MySection, MyKeyName)
		Dim Value: Value = Empty
		
		' Find [Section] specified.
		PosSection = InStr(1, FileContents, "[" & MySection & "]", vbTextCompare)
		
		If PosSection > 0 Then ' Section exists. 
			PosEndSection = InStr(PosSection, FileContents, vbCrLf & "[") ' Find end of section.
			
			' Is this last section? If so, mark the end of it as the end of the String (file's contents).
			If PosEndSection = 0 Then PosEndSection = Len(FileContents) + 1
			
			Dim SectionContents ' Separate section contents.
    		SectionContents = Mid(FileContents, PosSection, PosEndSection - PosSection)
			
			If InStr(1, SectionContents, vbCrLf & MyKeyName & "=", vbTextCompare) > 0 Then
				Value = SeparateField(SectionContents, vbCrLf & MyKeyName & "=", vbCrLf) ' Separate value of a key.
			End If
		End If
		
		GetValue = Value ' Return the corresponding value for the key specified.
	End Function
	
	Public Sub SetValue(MySection, MyKeyName, MyValue)
		' Find [Section] specified.
		PosSection = InStr(1, FileContents, "[" & MySection & "]", vbTextCompare)
		
		If PosSection > 0 Then ' Section exists.
			PosEndSection = InStr(PosSection, FileContents, vbCrLf & "[") ' Find end of section.
			
			' Is this last section? If so, mark the end of it as the end of the String (file's contents).
			If PosEndSection = 0 Then PosEndSection = Len(FileContents) + 1
			
			Dim OldSectionContents, NewSectionContents, Line ' Separate section contents
			OldSectionContents = Mid(FileContents, PosSection, PosEndSection - PosSection)
			OldSectionContents = Split(OldSectionContents, vbCrLf)
			
			Dim KeyName, Found
			KeyName = LCase(MyKeyName & "=") ' Temp variable to find a key.
			
			For Each Line In OldSectionContents
				If LCase(Left(Line, Len(KeyName))) = KeyName Then
					Line = MyKeyName & "=" & MyValue
					Found = True
				End If
				
				NewSectionContents = NewSectionContents & Line & vbCrLf
			Next
			
			If IsEmpty(Found) Then ' Key not found.
				NewSectionContents = NewSectionContents & MyKeyName & "=" & MyValue ' Add it at the end of section.
			Else ' Key found.
				NewSectionContents = Left(NewSectionContents, Len(NewSectionContents) - 2) ' Remove last vbCrLf. vbCrLf is at PosEndSection.
			End If
			
			' Combine pre-section, new section, and post-section data.
			FileContents = Left(FileContents, PosSection-1) & NewSectionContents & Mid(FileContents, PosEndSection)
		Else ' Section doesn't exist.
			If Right(FileContents, 2) <> vbCrLf And Len(FileContents) > 0 Then
				FileContents = FileContents & vbCrLf ' Add section data at the end of file contents.
			End If
			
			FileContents = FileContents & "[" & MySection & "]" & vbCrLf & MyKeyName & "=" & MyValue
		End If ' If PosSection > 0 Then
	End Sub
	
	Public Sub Load
		FileContents = GetFileContents
	End Sub
	
	Public Sub Save
		WriteFileContents(FileContents)
	End Sub
End Class