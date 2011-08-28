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
	Private FSO, FileName, FileContents
	
	' Extract a key's value from its line in the INI file
	Private Function ExtractValue(ByVal MyFrom, ByVal MyStart, ByVal MyEnd)
		Dim PosS: PosS = InStr(1, MyFrom, MyStart, 1)
		
		If PosS > 0 Then
			PosS = PosS + Len(MyStart)
			Dim PosE: PosE = InStr(PosS, MyFrom, MyEnd, 1)
			If PosE = 0 Then PosE = InStr(PosS, MyFrom, vbCrLf, 1)
			If PosE = 0 Then PosE = Len(MyFrom) + 1
			ExtractValue = Mid(MyFrom, PosS, PosE - PosS)
		Else
			ExtractValue = vbNullString
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
	
	Private Function GetSectionContents(MySection)
		Dim SectionContents: SectionContents = vbNullString
		Dim PosSection: PosSection = 0
		Dim PosEndSection: PosEndSection = 0
		
		' Find [Section] specified.
		PosSection = InStr(1, FileContents, "[" & MySection & "]", vbTextCompare)
		
		If PosSection > 0 Then ' Section exists.
			PosEndSection = InStr(PosSection, FileContents, vbCrLf & "[") ' Find end of section.
			
			' Is this last section? If so, mark the end of it as the end of the String (file's contents).
			If PosEndSection = 0 Then PosEndSection = Len(FileContents) + 1
			
			' Separate section contents.
    		SectionContents = Mid(FileContents, PosSection, PosEndSection - PosSection)
		End If
		
		GetSectionContents = SectionContents
	End Function
	
	' Write contents to file, overwriting previous file contents. If the file doesn't already exist, it is created.
	Private Sub WriteFileContents(ByVal MyContents)
		Dim FileStream: Set FileStream = FSO.OpenTextFile(FileName, ForWriting, True)
		FileStream.Write MyContents
		FileStream.Close()
	End Sub
	
	Public Default Function Init(MyFileName)
		Set FSO  = CreateObject("Scripting.FileSystemObject")
		FileName = MyFileName
		
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
	
	Public Function GetSections()
		Dim SectionsRegExp: Set SectionsRegExp = New RegExp
		
		' Matches a [Section] on its own line. Could be at the very beginning of the file,
		' in the middle of the file, or at the very end of the file (an empty [Section]).
		SectionsRegExp.Pattern = "([\r\n]\[|^\[)([^\]]*)(\][\r\n]|\]$)"
		
		' Matches all occurrences, not just the first one.
		SectionsRegExp.Global = True
		
		Dim SectionMatches: Set SectionMatches = SectionsRegExp.Execute(FileContents)
		
		Dim Sections: Sections = Array()
		Dim Index
		
		If SectionMatches.Count > 0 Then
			For Index = 0 To SectionMatches.Count - 1
				ReDim Preserve Sections(Index)
				Sections(Index) = SectionMatches.Item(Index).SubMatches(1)
			Next
		End If
		
		GetSections = Sections
	End Function
	
	Public Function GetKeys(MySection)
		' Grab the contents of the specified [Section]
		Dim SectionContents: SectionContents = GetSectionContents(MySection)
		
		Dim KeysRegExp: Set KeysRegExp = New RegExp
		
		' Matches a key= on its own line; captures the name of the key.
		KeysRegExp.Pattern = "[\r\n]{1,2}([^=]*)="
		
		' Matches all occurrences, not just the first one.
		KeysRegExp.Global = True
		
		Dim KeyMatches: Set KeyMatches = KeysRegExp.Execute(SectionContents)
		
		Dim Keys: Keys = Array()
		Dim Index
		
		If KeyMatches.Count > 0 Then
			For Index = 0 To KeyMatches.Count - 1
				ReDim Preserve Keys(Index)
				Keys(Index) = KeyMatches.Item(Index).SubMatches(0)
			Next
		End If
		
		GetKeys = Keys
	End Function
	
	Public Function GetValue(MySection, MyKeyName)
		Dim Value
		
		' Grab the contents of the specified [Section]
		Dim SectionContents: SectionContents = GetSectionContents(MySection)
		
		' Look for the key=
		If InStr(1, SectionContents, vbCrLf & MyKeyName & "=", vbTextCompare) > 0 Then
			' Extract the value from the key= line.
			Value = ExtractValue(SectionContents, vbCrLf & MyKeyName & "=", vbCrLf)
		End If
		
		GetValue = Value
	End Function
	
	' Unlike Dictionary's Add() method, SetValue() can change the value of an existing key.
	Public Sub SetValue(MySection, MyKeyName, MyValue)
		' Grab the contents of the specified [Section]
		Dim OldSectionContents: OldSectionContents = GetSectionContents(MySection)
		
		If OldSectionContents <> vbNullString Then ' Section exists.
			Dim KeyName, Line, Found, NewSectionContents
			
			OldSectionContents = Split(OldSectionContents, vbCrLf)
			KeyName = LCase(MyKeyName & "=") ' Temp variable to find a key.
			
			' Copy each line over; if the key matches, change its value first.
			For Each Line In OldSectionContents
				If LCase(Left(Line, Len(KeyName))) = KeyName Then
					Line = KeyName & MyValue
					Found = True
				End If
				
				NewSectionContents = NewSectionContents & Line & vbCrLf
			Next
			
			If IsEmpty(Found) Then ' Key not found.
				' Append it to the [Section].
				NewSectionContents = NewSectionContents & KeyName & MyValue
			Else ' Key found.
				' Remove last vbCrLf. There's already a vbCrLf at the end of the [Section].
				NewSectionContents = Left(NewSectionContents, Len(NewSectionContents) - 2)
			End If
			
			' Combine pre-section, new section, and post-section data.
			FileContents = Left(FileContents, PosSection-1) & NewSectionContents & Mid(FileContents, PosEndSection)
		Else ' Section doesn't exist.
			' If the file doesn't already end in a new line, and if the file isn't empty...
			If Right(FileContents, 2) <> vbCrLf And Len(FileContents) > 0 Then
				' Add a new line to the end of the file
				FileContents = FileContents & vbCrLf
			End If
			
			' Add section data at the end of file contents.
			FileContents = FileContents & "[" & MySection & "]" & vbCrLf & MyKeyName & "=" & MyValue
		End If ' If OldSectionContents <> vbNullString Then
	End Sub
	
	Public Sub Load
		FileContents = GetFileContents
	End Sub
	
	Public Sub Save
		WriteFileContents(FileContents)
	End Sub
End Class