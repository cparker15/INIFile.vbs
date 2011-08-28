INIFile.vbs is a VBScript class for reading from and writing to INI files. It can be used anywhere VBScript can be used, such as in an HTML Application (HTA) or in a Windows Script Host (WSH) script.

Usage Examples
==============

SetValue(), GetValue(), Load(), Save()
--------------------------------------

```vb.net
Dim INI: Set INI = (New INIFile)("MyIniFile.ini") ' automatically loads the contents of the file
INI.SetValue "test section 1", "test key 1", "test value 1"
INI.SetValue "test section 2", "test key 2", "test value 2"
INI.SetValue "test section 3", "test key 3", "test value 3"
INI.SetValue "test section 4", "test key 4", "test value 4"
INI.Save ' file remains untouched until here
INI.Load ' reloads contents of the file
INI.SetValue "test section 5", "test key 5", "test value 5"
INI.Save

MsgBox "test section 5, test key 5: " & INI.GetValue("test section 5", "test key 5")
```

This example results in an INI file named "MyIniFile.ini", located in the current working directory, that looks like the following output.

```
[test section 1]
test key 1=test value 1
[test section 2]
test key 2=test value 2
[test section 3]
test key 3=test value 3
[test section 4]
test key 4=test value 4
[test section 5]
test key 5=test value 5
```

It then displays a message box with the text `test section 5, test key 5: test value 5`.

GetSections(), GetKeys()
------------------------

This example demonstrates how one could use the GetSections() and GetKeys() methods to create an outline of an INI file.

```vb.net
Dim INI: Set INI = (New INIFile)("MyIniFile.ini") ' automatically loads the contents of the file
Dim Msg: Msg = "INI File Outline:" & vbCrLf & vbCrLf

Dim Sections: Sections = INI.GetSections()
Dim Section

If UBound(Sections) > -1 Then
	For Each Section In Sections
		Msg = Msg & "- " & Section & vbCrLf
		
		Dim Keys: Keys = INI.GetKeys(Section)
		Dim Key
		
		For Each Key In Keys
			Msg = Msg & "  - " & Key & vbCrLf
		Next
	Next
Else
	Msg = "INI File Is Empty"
End If

MsgBox Msg
```

Assuming the **SetValue(), GetValue(), Load(), Save()** example code was just run, this code produces a message box with the following text:

```
INI File Outline:

- test section 1
  - test key 1
- test section 2
  - test key 2
- test section 3
  - test key 3
- test section 4
  - test key 4
- test section 5
  - test key 5
```

Notes
=====

- INIFile.vbs does not support comments in INI files. It has not been tested with INI files containing comments.

- INIFile.vbs is non-destructive. Any sections and key/value pairs that are added via INIFile.vbs are appended to the INI file.

- *IMPORTANT!* The encoding of INIFile.vbs *MUST NOT* be changed to UTF-8. Doing so will result in the following error under Windows Script Host, which cannot handle scripts encoded in UTF-8.

```
Script: INIFile.vbs
Line: 1
Char: 1
Error: Invalid character
Code: 800A0408
Source: Microsoft VBScript compilation error
```