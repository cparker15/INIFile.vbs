INIFile.vbs is a VBScript class for reading from and writing to INI files. It can be used anywhere VBScript can be used, such as in an HTML Application (HTA) or in a Windows Script Host (WSH) script.

Usage example:

    Dim INI: Set INI = New INIFile
    INI.SetFileName "C:\\MyIniFile.ini"
    INI.SetValue "test section 1", "test key 1", "test value 1"
    INI.SetValue "test section 1", "test key 2", "test value 2"
    INI.SetValue "test section 2", "test key 3", "test value 3"
    INI.SetValue "test section 2", "test key 4", "test value 4"

The above code snippet results in an INI file at C:\MyIniFile.ini that looks like this:

    [test section 1]
    test key 1=test value 1
    test key 2=test value 2
    [test section 2]
    test key 3=test value 3
    test key 4=test value 4

Note that this class is non-destructive. Any sections and key/value pairs that are added via this class are added to the INI file. If you need to wipe the INI file clean for some reason before writing to it, then the above usage example would look like this:

    Dim INI: Set INI = New INIFile
    INI.SetFileName "C:\\MyIniFile.ini"
    INI.WriteFileContents ""
    INI.SetValue "test section 1", "test key 1", "test value 1"
    INI.SetValue "test section 1", "test key 2", "test value 2"
    INI.SetValue "test section 2", "test key 3", "test value 3"
    INI.SetValue "test section 2", "test key 4", "test value 4"