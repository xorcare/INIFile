VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INIFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'################################################################################
'#                                                                              #
'#    Copyright 2017 Vasilyuk Vasiliy vasilyukvasiliy@gmail.com                 #
'#                                                                              #
'#    Licensed under the Apache License, Version 2.0 (the "License");           #
'#    you may not use this file except in compliance with the License.          #
'#    You may obtain a copy of the License at                                   #
'#                                                                              #
'#        http://www.apache.org/licenses/LICENSE-2.0                            #
'#                                                                              #
'#    Unless required by applicable law or agreed to in writing, software       #
'#    distributed under the License is distributed on an "AS IS" BASIS,         #
'#    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  #
'#    See the License for the specific language governing permissions and       #
'#    limitations under the License.                                            #
'#                                                                              #
'################################################################################

Const TYPE_OF_DICTIONARY As Integer = 2
Const TYPE_OF_STRING As Integer = 1
Const TYPE_OF_NOT_EXIST As Integer = 0
Const TYPE_OF_NOT_DEFINED As Integer = -1
Const DEFAULT_SECTION As String = "DEFAULT-INI-FILE-SECTION"

'Format:
'*  0  = TristateFalse - Default. Open the file as ASCII
'*  -1 = TristateTrue - Open the file as Unicode
'*  -2 = TristateUseDefault - Open the file using the system default
Public Format As Integer

'Create the file if it does not exist
'* True - Create
'* False - Ignore
Public Create As Boolean

Private private_data As Object
Private private_path As String
Private private_loaded As Boolean
Private private_compareMode As Integer
Private Const PRIVATE_IOMODE As Integer = 1
Private FSO As Object

Public Function GetCompareMode() As Byte
    GetCompareMode = private_compareMode
End Function

Public Function SetCompareMode(ByVal mode As Byte) As Boolean
    SetCompareMode = False
    If private_loaded Then
        Exit Function
    End If
    If mode = vbTextCompare Or mode = vbBinaryCompare Then
        private_compareMode = CInt(mode)
        SetCompareMode = True
    End If
End Function

'Set path to ini file
Public Sub SetPath(ByVal inputPath As String)
    If Not private_loaded Then
        private_path = FSO.GetAbsolutePathName(inputPath)
    End If
End Sub

'Get string path for ini file
Public Function GetPath() As String
    GetPath = private_path
End Function

'Loaded data from a file
Public Function Loaded() As Boolean
    Loaded = private_loaded
End Function

Private Sub Class_Initialize()
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set private_data = Nothing
    private_compareMode = vbTextCompare
    Format = 0
    Create = False
    private_path = ""
    Set private_data = CreateObject("Scripting.Dictionary")
    private_data.compareMode = private_compareMode
End Sub

Public Sub Clean()
    Set private_data = Nothing
    private_compareMode = vbTextCompare
    Format = 0
    Create = False
    private_path = ""
    Set private_data = CreateObject("Scripting.Dictionary")
    Let private_data.compareMode = private_compareMode
End Sub

Private Sub Class_Terminate()
'Class_Terminate
End Sub

Public Function LoadFrom(ByVal inputPath As String) As Boolean
    private_path = FSO.GetAbsolutePathName(inputPath)
    LoadFrom = Load()
End Function

'Checks whether there is property inside the section (small call)
Public Function E(ByVal sectionName As String, Optional propertyName As String = "") As Boolean
    E = ExistProperty(sectionName, propertyName)
End Function

'Checks whether there is property inside the section
Public Function ExistProperty(ByVal sectionName As String, Optional propertyName As String = "") As Boolean
    ExistProperty = False
    If Not CheckLocation(sectionName, propertyName) Then
        Exit Function
    End If
    
    If private_data.Exists(sectionName) Then
        Set Section = private_data.item(sectionName)
        If Section.Exists(propertyName) Then
            ExistProperty = True
        End If
    End If
End Function

'Returns the value of the property in a certain section (small call)
Public Function G(ByVal sectionName As String, Optional propertyName As String = "") As String
    G = GetProperty(sectionName, propertyName)
End Function

'Returns the value of the property in a certain section
Public Function GetProperty(ByVal sectionName As String, Optional propertyName As String = "") As String
    GetProperty = ""
    
    If Not CheckLocation(sectionName, propertyName) Then
        Exit Function
    End If
    
    If TypeOfProperty(sectionName, propertyName) <> TYPE_OF_STRING Then
        Exit Function
    End If
    
    If private_data.Exists(sectionName) Then
        Set Section = private_data.item(sectionName)
        If Section.Exists(propertyName) Then
            GetProperty = CStr(Section.item(propertyName))
        End If
    End If
End Function

'Checks for the existence of the section (small call)
Public Function ES(ByVal sectionName As String) As Boolean
    ES = ExistSection(sectionName)
End Function

'Checks for the existence of the section
Public Function ExistSection(ByVal sectionName As String) As Boolean
    ExistSection = False
    KeyCase sectionName
    If private_data.Exists(sectionName) Then
        ExistSection = True
    End If
End Function

'Returns entire section as a dictionary (small call)
Public Function GS(ByVal sectionName As String) As Object
    GS = Nothing
    Set GS = GetSection(sectionName)
End Function

'Returns entire section as a dictionary
Public Function GetSection(ByVal sectionName As String) As Object
    Set GetSection = CreateObject("Scripting.Dictionary")
    KeyCase sectionName
    If private_data.Exists(sectionName) Then
        Set GetSection = private_data.item(sectionName)
    End If
End Function

'Get the type of properties (small call)
Public Function TOP(ByVal sectionName As String, Optional propertyName As String = "") As Integer
    TOP = TypeOfProperty(sectionName, propertyName)
End Function

'Get the type of properties
Public Function TypeOfProperty(ByVal sectionName As String, Optional propertyName As String = "") As Integer
    On Error GoTo ErrHandlerTypeOfProperty
    TypeOfProperty = TYPE_OF_NOT_DEFINED
    If Not ExistProperty(sectionName, propertyName) Then
        TypeOfProperty = TYPE_OF_NOT_EXIST
        Exit Function
    End If
    Set Section = private_data.item(sectionName)
    If LCase(TypeName(Section.item(propertyName))) = "dictionary" Then
        TypeOfProperty = TYPE_OF_DICTIONARY
        Exit Function
    End If
    If LCase(TypeName(Section.item(propertyName))) = "string" Then
        TypeOfProperty = TYPE_OF_STRING
        Exit Function
    End If
    Exit Function
ErrHandlerTypeOfProperty:
    TypeOfProperty = TYPE_OF_NOT_DEFINED
End Function

'Reading and parsing a ini file
Public Function Load() As Boolean
    Load = False
    On Error GoTo LoadErrHandler
    If private_loaded Then
        Exit Function
    End If
    Set textStream = FSO.OpenTextFile(private_path, PRIVATE_IOMODE, Create, Format)
    Set Regexp = CreateObject("VBScript.RegExp")
    private_data.compareMode = private_compareMode
    Dim keyString As String
    Dim PartTitle As String
    Dim value As String
    PartTitle = DEFAULT_SECTION
    Do While Not textStream.AtEndOfStream
        fileStr = textStream.ReadLine
        fileStr = Trim(fileStr)
        If Len(fileStr) > 0 And InStr(0, fileSrt, ";", 1) = 0 And InStr(0, fileSrt, ";", 1) = 0 Then
            Matched = False
            Regexp.Pattern = "^\[([^\[\]]{1,})\][ \t]*([;#]{1,}.*|)$"
            If Regexp.Test(fileStr) And Matched = False Then
                Set Matches = Regexp.Execute(fileStr)
                For Each Match In Matches
                    PartTitle = CStr(Match.SubMatches(0))
                    KeyCase PartTitle
                Next
                Matched = True
            End If

            Regexp.Pattern = "^([^\[\]]{1,})\[\]\=([^"";#]*)[ \\t]*([;#]{1,}.*|)$"
            If Regexp.Test(fileStr) And Matched = False Then
                Set Matches = Regexp.Execute(fileStr)
                For Each Match In Matches
                    keyString = CStr(Match.SubMatches(0))
                    value = CStr(Match.SubMatches(1))
                    KeyCase keyString
                    SetDictionaryProperty PartTitle, keyString, value
                Next
                Matched = True
            End If
            
            Regexp.Pattern = "^([^\[\]]{1,})\[\]\=""(.*)""[ \\t]*([;#]{1,}.*|)$"
            If Regexp.Test(fileStr) And Matched = False Then
                Set Matches = Regexp.Execute(fileStr)
                For Each Match In Matches
                    keyString = CStr(Match.SubMatches(0))
                    value = CStr(Match.SubMatches(1))
                    KeyCase keyString
                    SetDictionaryProperty PartTitle, keyString, value
                Next
                Matched = True
            End If
            
            Regexp.Pattern = "^([^\[\]]*)\=""(.*)""[ \\t]*([;#]{1,}.*|)$"
            If Regexp.Test(fileStr) And Matched = False Then
                Set Matches = Regexp.Execute(fileStr)
                For Each Match In Matches
                    keyString = CStr(Match.SubMatches(0))
                    value = CStr(Match.SubMatches(1))
                    KeyCase keyString
                    SetStringProperty PartTitle, keyString, value
                Next
                Matched = True
            End If
            
            Regexp.Pattern = "^([^\[\]]*)\=([^"";#]*)[ \\t]*([;#]{1,}.*|)$"
            If Regexp.Test(fileStr) And Matched = False Then
                Set Matches = Regexp.Execute(fileStr)
                For Each Match In Matches
                    keyString = CStr(Match.SubMatches(0))
                    value = CStr(Match.SubMatches(1))
                    KeyCase keyString
                    SetStringProperty PartTitle, keyString, value
                Next
                Matched = True
            End If
        End If
        keyString = ""
        value = ""
    Loop
    textStream.Close
    Set LoadIniFile = private_data
    Load = True
    private_loaded = True
    Exit Function
LoadErrHandler:
    Load = False
End Function

Private Sub SetStringProperty(ByVal PartTitle As String, ByVal keyString As String, ByVal value As String)
    If private_data.Exists(PartTitle) Then
        Set properties = private_data.item(PartTitle)
    Else
        Set properties = CreateObject("Scripting.Dictionary")
    End If
    If Not properties.Exists(keyString) Then
        properties.Add keyString, value
    Else
        properties.item(keyString) = value
    End If
    Set private_data.item(PartTitle) = properties
End Sub

Private Sub SetDictionaryProperty(ByVal PartTitle As String, ByVal keyString As String, ByVal value As String)
    If private_data.Exists(PartTitle) Then
        Set properties = private_data.item(PartTitle)
    Else
        Set properties = CreateObject("Scripting.Dictionary")
    End If
    If Not properties.Exists(keyString) Then
        Set dictionaryArray = CreateObject("Scripting.Dictionary")
        dictionaryArray.Add CStr(dictionaryArray.count), value
        properties.Add keyString, dictionaryArray
    Else
        Set dictionaryArray = properties.item(keyString)
        dictionaryArray.Add CStr(dictionaryArray.count), value
        Set properties.item(keyString) = dictionaryArray
    End If
    Set private_data.item(PartTitle) = properties
End Sub

'To get array data in the form of a dictionary (small call)
Public Function GD(ByVal sectionName As String, Optional propertyName As String = "") As Object
    Set GD = GetDictionary(sectionName, propertyName)
End Function

'To get array data in the form of a dictionary
Public Function GetDictionary(ByVal sectionName As String, Optional propertyName As String = "") As Object
    Set GetDictionary = CreateObject("Scripting.Dictionary")
    GetDictionary.compareMode = private_compareMode
    
    If Not CheckLocation(sectionName, propertyName) Then
        Exit Function
    End If
    
    If TypeOfProperty(sectionName, propertyName) = TYPE_OF_DICTIONARY Then
        Set sectionDictionary = GetSection(sectionName)
        Set GetDictionary = sectionDictionary.item(propertyName)
        Exit Function
    End If
End Function

'To get array data in the form of a array (small call)
Public Function GA(ByVal sectionName As String, Optional propertyName As String = "")
    GA = GetArray(sectionName, propertyName)
End Function

'To get array data in the form of a array
Public Function GetArray(ByVal sectionName As String, Optional propertyName As String = "")
    If Not CheckLocation(sectionName, propertyName) Then
        Exit Function
    End If
    
    Dim tmp() As String
    
    If TypeOfProperty(sectionName, propertyName) = TYPE_OF_STRING Then
        ReDim tmp(1) As String
        tmp(0) = GetProperty(sectionName, propertyName)
        GetArray = tmp
        Exit Function
    End If
    
    If TypeOfProperty(sectionName, propertyName) = TYPE_OF_DICTIONARY Then
        Set dictionary = GetDictionary(sectionName, propertyName)
        ReDim tmp(CLng(dictionary.count)) As String
        Dim count As Long
        count = 0
        For Each Key In dictionary.Keys
            Dim item As String
            item = CStr(dictionary.item(CStr(Key)))
            tmp(count) = item
            count = count + 1
        Next
        GetArray = tmp
        Exit Function
    End If
End Function

'To export a full data dictionary
Public Function Export() As Object
    Set Export = private_data
End Function

Private Function CheckLocation(ByRef sectionName As String, ByRef propertyName As String) As Boolean
    CheckLocation = False
    If sectionName <> "" Then
        CorrectLocation sectionName, propertyName
        CheckLocation = True
    End If
End Function

Private Sub CorrectLocation(ByRef sectionName As String, ByRef propertyName As String)
    If propertyName = "" And sectionName <> "" Then
        propertyName = sectionName
        sectionName = DEFAULT_SECTION
    End If
        KeyCase sectionName
        KeyCase propertyName
End Sub

Private Sub KeyCase(ByRef keyName As String)
    If private_compareMode = vbTextCompare Then
        keyName = LCase(keyName)
    End If
End Sub

'
'     .d8888b.   .d88888b.   .d88888b.  8888888b.      888      888     888  .d8888b.  888    d8P
'    d88P  Y88b d88P" "Y88b d88P" "Y88b 888  "Y88b     888      888     888 d88P  Y88b 888   d8P
'    888    888 888     888 888     888 888    888     888      888     888 888    888 888  d8P
'    888        888     888 888     888 888    888     888      888     888 888        888d88K
'    888  88888 888     888 888     888 888    888     888      888     888 888        8888888b
'    888    888 888     888 888     888 888    888     888      888     888 888    888 888  Y88b
'    Y88b  d88P Y88b. .d88P Y88b. .d88P 888  .d88P     888      Y88b. .d88P Y88b  d88P 888   Y88b
'     "Y8888P88  "Y88888P"   "Y88888P"  8888888P"      88888888  "Y88888P"   "Y8888P"  888    Y88b
'
