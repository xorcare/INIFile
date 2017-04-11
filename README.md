# INIFile DLL (INIFI)

## Class for parsing ini files in Visual Basic 5-6
**Also the library INIFile.dll is compatible with 1C 77**

### Set/Get path to ini file

~~~
Public Sub SetPath(inputPath)
Public Function GetPath() As String
~~~

### Reading and parsing a ini file

~~~
Public Function Load() As Boolean
Public Function LoadFrom(filePath) As Boolean
~~~

### Loaded data from a file

~~~
Public Function Loaded() As Boolean
~~~

### Checks whether there is property inside the section

~~~
Public Function E(sectionName, Optional propertyName) As Boolean
Public Function ExistProperty(sectionName, Optional propertyName) As Boolean
~~~

### Returns the value of the property in a certain section

~~~
Public Function G(sectionName, Optional propertyName) As String
Public Function GetProperty(sectionName, Optional propertyNam) As String
~~~

### Dictionary CompareMode

~~~
Public Function GetCompareMode() As Byte
Public Function SetCompareMode(mode) As Boolean
~~~

### Get the type of properties

~~~
Public Function TOP(sectionName, Optional propertyName) As Integer
Public Function TypeOfProperty(sectionName, Optional propertyName) As Integer
~~~

**Const types:**

 * TYPE_OF_NOT_DEFINED = -1
 * TYPE_OF_DICTIONARY = 2
 * TYPE_OF_NOT_EXIST = 0
 * TYPE_OF_STRING = 1

### Get array data in the form of a dictionary

~~~
Public Function GD(sectionName, Optional propertyName) As Object
Public Function GetDictionary(sectionName, Optional propertyName) As Object
~~~

### Get array data in the form of a array

~~~
Public Function GA(sectionName, Optional propertyName)
Public Function GetArray(sectionName, Optional propertyName)
~~~

### Class properties

**Format:**
* 0 = TristateFalse - Default. Open the file as ASCII
* -1 = TristateTrue - Open the file as Unicode
* -2 = TristateUseDefault - Open the file using the system default

~~~
Public Format As Integer
~~~

**Create the file if it does not exist**
* True - Create
* False - Ignore

~~~
Public Create As Boolean
~~~

### To export a full data dictionary

~~~
Public Function Export() As Object
~~~

## Example

~~~
Set FSO = CreateObject("Scripting.FileSystemObject")
Set ini = CreateObject("INIFile.INIFI")
ini.SetPath (FSO.GetAbsolutePathName(FSO.BuildPath(app.Path, app.EXEName & ".ini")))
ini.Create = True
ini.Format = -2 'Cyrillic
status = ini.SetCompareMode(vbTextCompare)
MsgBox CStr(ini.Load())
~~~