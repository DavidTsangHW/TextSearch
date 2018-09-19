'26 AUG 2018
'TEXT SEARCH
'SEARCH TEXT BY SEARCHING THROUGH ALL FOLDER AND SUBFOLDERS

'David Tsang
'Linkedin: https://www.linkedin.com/in/david-tsang-306b7433/

'***********************IMPORTANT******************************

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301, USA.

'USAGE
'1. COPY THIS VBSCRIPT TO THE ROOT FOLDER TO BE SEARCHED
'2. ENTER SEARCH TEXT
'3. ENTER FILE TYPE, IE. TXT, CPP, JAVA, PY...
'4. SIT BACK AND RELAX
'5. FIND THE SEARCH RESULT IN RESULT.XLS

Dim message
Dim answer 
Dim CurrentDirectory
Dim fs

Dim fname 
Dim fileName
Dim headerFlag


Dim MyFiles
Dim MyFile
Dim validFiles
Dim filescount
Dim searchText

Dim searchResult
Dim instrResult

Dim resultCount
Dim programname

programname = "Text search"

resultCount = 0

set fs=CreateObject("Scripting.FileSystemObject")

outputFilename = "result.xls"

searchText = inputbox("This vbscript finds text by searching through all folders and subfolders." & vbcrlf & vbcrlf & "Enter text",programname,"")

if len(searchText) = 0 then
	
	Msgbox "No search text is entered." & vbcrlf & "Program aborted", 0, programname
	Wscript.quit

end if

filetype = inputbox("Enter file type (ie, txt, cpp, java, py, vbs)", programname, "py")

set fname=fs.CreateTextFile(outputFilename,true)

CurrentDirectory = fs.GetAbsolutePathName(".")

call ReadFiles(fs.GetFolder(CurrentDirectory)) 

sub ReadFiles(MyFolder)
    
    Dim SubFolder
    
    For Each SubFolder In MyFolder.SubFolders
		On Error Resume Next
		call ReadFiles(SubFolder)
    Next
	
	
	For Each MyFile In MyFolder.Files
	
		GetAnExtension = fs.GetExtensionName(MyFile.name)
		GetAnExtension = lcase(GetANExtension)
		
		if lcase(outputFilename) <> MyFile.name then

			if ((myFile.Attributes And 2) <> 2) and (GetANExtension = filetype) then
				path = fs.GetAbsolutePathName(MyFolder) & "\" & MyFile.name 
				Set objFile = fs.OpenTextFile(path, 1)
				linenum = 1
				Do Until objFile.AtEndOfStream
					myLine = objFile.ReadLine	
					instrResult = instr(ucase(myLine), ucase(searchText)) 
					if instrResult > 0 then
						fname.writeLine path & "," & linenum & "," & myLine 
						resultCount = resultCount + 1
					end if
					linenum = linenum + 1
				Loop
				objFile.Close
			end if
		
		end if 
	Next

end sub

msgbox resultCount & " results" & vbCrlf & vbCrlf & "See results.xls", 0, programname

set fname=nothing
