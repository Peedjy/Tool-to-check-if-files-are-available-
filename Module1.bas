Attribute VB_Name = "Module1"
'///////////////////////////////////////////
'///////////////////////////////////////////
'
'Written by Jean-Pierre Crozemarie
'January 2024
'Version : Demo
'Description: This tools help me to quickly check if files are available in different folders.
'
'
'
'///////////////////////////////////////////
'///////////////////////////////////////////

Function IsFileAvailable(FileType)
Dim myDate As String

'Get the date
myDate = Format(Now, "YYYYMMDD")
'MsgBox ("Variable myDate: " & myDate)

'DEBUG :Check if FileType is received by the function
'MsgBox "File type is " & FileType

'Path used for demo only, replace by your path
PathFT = "\\server_SAP_EU\folder Toto\Folder TiTi\Folder Tata\FT Report Folder\" & myDate & "\"      'Attention these reports are saved in a specific folder (date in format YYYYMMDD)
PathFTNIO = "\\server_SAP_EU\folder Toto\Folder TiTi\Folder Tata\FT Report Folder\" & myDate & "\"   'Attention these reports are saved in a specific folder (date in format YYYYMMDD)
PathPT = "\\server_SAP_EU\folder Toto\Folder TiTi\Folder Tata\PT Report Folder\"
PathSPD = "\\server_SAP_EU\folder Toto\Folder TiTi\Folder Tata\SPD Report Folder\"
PathDR = "\\server_SAP_EU\folder Toto\Folder TiTi\Folder Tata\DR Report Folder\"
PathFB = "\\server_SAP_EU\folder Toto\Folder TiTi\Folder Tata\FB Report Folder\"

'Report Names for démo only
File_FT = "FT_Report_FR.zip"  'In this case files are zipped
File_SPD = "SPD_Report_FR.xlsx"
File_PT = "PT_Report_FR.xlsx"
File_DR = "DR_Report_FR.xlsx"
File_FB = "FB_Report_FR.xlsx"


Select Case FileType
    Case "SPD"
    
        Path = Path_SPD
        File = File_SPD
        Check_File Path, File 'Note: function check takes two arguments, make sure to pass them without parentheses.
    
    Case "FT"
    ' In case of FT we need to check both FT and FTNIO
    
        Path = Path_FT
        File = File_FT
        Check_File Path, File
        
        Path = Path_FTNIO
        File = File_FTNIO
        Check_File Path, File
        
    Case "PT"
        Path = Path_PT
        File = File_PT
        Check_File Path, File
    
    Case Else
        MsgBox "File type is not recognized"
End Select


End Function


Function Check_File(Path, File)
FileExiste = (Dir(Path & File) <> "") 'Return True or False
'DEBUG
'MsgBox "File_Existe is " & FileExiste

    If FileExiste Then
        MsgBox File & " exists in directory", vbInformation
    ' if you need to open the folder to get the file uncoment the line below
       'Shell "explorer.exe """ & Path & """", vbNormalFocus
        
    Else
        MsgBox "File is not yet available", vbExclamation
    End If
End Function

