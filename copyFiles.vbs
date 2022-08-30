'   Author: Flatrock Designs
'   Site: FlatrockDesigns.com
'   Contact: contact@flatrockdesigns.com
Dim FSO
Set FSO = CreateObject("Scripting.FileSystemObject")
Set speech = CreateObject("SAPI.SpVoice")

sourcePath = "C:\Users\skipp\OneDrive\Documents\Scripts\"   'Source Path for the target file
file = "Birthday placeings.pdf"                             'Target File Name with extension
destPath = "C:\Users\skipp\OneDrive\Documents\Scripts\test\"'Desination path for target duplicate
speech.Speak "Settings established"                         'Audible confirmation
FSO.CopyFile sourcePath&file, destPath                      'Execute action