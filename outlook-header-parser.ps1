##### The script uses MS Outlook API to parse the MSG files via "http://schemas.microsoft.com/mapi/proptag/0x007D001E" #####
##Hristiyan Lazarov##
# MSG files are enumerated in a current directory
$files = Get-ChildItem -Filter *.msg | % {$_.FullName}
# Enumerated files are parsed
ForEach ($file in $files)   {
    # Creates COM outlook object
    $outlook = New-Object -comobject outlook.application 
    $msg = $outlook.CreateItemFromTemplate($file)
    # Parses MSG headers via Microsoft Outlook Schema online (Internet is connection is required)
    # Exports Internet Headers to separate text files named as MSG files in the current directory
    $msg.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E") | Out-File "$file.txt" 
    }
   
