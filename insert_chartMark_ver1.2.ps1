cls
Set-ExecutionPolicy Unrestricted -scope Process

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()


$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Filter = "Text files (*.txt)|*.txt"
$OpenFileDialog.Title = "Import setting file"
$OpenFileDialog.InitialDirectory = $PWD


if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $dictionaryPath = $OpenFileDialog.FileName
} else {
    Write-Host "File Selection canceld"
    exit
}


$bookmarkMapping = @{}
Get-Content $dictionaryPath | ForEach-Object {
    if ($_ -match '^(.*?) = (.*?)$') {
        $bookmarkMapping[$matches[1]] = $matches[2]
    }
}


$FileDialog = New-Object System.Windows.Forms.OpenFileDialog
$FileDialog.Filter = "Word files (*.doc;*.rtf)|*.doc;*.rtf"
$FileDialog.Multiselect = $true
$FileDialog.Title = "Change bookmark file"
$FileDialog.InitialDirectory = $PWD


if ($FileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $wordFiles = $FileDialog.FileNames
} else {
    Write-Host "WOrd Selection canceled"
    exit
}


$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false


$counter = 0
$totalFiles = $wordFiles.Count

foreach ($file in $wordFiles) {
    $counter++
    Write-Progress -Activity "processing..." -Status "process $file ($counter/$totalFiles)" -PercentComplete (($counter / $totalFiles) * 100)
    
    $document = $wordApp.Documents.Open($file)

    
    foreach ($bookmark in $document.Bookmarks) {
        $bookmarkName = $bookmark.Name

        
        if ($bookmarkMapping.ContainsKey($bookmarkName)) {
            $newBookmarkName = $bookmarkMapping[$bookmarkName]
            Write-Host "rename bookmarkName from '$bookmarkName' to '$newBookmarkName' ..."
            
            $bookmark.Range.Bookmarks.Add($newBookmarkName) | Out-Null
            $bookmark.Delete()  
        } else {
            Write-Host "'$bookmarkName' passed"
        }
    }

    
    $document.Save()
    $document.Close()
}


$wordApp.Quit()

Write-Host "Finished all the process"
Read-Host "Press Enter key to exit the program"
