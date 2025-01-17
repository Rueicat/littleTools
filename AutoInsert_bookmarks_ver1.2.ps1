cls
Set-ExecutionPolicy Unrestricted -Scope Process

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()


#Import database txt file (with "[]" marks)
$openfileDialog = New-Object System.Windows.Forms.OpenFileDialog

$openfileDialog.Title = "Select database file"
$openfileDialog.InitialDirectory = $PWD
$openfileDialog.Filter = "txt files (*.txt)|*.txt"

if ($openfileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
    $content = Get-Content -Path $openfileDialog.FileName
    $matches = [regex]::Matches($content, '(?s)\[(.*?)\]')

    $arraytextsWithBrackets = @()
    foreach ($match in $matches){
        $arraytextsWithBrackets += $match.Groups[0].Value
    }

    #Then Import ms word file(Could mutiple import the files)
    $openfileDialog.Title = "Select the word file to process"
    $openfileDialog.Filter = "word (*.doc;*.rtf)|*.doc;*.rtf"
    $openfileDialog.InitialDirectory = $PWD
    $openfileDialog.Multiselect = $true

    if ($openfileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $wordApp = New-Object -ComObject word.application

        $wordApp.Visible = $false

        ###forEach word file to process
      forEach ($file in $openfileDialog.FileNames){
            
        $document = $wordApp.Documents.Open($file)

        Write-Host "處理報告路徑"
        Write-Host $file
        Write-Host ""

        forEach ($paragraph in $document.Paragraphs){
            $text = $paragraph.Range.Text
            
            forEach ($bracketText in $arraytextsWithBrackets){
                if ($text -match [regex]::Escape($bracketText)){
                    $NobracketText = $bracketText -replace '\[|\]', ''
                 
                 ###Delete contents of bookmark without deleting bookmark
                 $range = $paragraph.Range
                 $startPosition = $range.Start + $text.IndexOf($bracketText)
                 $endPosition = $range.Start + ($text.IndexOf($bracketText) + $bracketText.Length)
                 $range.SetRange($startPosition, $endPosition)

                 $range.Text = ''

                 #set the point again
                 $range.SetRange($startPosition, $startPosition)

                 $bookmarkName = "Loc_" + $NobracketText
                 $document.Bookmarks.Add($bookmarkName, $range)
                 ###
                }
            }
        }
        $document.Save()
        $document.Close()
      }
    } else {
        Write-Host "Word file selection canceled"
    }
    $wordApp.Quit()
} else{
Write-Host "File selection canceled"
}
Write-Host ""
Read-Host "Press Enter to exit the program"
