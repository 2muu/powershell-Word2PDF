# PowerShell ConvertWord

Add-Type -AssemblyName System.Windows.Forms

$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = 'Select a Folder'
$folderBrowser.RootFolder = [Environment+SpecialFolder]::Desktop
$folderBrowser.ShowNewFolderButton = $false

$result = $folderBrowser.ShowDialog()

if ($result -eq [Windows.Forms.DialogResult]::OK) {
    $folderPath = $folderBrowser.SelectedPath
    $word_app = New-Object -ComObject Word.Application
    $word_app.Visible = $false
    
    $files = Get-ChildItem -Path $folderPath -Recurse -Filter *.doc*
    $fileCount = $files.Count
    $currentFileNumber = 0

    $files | ForEach-Object {
        $currentFileNumber++
        $docPath = $_.FullName
        $pdfPath = [System.IO.Path]::ChangeExtension($docPath, 'pdf')
        
        if (-Not (Test-Path $pdfPath)) {
            Write-Progress -PercentComplete (($currentFileNumber / $fileCount) * 100) -Status "Processing $docPath" -Activity "$currentFileNumber of $fileCount files processed"
            $doc = $word_app.Documents.Open($docPath)
            $doc.SaveAs([ref]$pdfPath, [ref]17)
            $doc.Close($false)
        }
        else {
            Write-Progress -PercentComplete (($currentFileNumber / $fileCount) * 100) -Status "Skipping $docPath (PDF already exists)" -Activity "$currentFileNumber of $fileCount files processed"
        }
    }

    $word_app.Quit()
    Write-Output "Conversion completed!"
}