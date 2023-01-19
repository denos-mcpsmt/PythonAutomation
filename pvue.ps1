$FolderName = "C:\Users\Pearson\Desktop\Pearson Vue\"
if (Test-Path $FolderName) {
    Write-Host "Folder Exists, removing"
    Remove-Item $FolderName -Force
}

New-Item -Path 'C:\Users\Pearson\Desktop\Pearson Vue' -ItemType Directory
python "C:\Users\Pearson\PycharmProjects\Python Automation\autoWord.py"

$cont= Read-Host -Prompt "Would you like to send pdf files to printer? (y/n)"

if ($cont -eq "n"){
    exit
}

$pdffiles = Get-ChildItem $FolderName -Filter *.pdf
foreach ($f in $pdffiles){
    Start-Process -FilePath $f.FullName -Verb print
}

