# Word uygulamasını başlat ve görünmez yap
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Belirli bir klasördeki tüm Word belgelerini işle
$folderPath = "C:\desktop"
$files = Get-ChildItem -Path $folderPath -Filter *.docx

foreach ($file in $files) {
    Write-Host "Processing file: $($file.FullName)"
    try {
        # Her belgeyi aç
        $document = $word.Documents.Open($file.FullName)
        
        # Belirli bir metni alt bilgilerde ara ve sil
        $textFound = $false
        foreach ($section in $document.Sections) {
            foreach ($footer in $section.Footers) {
                $range = $footer.Range
                $find = $range.Find
                $find.Text = "!!!Your Keyword!!!"
                $find.Replacement.Text = ""
                $find.Forward = $true
                $find.Wrap = 1 # wdFindContinue
                $find.Format = $false
                $find.MatchCase = $false
                $find.MatchWholeWord = $true
                $find.MatchWildcards = $false
                $find.MatchSoundsLike = $false
                $find.MatchAllWordForms = $false

                $result = $find.Execute()
                if ($result -eq $true) {
                    $textFound = $true
                    $range.Text = $find.Replacement.Text
                    Write-Host "Text found and replaced in footer of $($file.FullName)"
                } else {
                    Write-Host "Text not found in footer of $($file.FullName)"
                }
            }
        }
        
        if ($textFound) {
            # Değişiklikleri kaydet
            $document.Save()
        }
        # Belgeyi kapat
        $document.Close()
    } catch {
        Write-Host "Error processing file: $($file.FullName) - $_"
    }
}

# Word uygulamasını kapat
$word.Quit()

# COM nesnelerini serbest bırak
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable word
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()