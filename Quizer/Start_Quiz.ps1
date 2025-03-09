# Funksjon for å vise bilde
function Show-Image {
    param([string]$imagePath)

    if (-not (Test-Path $imagePath)) {
        Write-Host "Image file not found: $imagePath" -ForegroundColor Red
        return $null
    }

    Add-Type -AssemblyName System.Windows.Forms
    $image = [System.Drawing.Image]::FromFile($imagePath)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Image Viewer'
    $form.AutoSize = $true
    $form.AutoSizeMode = 'GrowAndShrink'

    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.SizeMode = 'Zoom'
    $pictureBox.Width = 768
    $pictureBox.Height = ($image.Height / $image.Width) * $pictureBox.Width
    $pictureBox.Image = $image
    $form.Controls.Add($pictureBox)

    # Viser formen uten å bruke ShowDialog
    $form.Show()
    return $form
}

# Hent stien til scriptets katalog
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Hent alle CSV-filer i scriptets katalog
$csvFiles = Get-ChildItem -Path $scriptPath -Filter "*.csv"

# Sjekk om det finnes CSV-filer
if ($csvFiles.Count -eq 0) {
    Write-Host "No CSV files found in the directory." -ForegroundColor Red
    exit
}

# Bruker velger en quizfil
Write-Host "Vennligst velg en quizfil ved å taste inn nummeret:"
for ($i = 0; $i -lt $csvFiles.Count; $i++) {
    Write-Host "$($i + 1): $($csvFiles[$i].Name)"
}
$userChoice = Read-Host "Tast inn ditt valg (1-$($csvFiles.Count))"

# Valider brukerinput
if ($userChoice -lt 1 -or $userChoice -gt $csvFiles.Count) {
    Write-Host "Ugyldig valg." -ForegroundColor Red
    exit
}

# Last inn den valgte CSV-filen
$selectedCsvFile = $csvFiles[$userChoice - 1].FullName

# Funksjon for å prøve import med forskjellige delimiters
function TryImportCsv {
    param([string]$filePath, [string]$delimiter)
    try {
        $csv = Import-Csv $filePath -Delimiter $delimiter
        return $csv
    } catch {
        return $null
    }
}

# Funksjon for å sjekke om CSV inneholder nødvendige kolonner
function CheckCsvColumns {
    param([pscustomobject[]]$csvData, [string[]]$requiredColumns)
    if ($csvData -eq $null -or $csvData.Count -eq 0) {
        return $false
    }

    foreach ($col in $requiredColumns) {
        if (-not ($csvData[0].PSObject.Properties.Name -contains $col)) {
            return $false
        }
    }

    return $true
}

# Last inn den valgte CSV-filen
$selectedCsvFile = $csvFiles[$userChoice - 1].FullName

# Prøv å importere med komma som delimiter, deretter med semikolon hvis det feiler
$questions = TryImportCsv -filePath $selectedCsvFile -delimiter ","
if (-not (CheckCsvColumns -csvData $questions -requiredColumns $requiredColumns)) {
    $questions = TryImportCsv -filePath $selectedCsvFile -delimiter ";"
}

# Sjekk om importen var vellykket
if (-not (CheckCsvColumns -csvData $questions -requiredColumns $requiredColumns)) {
    Write-Host "CSV file is missing required columns or could not be imported with known delimiters." -ForegroundColor Red
    exit
}

# Initialiser tellere
$fullyCorrectAnswers = 0
$partiallyCorrectAnswers = 0
$incorrectAnswers = 0

# Gjennomfør quizen i en uendelig løkke
while ($true) {
    # Velg et tilfeldig spørsmål
    $item = Get-Random -InputObject $questions

    # Vis separator før spørsmålet
    Write-Host "----------------------------------------------------------------------------------------------------------" -ForegroundColor Cyan

    # Vis spørsmålet
    $questionText = $item.Question
    Write-Host "`n$questionText`n"

    # Initialiser $form variabelen
    $form = $null

    # Vis bilde hvis tilgjengelig
    if (![string]::IsNullOrEmpty($item.ImagePath)) {
        $imagePath = Join-Path -Path $scriptPath -ChildPath $item.ImagePath
        $form = Show-Image -imagePath $imagePath
    }

    # Vis alternativene, hvert på en ny linje
    $options = $item.Options -split ', '
    foreach ($option in $options) {
        Write-Host "$option"
    }

    # Hent brukerens svar
    Write-Host "`nSelect an answer(s) separated by spaces (type 'end' to quit): "
    $userInput = Read-Host

    # Lukk bildet hvis det er åpent
    if ($form -ne $null) {
        $form.Close()
    }

    # Sjekk om brukeren ønsker å avslutte quizen
    if ($userInput -eq "end") {
        Write-Host "Quiz ended by user." -ForegroundColor Yellow
        break
    }

    # Behandle svarene
    $userAnswers = $userInput -split ' '

    # Konverter brukerens svar til en kommaseparert streng
    $userAnswerString = $userAnswers -join ','

    $correctAnswersArray = $item.'Correct Answer'.ToUpper().Split(',').Trim() | Sort-Object

    # Sjekk om brukerens svar stemmer overens med de riktige svarene
    if ($userAnswerString -eq ($correctAnswersArray -join ',')) {
        Write-Host "`nFully Correct!`n" -ForegroundColor Green
        $fullyCorrectAnswers++
    } elseif ($userAnswers | Where-Object { $_ -in $correctAnswersArray }) {
        Write-Host "`nPartially Correct. You answered correctly but missed some correct options.`n" -ForegroundColor Yellow
        $partiallyCorrectAnswers++
    } else {
        Write-Host "`nIncorrect. The correct answer(s) is/are: $($correctAnswersArray -join ', ')`n" -ForegroundColor Red
        $incorrectAnswers++
    }

    # Vis forklaringen gradvis
    $explanation = $item.Explanation -split '\n'
    foreach ($line in $explanation) {
        Write-Host $line -ForegroundColor Blue  # Endret til blå farge
        Start-Sleep -Seconds 1  # Forsinkelse mellom forklaringslinjene
    }

    # Slutten av spørsmålsløkken
}

# Vis quizresultatene
Write-Host "`nQuiz completed! You answered $fullyCorrectAnswers fully correct, $partiallyCorrectAnswers partially correct, and $incorrectAnswers incorrect.`n" -ForegroundColor Cyan
