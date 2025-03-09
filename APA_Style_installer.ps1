# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to detect all potential paths for citation styles in Microsoft Word
function Get-StylePaths {
    $stylePaths = @(
        "$env:ProgramFiles\Microsoft Office\Root\Office16\Bibliography\Style",
        "$env:ProgramFiles(x86)\Microsoft Office\Root\Office16\Bibliography\Style",
        "$env:ProgramFiles\Microsoft Office\Office16\Bibliography\Style",
        "$env:ProgramFiles(x86)\Microsoft Office\Office16\Bibliography\Style",
        "$env:APPDATA\Microsoft\Bibliography\Style"
    )

    $existingPaths = @()

    foreach ($path in $stylePaths) {
        if (Test-Path -Path $path) {
            $existingPaths += $path
        }
    }

    if ($existingPaths.Count -eq 0) {
        Write-Output "No existing Office citation style directories were found."
    }

    return $existingPaths
}

# Function to check if a specific .xsl style is installed in a given path
function Test-Style {
    param (
        [string]$styleName,
        [string]$styleFileName,
        [array]$stylePaths
    )
    
    $status = "Not Installed"
    foreach ($stylePath in $stylePaths) {
        $fullPath = "$stylePath\$styleFileName"
        if (Test-Path -Path $fullPath) {
            $status = "Installed"
            break
        }
    }

    return $status
}

# Function to download and install a citation style with progress indication
function Install-Style {
    param (
        [string]$styleName,
        [string]$styleFileName,
        [string]$styleUrl,
        [array]$stylePaths
    )

    $installSuccess = $false
    $webClient = New-Object System.Net.WebClient

    foreach ($stylePath in $stylePaths) {
        $fullPath = "$stylePath\$styleFileName"

        Write-Output "Downloading $styleName from $styleUrl..."
        try {
            # Download the file with progress
            $webClient.DownloadFile($styleUrl, $fullPath)
            Write-Output "$styleName downloaded successfully at $fullPath."

            # Verify if the file was successfully installed
            if (Test-Path -Path $fullPath) {
                Write-Output "$styleName installed successfully at $fullPath."
                $installSuccess = $true
            } else {
                Write-Error "Verification failed: $styleName was not found at $fullPath after installation attempt."
            }
        } catch {
            Write-Error "Failed to install ${styleName} at ${stylePath}: $($_.Exception.Message)"
            $installSuccess = $false
        }
    }

    return $installSuccess
}

# Function to remove a citation style from all detected paths
function Remove-Style {
    param (
        [string]$styleName,
        [string]$styleFileName,
        [array]$stylePaths
    )

    $removeSuccess = $true

    foreach ($stylePath in $stylePaths) {
        $fullPath = "$stylePath\$styleFileName"
        if (Test-Path -Path $fullPath) {
            Write-Output "Removing $styleName from $stylePath..."
            try {
                Remove-Item -Path $fullPath -Force -ErrorAction Stop
                Write-Output "$styleName removed successfully from $fullPath."
            } catch {
                Write-Error "Failed to remove ${styleName} from ${stylePath}: $($_.Exception.Message)"
                $removeSuccess = $false
            }
        } else {
            Write-Output "$styleName is not installed at $fullPath."
        }
    }

    return $removeSuccess
}

# Create the form
$form = New-Object Windows.Forms.Form
$form.Text = "Manage Citation Styles"
$form.Size = New-Object Drawing.Size(600,400)

# Create the DataGridView to list styles
$dataGridView = New-Object Windows.Forms.DataGridView
$dataGridView.Size = New-Object Drawing.Size(580, 300)
$dataGridView.Location = New-Object Drawing.Point(10, 10)
$dataGridView.AutoSizeColumnsMode = [Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGridView.SelectionMode = [Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridView.MultiSelect = $true

# Create columns
$col1 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$col1.HeaderText = "Name"
$col1.Name = "Name"
$col1.ReadOnly = $true
$col2 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$col2.HeaderText = "FileName"
$col2.Name = "FileName"
$col2.ReadOnly = $true
$col3 = New-Object Windows.Forms.DataGridViewTextBoxColumn
$col3.HeaderText = "Status"
$col3.Name = "Status"
$col3.ReadOnly = $true

$dataGridView.Columns.AddRange($col1, $col2, $col3)

# Create the Install button
$installButton = New-Object Windows.Forms.Button
$installButton.Text = "Install"
$installButton.Size = New-Object Drawing.Size(75, 30)
$installButton.Location = New-Object Drawing.Point(10, 320)

# Create the Remove button
$removeButton = New-Object Windows.Forms.Button
$removeButton.Text = "Remove"
$removeButton.Size = New-Object Drawing.Size(75, 30)
$removeButton.Location = New-Object Drawing.Point(100, 320)

# Populate DataGridView with styles
$stylePaths = Get-StylePaths
$styles = @(
    [PSCustomObject]@{ Name = "APA 7th Edition"; FileName = "APASeventhEdition.xsl"; Url = "https://raw.githubusercontent.com/briankavanaugh/APA-7th-Edition/main/APASeventhEdition.xsl"; Status = Test-Style "APA 7th Edition" "APASeventhEdition.xsl" $stylePaths },
    [PSCustomObject]@{ Name = "Chicago 16th Edition"; FileName = "CHICAGO.XSL"; Url = "https://raw.githubusercontent.com/citation-style-repository/chicago-author-date/main/chicago.xsl"; Status = Test-Style "Chicago 16th Edition" "CHICAGO.XSL" $stylePaths },
    [PSCustomObject]@{ Name = "MLA 7th Edition"; FileName = "MLASeventhEditionOfficeOnline.xsl"; Url = "https://raw.githubusercontent.com/citation-style-repository/mla/main/mla.xsl"; Status = Test-Style "MLA 7th Edition" "MLASeventhEditionOfficeOnline.xsl" $stylePaths }
)

foreach ($style in $styles) {
    $row = $dataGridView.Rows.Add()
    $dataGridView.Rows[$row].Cells[0].Value = $style.Name
    $dataGridView.Rows[$row].Cells[1].Value = $style.FileName
    $dataGridView.Rows[$row].Cells[2].Value = $style.Status
}

# Add event handler for the Install button
$installButton.Add_Click({
    foreach ($selectedRow in $dataGridView.SelectedRows) {
        $styleName = $selectedRow.Cells[0].Value
        $fileName = $selectedRow.Cells[1].Value
        $styleUrl = $styles | Where-Object { $_.FileName -eq $fileName } | Select-Object -ExpandProperty Url

        if ($styleUrl) {
            $result = Install-Style -styleName $styleName -styleFileName $fileName -styleUrl $styleUrl -stylePaths $stylePaths
            if ($result) {
                $selectedRow.Cells[2].Value = "Installed"
            } else {
                $selectedRow.Cells[2].Value = "Failed to Install"
            }
        } else {
            Write-Error "URL not found for $styleName."
        }
    }
})

# Add event handler for the Remove button
$removeButton.Add_Click({
    foreach ($selectedRow in $dataGridView.SelectedRows) {
        $styleName = $selectedRow.Cells[0].Value
        $fileName = $selectedRow.Cells[1].Value
        
        $result = Remove-Style -styleName $styleName -styleFileName $fileName -stylePaths $stylePaths
        if ($result) {
            $selectedRow.Cells[2].Value = "Not Installed"
        } else {
            $selectedRow.Cells[2].Value = "Failed to Remove"
        }
    }
})

# Add controls to the form
$form.Controls.Add($dataGridView)
$form.Controls.Add($installButton)
$form.Controls.Add($removeButton)

# Show the form
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

# Final message to indicate completion
Write-Output "Citation style management process completed."
