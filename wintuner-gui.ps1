# Ensure the necessary module is installed
# Import-Module wintuner

# Load .NET Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Settings file to save the package folder
$settingsFile = "$env:USERPROFILE\wingetintune_settings.json"

# Define color variables for easier usage
$colorRed = [System.Drawing.Color]::FromName("Red")
$colorYellow = [System.Drawing.Color]::FromName("Yellow")
$colorWhite = [System.Drawing.Color]::FromName("White")
$colorGreen = [System.Drawing.Color]::FromName("Green")


# Function to save the current settings to file
function Save-Settings {
    param ($usernames, $lastPackageFolder)

    # Save usernames and the last used package folder to settings
    $settings = @{
        Usernames = $usernames
        LastPackageFolder = $lastPackageFolder
    }

    $settings | ConvertTo-Json | Set-Content -Path $settingsFile
}
# Function to load settings from file
function Load-Settings {
    if (Test-Path $settingsFile) {
        $settings = Get-Content -Path $settingsFile | ConvertFrom-Json
        return $settings
    }
    # Return default settings if no settings saved
    return @{
        Usernames = @()
        LastPackageFolder = ""
    }
}

# Function to automatically select the first search result
function Select-FirstSearchResult {
    if ($comboBoxPackages.Items.Count -gt 0) {
        $comboBoxPackages.SelectedIndex = 0  # Automatically select the first item
    }
}

# Function to search Winget packages using winget.run API
function Search-WingetPackages {
    param ($query)
    $url = "https://api.winget.run/v2/packages?query=$query"
    $response = Invoke-WebRequest -Uri $url -Method Get
    $jsonContent = $response.Content

    # Replace one of the duplicate keys in the JSON string
    $jsonContent = $jsonContent -replace '"CreatedAt"', '"CreatedAt_API"'
    $jsonContent = $jsonContent -replace '"UpdatedAt"', '"UpdatedAt_API"'

    # Parse the modified JSON content
    $json = $jsonContent | ConvertFrom-Json

    if ($json -and $json.Packages) {
        return $json.Packages  # Return the list of packages
    } else {
        return @()  # Return an empty array if no packages found
    }
}

# Function to handle the package command and retry with x86 architecture if needed
function Package-App {
    param ($packageId, $packageFolder)

    # Define output and error file paths for capturing process output
    $outputFile = [System.IO.Path]::GetTempFileName()
    $errorFile = [System.IO.Path]::GetTempFileName()

    try {
        # Run the initial package command and capture output and errors
        $packageCmd = "wintuner package $packageId --package-folder `"$packageFolder`""

        # Start the process with redirection
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command", $packageCmd -RedirectStandardOutput $outputFile -RedirectStandardError $errorFile -NoNewWindow -Wait

        # Capture standard output and error
        $packageOutput = Get-Content $outputFile
        $packageError = Get-Content $errorFile

        # Check if there is any output and append to the console
        if ($packageOutput) {
            Append-ConsoleOutput -text $packageOutput -color $colorWhite
        }

        # Check if any error occurred
        if ($packageError) {
            Append-ConsoleOutput -text "Error: $packageError" -color $colorRed
            
            # Check if the error contains the "No installer found" message
            if ($packageError -like "*No installer found for*") {
                Append-ConsoleOutput -text "Retrying with x86 architecture..." -color $colorYellow
                
                # Retry with x86 architecture
                $packageCmdX86 = "wintuner package $packageId --package-folder `"$packageFolder`" --architecture X86"
                
                # Start the process again with x86 architecture
                $processX86 = Start-Process -FilePath "powershell.exe" -ArgumentList "-Command", $packageCmdX86 -RedirectStandardOutput $outputFile -RedirectStandardError $errorFile -NoNewWindow -Wait
                
                # Capture the output of the retry
                $packageOutputX86 = Get-Content $outputFile
                $packageErrorX86 = Get-Content $errorFile

                # Append the result of retry to the console output
                if ($packageOutputX86) {
                    Append-ConsoleOutput -text $packageOutputX86 -color $colorWhite
                }
                if ($packageErrorX86) {
                    Append-ConsoleOutput -text "Error: $packageErrorX86" -color $colorRed
                }
            }
        }
    }
    finally {
        # Clean up temporary files
        Remove-Item $outputFile
        Remove-Item $errorFile
    }
}

# Function to append text to the RichTextBox console output with optional color
function Append-ConsoleOutput {
    param (
        [string]$text,
        [System.Drawing.Color]$color = $colorWhite  # Default color is white
    )
    
    # Make sure the RichTextBox allows formatting and colors
    $richTextBoxConsoleOutput.SelectionStart = $richTextBoxConsoleOutput.Text.Length
    $richTextBoxConsoleOutput.SelectionColor = $color  # Apply the desired color
    $richTextBoxConsoleOutput.AppendText($text + [Environment]::NewLine)
    
    # Ensure scrolling to the latest line
    $richTextBoxConsoleOutput.SelectionStart = $richTextBoxConsoleOutput.Text.Length
    $richTextBoxConsoleOutput.ScrollToCaret()
}

# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Wintuner.GUI"
$form.Size = New-Object System.Drawing.Size(500, 600)
$form.StartPosition = "CenterScreen"

# Create a label and a ComboBox for the username input, now as the first element in the GUI
$labelUsername = New-Object System.Windows.Forms.Label
$labelUsername.Text = "Username:"
$labelUsername.AutoSize = $true
$labelUsername.Location = New-Object System.Drawing.Point(20, 20)  # Positioning it at the top
$form.Controls.Add($labelUsername)

$comboBoxUsername = New-Object System.Windows.Forms.ComboBox
$comboBoxUsername.Location = New-Object System.Drawing.Point(120, 20)  # Corresponding text box position
$comboBoxUsername.Size = New-Object System.Drawing.Size(250, 20)
$comboBoxUsername.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
$form.Controls.Add($comboBoxUsername)

# Create controls for package search input
$labelSearch = New-Object System.Windows.Forms.Label
$labelSearch.Text = "Search Package:"
$labelSearch.AutoSize = $true
$labelSearch.Location = New-Object System.Drawing.Point(20, 60) 
$form.Controls.Add($labelSearch)

$textBoxSearch = New-Object System.Windows.Forms.TextBox
$textBoxSearch.Location = New-Object System.Drawing.Point(120, 60)
$textBoxSearch.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($textBoxSearch)

# Create a button for searching packages
$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = "Search"
$buttonSearch.Location = New-Object System.Drawing.Point(380, 60)
$form.Controls.Add($buttonSearch)

# Create a combo box to show search results
$comboBoxPackages = New-Object System.Windows.Forms.ComboBox
$comboBoxPackages.Location = New-Object System.Drawing.Point(120, 100)
$comboBoxPackages.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($comboBoxPackages)

# Add new category dropdown (for Intune category)
$labelCategory = New-Object System.Windows.Forms.Label
$labelCategory.Text = "Category:"
$labelCategory.AutoSize = $true
$labelCategory.Location = New-Object System.Drawing.Point(20, 220)
$form.Controls.Add($labelCategory)

$comboBoxCategory = New-Object System.Windows.Forms.ComboBox
$comboBoxCategory.Location = New-Object System.Drawing.Point(120, 220)
$comboBoxCategory.Size = New-Object System.Drawing.Size(250, 20)
$comboBoxCategory.Items.AddRange(@("Business", "Computerverwaltung", "Datenverwaltung", "Entwicklung & Design", "Fotos und Medien", "Utilities", "Produktivität", "Zusammenarbeit und soziale Netzwerke"))
$comboBoxCategory.SelectedIndex = 0  # Select the first item by default
$form.Controls.Add($comboBoxCategory)


# Create controls for package folder input
$labelPackageFolder = New-Object System.Windows.Forms.Label
$labelPackageFolder.Text = "Package Folder:"
$labelPackageFolder.AutoSize = $true
$labelPackageFolder.Location = New-Object System.Drawing.Point(20, 140)
$form.Controls.Add($labelPackageFolder)

$textBoxPackageFolder = New-Object System.Windows.Forms.TextBox
$textBoxPackageFolder.Location = New-Object System.Drawing.Point(120, 140)
$textBoxPackageFolder.Size = New-Object System.Drawing.Size(250, 20)
$form.Controls.Add($textBoxPackageFolder)

# Load the saved package folder from settings
$textBoxPackageFolder.Text = Load-Settings

# Create a button for selecting the package folder
$buttonBrowseFolder = New-Object System.Windows.Forms.Button
$buttonBrowseFolder.Text = "Browse"
$buttonBrowseFolder.Location = New-Object System.Drawing.Point(380, 140)
$buttonBrowseFolder.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textBoxPackageFolder.Text = $folderBrowser.SelectedPath
    }
})
$form.Controls.Add($buttonBrowseFolder)

# Create radio buttons for deployment type selection
$labelDeploymentType = New-Object System.Windows.Forms.Label
$labelDeploymentType.Text = "Deployment Type:"
$labelDeploymentType.AutoSize = $true
$labelDeploymentType.Location = New-Object System.Drawing.Point(20, 180)
$form.Controls.Add($labelDeploymentType)

$radioAvailable = New-Object System.Windows.Forms.RadioButton
$radioAvailable.Text = "Available"
$radioAvailable.Location = New-Object System.Drawing.Point(120, 180)
$radioAvailable.Checked = $true  # Default selection
$form.Controls.Add($radioAvailable)

$radioRequired = New-Object System.Windows.Forms.RadioButton
$radioRequired.Text = "Required"
$radioRequired.Location = New-Object System.Drawing.Point(220, 180)
$form.Controls.Add($radioRequired)

$radioNone = New-Object System.Windows.Forms.RadioButton
$radioNone.Text = "None"
$radioNone.Location = New-Object System.Drawing.Point(320, 180)
$form.Controls.Add($radioNone)

# Create a RichTextBox for console output
$richTextBoxConsoleOutput = New-Object System.Windows.Forms.RichTextBox
$richTextBoxConsoleOutput.Multiline = $true
$richTextBoxConsoleOutput.ScrollBars = "Both"  # Enable both vertical and horizontal scrollbars
$richTextBoxConsoleOutput.Font = New-Object System.Drawing.Font("Consolas", 10)  # Use monospaced font like PowerShell
$richTextBoxConsoleOutput.BackColor = [System.Drawing.Color]::Black  # PowerShell-like background color
$richTextBoxConsoleOutput.ForeColor = $colorWhite  # Default text color
$richTextBoxConsoleOutput.Location = New-Object System.Drawing.Point(20, 280)
$richTextBoxConsoleOutput.Size = New-Object System.Drawing.Size(440, 250)
$form.Controls.Add($richTextBoxConsoleOutput)



# Create a button to trigger the package and publish commands
$buttonExecute = New-Object System.Windows.Forms.Button
$buttonExecute.Text = "Publish"
$buttonExecute.Location = New-Object System.Drawing.Point(120, 250)

$buttonExecute.Add_Click({
    $selectedPackageText = $comboBoxPackages.SelectedItem
    $packageFolder = $textBoxPackageFolder.Text
    $selectedCategory = $comboBoxCategory.SelectedItem
    $selectedUsername = $comboBoxUsername.Text


    # Add the username to the list if it's not already present
    if (-not $comboBoxUsername.Items.Contains($selectedUsername)) {
        $comboBoxUsername.Items.Add($selectedUsername)
    }

    # Save the updated settings (usernames and package folder)
    $usernamesList = @($comboBoxUsername.Items | ForEach-Object { $_ })
    Save-Settings -usernames $usernamesList -lastPackageFolder $packageFolder

    if (-not $selectedPackageText) {
        [System.Windows.Forms.MessageBox]::Show("Please select a package.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }
    if (-not $packageFolder) {
        [System.Windows.Forms.MessageBox]::Show("Package Folder is required.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    # Extract the package ID from the selected item
    $packageId = $selectedPackageText.Split(" - ")[0]

    # Determine deployment type
    if ($radioAvailable.Checked) {
        $deploymentOption = "--available alldevices"
    } elseif ($radioRequired.Checked) {
        $deploymentOption = "--required alldevices"
    } else {
        $deploymentOption = ""
    }

    Append-ConsoleOutput -text "Packaging the app for Intune..." -color $colorGreen

    # Call the Package-App function to package the app with error handling
    Package-App -packageId $packageId -packageFolder $packageFolder

    Append-ConsoleOutput -text "Publishing the app to Intune..." -color $colorGreen

    # Add category to the publish command
    $publishCmd = "wintuner publish $packageId --package-folder `"$packageFolder`" $deploymentOption --category `"$selectedCategory`" --username `"$selectedUsername`""
    $publishOutput = Invoke-Expression $publishCmd
    Append-ConsoleOutput -text $publishOutput -color $colorWhite
})

$form.Controls.Add($buttonExecute)

# Handle package search
$buttonSearch.Add_Click({
    $query = $textBoxSearch.Text
    if (-not $query) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a search query.", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        return
    }

    $comboBoxPackages.Items.Clear()
    Append-ConsoleOutput -text "Searching for packages..."

    try {
        # Calling the search function and getting the packages
        $packages = Search-WingetPackages -query $query

        if ($packages.Count -eq 0) {
            Append-ConsoleOutput -text "No packages found." -color $colorRed
        } else {
            foreach ($package in $packages) {
                # Add the package ID and name to the combo box
                $comboBoxPackages.Items.Add("$($package.Id) - $($package.Latest.Name)")
                Select-FirstSearchResult
            }
            Append-ConsoleOutput -text "Found $($packages.Count) packages." -color $colorWhite
        }
    } catch {
        Append-ConsoleOutput -text "Error searching packages: $_"
    }
})

# Load saved settings
$settings = Load-Settings

# Populate ComboBox for usernames
$comboBoxUsername.Items.AddRange($settings.Usernames)
if ($comboBoxUsername.Items.Count -gt 0) {
    $comboBoxUsername.SelectedIndex = 0  # Select the most recent username by default
}

# Set the last used package folder if available
if ($settings.LastPackageFolder) {
    $textBoxPackageFolder.Text = $settings.LastPackageFolder
}


# Run the form
[void] $form.ShowDialog()
