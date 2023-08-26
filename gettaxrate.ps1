# Load the necessary .NET assembly for Windows Forms
Add-Type -AssemblyName System.Windows.Forms

# Create a new Windows Form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Find Sales Tax Rate'
$form.Width = 300
$form.Height = 200

# Create a label for the zip code input
$label = New-Object System.Windows.Forms.Label
$label.Text = 'Enter Zip Code:'
$label.Location = New-Object System.Drawing.Point(10, 10)
$label.Width = 100

# Create a text box for the zip code input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(110, 10)
$textBox.Width = 150

# Create a button for submitting the form
$button = New-Object System.Windows.Forms.Button
$button.Text = 'Submit'
$button.Location = New-Object System.Drawing.Point(110, 40)
$button.Width = 150

# Define what happens when the button is clicked
$button.Add_Click({
    $csvPath = "C:\taxes\merged.csv"
    if (Test-Path $csvPath) {
        $taxData = Import-Csv -Path $csvPath
        $userZip = $textBox.Text
        $foundEntry = $taxData | Where-Object { $_.ZipCode -eq $userZip }
        if ($foundEntry) {
            $taxRate = $foundEntry.EstimatedCombinedRate
            [System.Windows.Forms.Clipboard]::SetText($taxRate)
            [System.Windows.Forms.MessageBox]::Show("The Estimated Combined Rate for zip code $userZip is $taxRate%. This value has been copied to your clipboard.")
        } else {
            [System.Windows.Forms.MessageBox]::Show("Zip code $userZip not found in the CSV file.")
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("The CSV file at $csvPath does not exist.")
    }
})

# Add controls to the form
$form.Controls.Add($label)
$form.Controls.Add($textBox)
$form.Controls.Add($button)

# Display the form
$form.ShowDialog()
