################### Welcome to the Work Notes Clipboard Script! This was written by Ambrosio Gomez IV, 2025 #####################


##################################################### Settings File Logic #########################################################

# Points the program to the settings file
$settingsPath = "$PSScriptRoot\userSettings.json"

# Establishes userSettings dictionary and sets a default settings profile
$global:userSettings = @{
sort = "Alpha"
autoPaste = $true
delayTime = 1.0
fontSize = 8
hoverTabs = $true
}

# If a settings file is present, assign its contents to the userSetetings dictionary
if (Test-Path $settingsPath){
$global:userSettings = Get-Content $settingsPath | ConvertFrom-Json
}

########################################################### Helpers #############################################################

# Function: establish a default button template to be used accross the program
# Parameters: $Text is the contents of the button, $X is the X coordinate location, $Y is the Y coordinate location
# Output: Returns a button object with with the given parameters
function New-Button {

param( [string]$Text , [int]$X , [int]$Y , [bool]$CopyPaste )

$button = New-Object System.Windows.Forms.Button
        $button.Text = $Text
        $button.Size = New-Object System.Drawing.Size(150, 60)
        $button.Location = New-Object System.Drawing.Point($X, $Y)
        $button.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:userSettings.fontSize, [System.Drawing.FontStyle]::Bold)

# automatically assigns the button to CopyAndPaste on click, can be reassigned manually

if ($CopyPaste){
        $button.Add_Click([scriptblock]::Create("CopyAndPaste '$Text'"))
}

return $button
}

# Function: establish a default label template to be used accross the program
# Parameters: $Text is the contents of the button, $X is the X coordinate location, $Y is the Y coordinate location
# Output: Returns a label object with with the given parameters
function New-Label {

param( [string]$Text , [int]$X , [int]$Y )

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point($X, $Y)
$label.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:userSettings.fontSize, [System.Drawing.FontStyle]::Bold)

    $label.AutoSize = $true
    $label.Text = $text

return $label

}

# Function: establish a default radio template to be used accross the program
# Parameters: $Text is the contents of the button, $X is the X coordinate location, $Y is the Y coordinate location
# Output: Returns a radio object with with the given parameters
function New-Radio {
param( [string]$Text , [int]$X , [int]$Y )

$radio = New-Object System.Windows.Forms.RadioButton
$radio.Text = $Text
$radio.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:userSettings.fontSize)
$radio.Location = New-Object System.Drawing.Point($X, $Y)

return $radio

}

# Function: establish a default text box template to be used accross the program
# Parameters: $X is the X coordinate location, $Y is the Y coordinate location
# Output: Returns a text box object with with the given parameters
function New-Input {

param ( [int]$X, [int]$Y )

$input = New-Object System.Windows.Forms.TextBox
$input.Location = New-Object System.Drawing.Point($X, $Y)
$input.Width = 300

return $input

}

function New-ListBox {

param ( [string[]]$ItemsArray, [int]$X, [int]$Y )

$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point($X, $Y)
$listBox.Size = New-Object System.Drawing.Size(300,60)
$listBox.Font = New-Object System.Drawing.Font("Microsoft Sans Serif", $global:userSettings.fontSize)
$listBox.SelectionMode = "One" 
$listBox.Items.AddRange($ItemsArray)

return $listBox

}

########################################################### Utilities #############################################################


# Function: to copy text to clipboard and paste it on the user's computer
# Parameters: $Text is the text that will be copied and/or pasted
# Output: Copies the string into the user's clipboard, and pastes it with an optional delay
function CopyAndPaste {
    # Accepts string variable
    param([string]$text)

    # Copies string to clipboard
    [System.Windows.Forms.Clipboard]::SetText($text)

    if ($global:userSettings.autoPaste){

    # Optional Delay
    Start-Sleep -Seconds $global:userSettings.delayTime

    # Sends CTRL - V (paste shortcut) to user's computer
    [System.Windows.Forms.SendKeys]::SendWait("^v")
    }

}

function Clear-DynamicControls {

param( [System.Windows.Forms.Panel]$Panel, [System.Windows.Forms.Form]$Form )
$count = 0

if ($Panel){
# Only clears incomePanel controls tagged with dynamic (the text boxes) upon each button click
    for ($i = $Panel.Controls.Count - 1; $i -ge 0; $i--) {

        $control = $Panel.Controls[$i]

        if ($control.Tag -eq 'dynamic') {
$Panel.Controls.Remove($control)
        }
    }
} elseif ($Form) {
# Only clears incomePanel controls tagged with dynamic (the text boxes) upon each button click
    for ($i = $Form.Controls.Count - 1; $i -ge 0; $i--) {

        $control = $Form.Controls[$i]

        if ($control.Tag -eq 'dynamic') {
$Form.Controls.Remove($control)
        }
    }
}
   

}

############################################# Main Tabs (Elig, Tpiv, Escl) ########################################################

# Function: adjust the sort of the notes based on user preference
# Parameters: $workNotes is the array that will be sorted
# Output: Returns a new array $newNotes that is sorted in accordance with user preference
function Get-SortedNotes {
param ( [array]$workNotes )

$newNotes = @{ }

    switch ($global:userSettings.sort) {
        "Alpha" { $newNotes = $workNotes | Sort-Object }
        "Most Used" { $newNotes = $workNotes }
    }

return $newNotes
}

function Get-TabDict {

$workNotesModPath = "$PSScriptRoot\workNotesModified.json"
$workNotesDefaultPath = "$PSScriptRoot\workNotesDefault.json"

$tabData = [ordered]@{}

if (Test-Path $workNotesModPath) {

    $json = Get-Content $workNotesModPath -Raw | ConvertFrom-Json
} elseif (Test-Path $workNotesDefaultPath) {

$json = Get-Content $workNotesDefaultPath -Raw | ConvertFrom-Json

} else {
    Write-Error "Configuration file '$worknotesPath' not found. Exiting script."
    exit 1
}

foreach ($item in $json.PSObject.Properties) {
$tabData[$item.Name] = $item.Value
}
return $tabData
}

# Function: creates the UI elements shown on the 3 main tabs
# Parameters: None
# Output: Shows the user the main tabs
function Render-Tabs {

$tabControl.TabPages.Clear()

$tabDict = Get-TabDict
# Loop that creates each tab page and button, loops for each work note tab
foreach ($tab in $tabDict.Keys){

# Creates a new tab page
    $tabPage = New-Object System.Windows.Forms.TabPage
    $tabPage.Text = $tab
    $tabPage.BackColor = [System.Drawing.Color]::FromName($tabDict[$tab].Color)

    # Create a scrollable panel
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.AutoScroll = $true

    # Instantiates the position of the Y axis for the button's creation, is incremented with every new button
    $yPos = 20

#sends the work notes to be sorted
$approvedNotes = Get-SortedNotes -workNotes $tabDict[$tab].approvedNotes
$denyNotes = Get-SortedNotes -workNotes $tabDict[$tab].denyNotes
# Loops through the approved notes to create buttons for each one
    foreach ($workNote in $approvedNotes){

$button = New-Button -Text $workNote -X 10 -Y $yPos -CopyPaste $true
$button.Anchor = 'Top,Left'
$button.BackColor = [System.Drawing.Color]::YellowGreen
$button.Tag = "Approve"
        $panel.Controls.Add($button)

        # Y position incrementor
        $yPos += 60
    }

# resets the y Position for the deny notes
    $yPos = 20

# Loops through the deny notes to create buttons for each one
    foreach ($workNote in $denyNotes){

$button = New-Button -Text $workNote -X 170 -Y $yPos -CopyPaste $true
$button.BackColor = [System.Drawing.Color]::PaleVioletRed
$button.Anchor = 'Top,Left'
$button.Tag = "Deny"
        $panel.Controls.Add($button)

        # Y position incrementor
        $yPos += 60
      }

# Adds the panel to the tab page
      $tabPage.Controls.Add($panel)

# Adds the tab page to the window
      $tabControl.TabPages.Add($tabPage)
}

# Renders the income tab
Render-IncomeTab
}

##################################################### Income Tab #################################################################

# Function: gets the average annual income based on the values of each text box and the amount of pay periods
# Parameters: $Panel is used to see the values of each text box, $Amount is the amount of pay periods
# Output: Returns $finalAmount, which is the average annual income
function Get-AverageIncome {

param ( [System.Windows.Forms.Panel]$AvgPanel, [int]$Amount )

#gets the sum of every text box
$sum = 0
foreach ($ctrl in $AvgPanel.Controls) {
if ($ctrl -is [System.Windows.Forms.TextBox] -and $ctrl.Tag -eq 'dynamic') {
[double]$val = 0
[void][double]::TryParse($ctrl.Text, [ref]$val)

if ($val -lt 9999999999){
$sum += $val
} else {

[System.Windows.Forms.MessageBox]::Show("One or more values is too high! Calculation Cancelled. ")
}
}
}

# gets the average amount from the pay period amounts
$avgAmt = $sum / $Amount

$finalAmount = 0
# calculates average annual income from the annual income, depending on pay period frequency
switch ($Amount) {
1 { $finalAmount = $avgAmt } #annually
3 { $finalAmount = $avgAmt * 12 } #monthly
6 { $finalAmount = $avgAmt * 26 } #biweekly
12 { $finalAmount = $avgAmt * 52 } #weekly
default { $finalAmount = 0 } #default
}

return $finalAmount
}

# Function: Updates the approval status with a color and message
# Parameters: $household and $annualAverage are used to determine approval status
# Output: Returns $Status, an array that holds the font color and message for the approval button based on the approval decision
function Update-ApprovalStatus {

param( [int]$Household, [int]$AnnualAverage, [string]$State )

# Household Size and their corresponding maximum income

$approvedHHAmounts = [ordered]@{

"48states" = @{
0 = 0.00
1 = 10000.00
2 = 20000.00
3 = 30000.00
4 = 40000.00
5 = 50000.00
6 = 60000.00
7 = 70000.00
8 = 80000.0
}

"Alaska" = @{
0 = 0.00
1 = 15000.00
2 = 25000.00
3 = 35000.00
4 = 45000.00
5 = 55000.00
6 = 65000.00
7 = 75000.00
8 = 85000.00
}

"Hawaii" = @{
0 = 0.00
1 = 17500.00
2 = 27500.00
3 = 37500.00
4 = 47500.00
5 = 57500.00
6 = 67500.00
7 = 77500.00
8 = 87500.00
}
}

# Initialize status dictionary
$status = @{
Message = ""
BackColor = ""
}

$correctApprovalList = @{}

switch ($State) {
"Continental US" { $correctApprovalList = $approvedHHAmounts["48States"] }
"Alaska" { $correctApprovalList = $approvedHHAmounts["Alaska"] }
"Hawaii" { $correctApprovalList = $approvedHHAmounts["Hawaii"] }
}


# Checks to see if the household size is within the range of given values, or if a new value must be calculated
if ($Household -gt 0 -and $Household -lt 100) {

# instantiates the approvable threshold
$threshold = 0

if ($Household -le 8) {
# sets the threshold to the amount listed in the dictionary for the provided household size
$threshold = $correctApprovalList[$Household]
} elseif ($Household -gt 8) {
# adds $10,000 for every added household member over 8
$adjustment = ($Household - 8) * 10000
$threshold = $correctApprovalList[8] + $adjustment
}
 
# updates the button text with approval message
if ($AnnualAverage -le $threshold) {

$status['Message'] ="ELIG - INCOME APPROVED (INCOME IS AT OR UNDER: {0:C}) " -f $threshold
$status['BackColor'] = 'YellowGreen'

} else {

$status['Message'] = "ELIG - INCOME TOO HIGH (INCOME IS ABOVE: {0:C}) " -f $threshold
$status['BackColor'] = 'PaleVioletRed'
}

} else {

# default value for Status
$status['Message'] = "Enter a valid household size"
$status['BackColor']  = 'Control'
}

return $status
}

# Function: Updates the input boxes based on the radio button pressed
# Parameters: $incomePanel is used to create/clear input boxes, $amtPayPeriods is used to see how many input boxes to create
# Output: Previous input boxes are cleared, and the correct amount of new input boxes are created
function Update-InputBoxes {

param( [System.Windows.Forms.Panel]$Panel, [int]$PayPeriods )

$yPos = 20

Clear-DynamicControls -Panel $Panel

    if ($PayPeriods -gt 0) {

$yPos += 215

# Creates text boxes based off of the amount of periods
        for ($i = 0; $i -lt $PayPeriods; $i++) {

$paymentsInput = New-Input -X 20 -Y $yPos
$paymentsInput.Tag = 'dynamic'
$Panel.Controls.Add($paymentsInput)
           
$yPos += 30
    }
}
}

function Remove-AllClickHandlers{

param ([System.Windows.Forms.Button]$button)

#Gets the private events field from the component class
$eventsField = [System.ComponentModel.Component].GetField("events",[System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Instance)

#Gets the actual eventhanderlist object from the button object
$eventHandlerList = $eventsField.GetValue($button)
#Gets the private static key taht winforms uses to identify click events
$clickEventKey = [System.Windows.Forms.Control].GetField("EventClick",[System.Reflection.BindingFlags]::NonPublic -bor [System.Reflection.BindingFlags]::Static).GetValue($null)

#Removes all ahnlders associated with the click event key
if ($eventHandlerList -ne $null -and $eventHandlerList[$clickEventKey] -ne $null){
$eventHandlerList.RemoveHandler($clickEventKey, $eventHandlerList[$clickEventKey])
}
}

# Function: creates the UI elements shown on the income tab
# Parameters: None
# Output: Shows the user the income tab
function Render-IncomeTab {
$yPos = 20

# Create Tab Page
$tabPage = New-Object System.Windows.Forms.TabPage
$tabPage.Text = "PAYSTUBS"

# Create a scrollable panel
$incomePanel = New-Object System.Windows.Forms.Panel
$incomePanel.Dock = 'Fill'
$incomePanel.AutoScroll = $true

Set-Variable -Name incomePanel -Scope Script -Value $incomePanel

# Create approval button
$approvalButton = New-Button -Text "Enter all necessary information" -X 85 -Y -7
$approvalButton.ForeColor = 'Black'
$approvalButton.Anchor = 'Bottom,Left'

# Store it in script scope
Set-Variable -Name approvalButton -Scope Script -Value $approvalButton

$incomePanel.Controls.Add($approvalButton)

# Create the label prommpting user to input household size
$hhLabel = New-Label -Text "Input the household size: " -X 20 -Y $yPos
$incomePanel.Controls.Add($hhLabel)
$yPos += 20
# Create thes input control
$hhInput = New-Input -X 20 -Y $yPos
$incomePanel.Controls.Add($hhInput)
Set-Variable -Name hhInput -Scope Script -Value $hhInput
$yPos += 30

# Create the label prommpting user to input household size
$stateLabel = New-Label -Text "Where are they located? " -X 20 -Y $yPos
$incomePanel.Controls.Add($stateLabel)
$yPos += 20

Set-Variable -Name hhState -Scope Script -Value ""
$hhStateLB = New-ListBox -ItemsArray "Continental US", "Alaska", "Hawaii" -X 20 -Y $yPos
$incomePanel.Controls.Add($hhStateLB)
Set-Variable -Name hhStateLB -Scope Script -Value $hhStateLB
$yPos += 60

# Create the label prommpting user to input household size
$pfLabel = New-Label -Text "Select the Paystub Frequency" -X 20 -Y $yPos
$incomePanel.Controls.Add($pfLabel)
$yPos += 20

# Create pay period buttonss

Set-Variable -Name amtPayPeriods -Scope Script -Value 0

$radio1 = New-Radio -Text "Annual Income" -X 20  -Y $yPos
$radio1.Add_CheckedChanged({
if($this.Checked){
Set-Variable -Name amtPayPeriods -Scope Script -Value 1
Update-InputBoxes -Panel $incomePanel -PayPeriods $amtPayPeriods
}
})

$radio3 = New-Radio -Text "Monthly Paystubs" -X 120  -Y $yPos
$radio3.Add_CheckedChanged({
if($this.Checked){
Set-Variable -Name amtPayPeriods -Scope Script -Value 3
Update-InputBoxes -Panel $incomePanel -PayPeriods $amtPayPeriods
}
})
$radio6 = New-Radio -Text "Biweekly Paystubs" -X 220  -Y $yPos
$radio6.Add_CheckedChanged({
if($this.Checked){
Set-Variable -Name amtPayPeriods -Scope Script -Value 6
Update-InputBoxes -Panel $incomePanel -PayPeriods $amtPayPeriods
}
})

$radio12 = New-Radio -Text "Weekly Paystubs" -X 320  -Y $yPos
$radio12.Add_CheckedChanged({
if($this.Checked){
Set-Variable -Name amtPayPeriods -Scope Script -Value 12
Update-InputBoxes -Panel $incomePanel -PayPeriods $amtPayPeriods
}
})
$yPos += 30
$incomePanel.Controls.AddRange(@($radio1, $radio3, $radio6, $radio12))

$approvalStatus = @{
Message = "Enter all necessary information"
FontColor = "Black"
}

# Create the calculate button
$calcButton = New-Button -Text "Calculate" -X 10 -Y 87
$calcButton.Size = New-Object System.Drawing.Size(75, 60)
$calcButton.Anchor = 'Bottom,Left'
$calcButton.Add_Click({

if ($amtPayPeriods -gt 0){

$hhSize = 0

$temp = 0
$inputText = $script:hhInput.Text.Trim()
if ([int]::TryParse($inputText, [ref]$temp)) {
$hhSize = $temp
}

$stateSelection = $hhStateLB.SelectedItem
if (-not $stateSelection) {
    [System.Windows.Forms.MessageBox]::Show("Please select a state.")
    return
}

# gets the annual average
$annualAverage = Get-AverageIncome $incomePanel $amtPayPeriods

# checks approval status
$approvalStatus = Update-ApprovalStatus $hhSize $annualAverage $stateSelection

# Update the existing button
$buttonText = [string]$approvalStatus['Message']
$approvalButton.Text = $buttonText
$approvalButton.BackColor = [System.Drawing.Color]::FromName($approvalStatus['BackColor'])

Remove-AllClickHandlers $approvalButton
$approvalButton.Add_Click([scriptblock]::Create("CopyAndPaste '$buttonText'"))

}
})

# adds calculator buttons
$incomePanel.Controls.Add($calcButton)

#creates the paystub label
$psLabel = New-Label -Text "Input the payment amount per paystub: " -X 20 -Y $yPos
$incomePanel.Controls.Add($psLabel)
$yPos += 30

# add paystubs incomePanel to the window
$tabPage.Controls.Add($incomePanel)
$tabControl.TabPages.Add($tabPage)
}

##################################################### Settings Window #############################################################

# Function: Saves user settings
# Parameters: None
# Output: Settings are saved to a JSON file, if a file is not present, one will be created
function Save-UserSettings {
$global:userSettings | ConvertTo-Json -Depth 3 | Set-Content -Path $settingsPath -Encoding UTF8
}

# Function: updates the visibility of the delay field
# Parameters: None
# Output: Shows/hides the delay field based on autoPaste toggle

function Update-DelayField {

if ($autoPasteToggle.Checked){
$global:userSettings.autoPaste = $true
Save-UserSettings
$settingsPanel.Controls.Add($delayLabel)
$settingsPanel.Controls.Add($delayInput)

} else {
$global:userSettings.autoPaste = $false
Save-UserSettings

Clear-DynamicControls -Panel $settingsPanel
}
}

# Function: creates the UI elements shown on the settings page
# Parameters: None
# Output: Shows the user the settings page
function Render-SettingsWindow {

# Create the settings Form, i.e. the settings page itself
$settingsForm = New-Object System.Windows.Forms.Form
$settingsForm.Text = "Settings"
$settingsForm.Size = New-Object System.Drawing.Size(400,350)
$settingsForm.StartPosition = "CenterParent"
$settingsForm.FormBorderStyle = "FixedDialog"
$settingsForm.TopMost = $true
$settingsForm.Add_FormClosing({ Save-UserSettings })

# Create a scrollable panel
    $settingsPanel = New-Object System.Windows.Forms.Panel
    $settingsPanel.Dock = 'Fill'
    $settingsPanel.AutoScroll = $true

$settingsYPos = 20

# Create the Sort Label
$sortLabel = New-Label -Text "Choose your preferrred sort: " -X 20 -Y $settingsYPos
$settingsPanel.Controls.Add($sortLabel)
$settingsYPos += 20

# Create Sort buttons
$radioAlpha = New-Radio -Text "Alphabetical Order" -X 20 -Y $settingsYPos
$radioMostUsed= New-Radio -Text "Most Used Order" -X 125 -Y $settingsYPos


# On button press, sort depending on button pressed, save the user's preference, and reload the tabs
$radioAlpha.Checked = ($global:userSettings.sort -eq "Alpha")
$radioAlpha.Add_CheckedChanged({
if ($radioAlpha.Checked) {
$global:userSettings.sort = "Alpha"
Save-UserSettings
Render-Tabs
}
})
$radioMostUsed.Checked = ($global:userSettings.sort -eq "Most Used")
$radioMostUsed.Add_CheckedChanged({
if ($radioMostUsed.Checked) {
$global:userSettings.sort = "Most Used"
Save-UserSettings
Render-Tabs
}
})

# Add the sort buttons to the settings page
$settingsPanel.Controls.AddRange( @( $radioAlpha, $radioMostUsed ) )
$settingsYPos += 30

# Create the label prompting the user to enter a font size
$fontLabel = New-Label -Text "Select the font size: " -X 20 -Y $settingsYPos
$settingsPanel.Controls.Add($fontLabel)
$settingsYPos += 20

$fontSButton = New-Button -Text "S" -X 20 -Y $settingsYPos
$fontSButton.Size = New-Object System.Drawing.Size(30, 20)
$fontSButton.Add_Click({
$global:userSettings.fontSize = 7
Save-UserSettings
Render-Tabs
})


$fontMButton = New-Button -Text "M" -X 50 -Y $settingsYPos
$fontMButton.Size = New-Object System.Drawing.Size(30, 20)
$fontMButton.Add_Click({
$global:userSettings.fontSize = 8
Save-UserSettings
Render-Tabs
})

$fontLButton = New-Button -Text "L" -X 80 -Y $settingsYPos
$fontLButton.Size = New-Object System.Drawing.Size(30, 20)
$fontLButton.Add_Click({
$global:userSettings.fontSize = 9
Save-UserSettings
Render-Tabs
})

$fontXLButton = New-Button -Text "XL" -X 110 -Y $settingsYPos
$fontXLButton.Size = New-Object System.Drawing.Size(30, 20)
$fontXLButton.Add_Click({
$global:userSettings.fontSize = 10
Save-UserSettings
Render-Tabs
})
$settingsPanel.Controls.AddRange( @($fontSButton, $fontMButton, $fontLButton, $fontXLButton ) )
$settingsYPos += 40

# Create check box for auto paste option
$hoverTabsToggle = New-Object System.Windows.Forms.CheckBox
$hoverTabsToggle.Text = "Open Tabs on Hover"
$hoverTabsToggle.Location = New-Object System.Drawing.Point(20, $settingsYPos)
$hoverTabsToggle.Size = New-Object System.Drawing.Size(150, 20)
$hoverTabsToggle.Checked = $global:userSettings.hoverTabs
$hoverTabsToggle.Add_CheckedChanged({
if ($hoverTabsToggle.Checked){
$global:userSettings.hoverTabs = $true

} elseif (-not $hoverTabsToggle.Checked) {

$global:userSettings.hoverTabs = $false
}

Save-UserSettings
}) # Shows/hides the delay input field depending on toggle state
$settingsPanel.Controls.Add($hoverTabsToggle)

$settingsYPos += 5
$hoverTabsLabel = New-Label -Text "<- Requires Restart!" -X 169 -Y $settingsYPos
$settingsPanel.Controls.Add($hoverTabsLabel)
$settingsYPos += 30

# Create Delay label
$delayLabel = New-Label -Text "Input Auto Paste Delay Time: " -X 20 -Y 190
$delayLabel.Tag = 'dynamic' # The delay label is tagged 'dynamic' because it appears and disappears on button click
$settingsPanel.Controls.Add($delayLabel)

# Create delay input text box
$delayInput = New-Input -X 20 -Y 210
$delayInput.Width = 30
$delayInput.Tag = 'dynamic'
$delayInput.Text = $global:userSettings.delayTime.ToString()
$delayInput.Add_TextChanged({
# Validate the user input
$parsedDelay = 0.0
    if ( [double]::TryParse($delayInput.Text, [ref]$parsedDelay) -and $parsedDelay -gt 0.0 ) {
# Update the delay and save user settings
        $global:userSettings.delayTime = $parsedDelay
Save-UserSettings
    }

})

# Create check box for auto paste option
$autoPasteToggle = New-Object System.Windows.Forms.CheckBox
$autoPasteToggle.Text = "Enable Auto-Paste"
$autoPasteToggle.Location = New-Object System.Drawing.Point(20, $settingsYPos)
$autoPasteToggle.Size = New-Object System.Drawing.Size(150, 20)
$autoPasteToggle.Checked = $global:userSettings.autoPaste
$autoPasteToggle.Add_CheckedChanged({ Update-DelayField }) # Shows/hides the delay input field depending on toggle state
Update-DelayField
$settingsPanel.Controls.Add($autoPasteToggle)

$settingsForm.Controls.Add($settingsPanel)
# Show settings window to the user
$settingsForm.ShowDialog()
}

function Render-SystemButtons {

# Button to open the edit window
$editButton = New-Button -Text "Edit" -X 280 -Y 104
$editButton.Tag = 'dynamic'
$editButton.Size = New-Object System.Drawing.Size(75, 30)
$editButton.Anchor = 'Bottom,Right'
$editButton.Add_Click({ Render-TabSelection })

# Button to open settings window
$settingsButton = New-Button -Text "Settings" -X 280 -Y 134
$settingsButton.Tag = 'dynamic'
$settingsButton.Anchor = 'Bottom,Right'
$settingsButton.Size = New-Object System.Drawing.Size(75, 30)
$settingsButton.Add_Click({ Render-SettingsWindow })

# Button to show users a tutorial message
$htuButton = New-Button -Text "How To Use" -X 280 -Y 164
$htuButton.Tag = 'dynamic'
$htuButton.Size = New-Object System.Drawing.Size(75, 30)
$htuButton.Anchor = 'Bottom,Right'
$htuButton.Add_Click({

[System.Windows.Forms.MessageBox]::Show(`
"How to use: `nClick the button, and it will automatically copy and paste for you! `nJust remember to click where you want it to be pasted. `nIf you don't click on time, the text will still be in your clipboard.", "How to Use"`
)

})

$form.Controls.AddRange(@($editButton, $settingsButton, $htuButton))

}

###################################################### Button Editor ##############################################################

# Function: Saves user edits
# Parameters: $NewTabDict is the dictionary that will be saved to the JSON, $Order is the order in which they will be saved in
# $Order is an optional parameter, because the order will not always be changed
# Output: User edits are saved to the workNotesModified.json file
function Save-UserWorknotes {

param( [hashtable]$NewTabDict, [array]$Order )

# Gets the current tab dictionary, before any edits are appended
$currentTabDict = Get-TabDict

# Creates a new empty tab dictionary to store the correct order
$orderedTabDict = [ordered]@{}

# If an explicit order is passed, append tabs from $NewTabDict to $OrderedTabDict in that order
if ($Order) {

# Iterates through the order that is passed
foreach ($item in $Order){

#If the dictionary contains the item listed in the order
if($NewTabDict.ContainsKey($item)){

# Adds the item to the orderedTabDict, ensuring every item is correctly stored in the desired order
$orderedTabDict[$item] = $NewTabDict[$item]
}
}

} else {
 
#If no order is passed, use $currentTabDict as a guide for already existing tabs
foreach ($key in $currentTabDict.Keys){

if($NewTabDict.ContainsKey($key)){
$orderedTabDict[$key] = $NewTabDict[$key]
}
}

# Adds each item that is not previously present
foreach ($key in $NewTabDict.Keys){
$orderedTabDict[$key] = $NewTabDict[$key]
}
}

#defines the file path
$workNotesModPath = "$PSScriptRoot\workNotesModified.json"

#writes the file
$orderedTabDict | ConvertTo-Json -Depth 10 | Set-Content -Path $workNotesModPath -Encoding UTF8
}

# Function: Provides a GUI that prompts the user to decide whether they will reset their edits
# Parameters: None
# Output: If user chooses to reset, overwrites with workNotesModified.json with workNotesDefault.json
function Reset-UserEdits {
 
# Creates a new form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Reset All Edits "
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = "CenterParent"
$form.FormBorderStyle = "FixedDialog"
$form.TopMost = $true

# Initializes the Y position
$yPos = 20

# Create the label
$nameLabel = New-Label -Text "Are you sure you want to Reset ALL Edits?`nThis CANNOT be undone. " -X 20 -Y $yPos
$form.Controls.Add($nameLabel)
$yPos += 35

# Create Yes Button
$yesButton = New-Button -Text "Yes" -X 200 -Y $yPos
$yesButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$yesButton.BackColor = [System.Drawing.Color]::PaleGreen
$form.AcceptButton = $yesButton
$form.Controls.Add($yesButton)

# Create No Button
$noButton = New-Button -Text "No" -X 20 -Y $yPos
$noButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$noButton.BackColor = [System.Drawing.Color]::PaleVioletRed
$form.CancelButton = $noButton
$form.Controls.Add($noButton)

# Assign Dialog to a variable to reference
$result = $form.ShowDialog()
# Only takes action if user presses OK
if ($result -eq [System.Windows.Forms.DialogResult]::OK){

# Define the path for the default worknotes
$workNotesDefaultPath = "$PSScriptRoot\workNotesDefault.json"

# Get the contents of the default worknotes file
$json = Get-Content $workNotesDefaultPath -Raw | ConvertFrom-Json

# initialize the new dict
$overwriteDict = [ordered]@{}

# for each item in the default notes, add them to the overwrite dictionary
foreach ($item in $json.PSObject.Properties) {
$overwriteDict[$item.Name] = $item.Value
}

# Define the path for the modified work notes
$workNotesModPath = "$PSScriptRoot\workNotesModified.json"

# Overwrite the modified work notes with the overwrite dictionary, wich now has the default configuration
$overwriteDict | ConvertTo-Json -Depth 10 | Set-Content -Path $workNotesModPath -Encoding UTF8

# Reset Tabs
Render-Tabs
# Close the form
$tsForm.Close()
}

}

# Function: Provides a GUI that prompts the user to input a string, either to add or edit a work note
# Parameters: OPTIONAL - if editing a string, $StringName is the String that is being edited
# Output: Returns user input as a string, as long as they input anything
function AddEdit-String {

param( [string]$StringName )

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Size = New-Object System.Drawing.Size(300,200)
$form.StartPosition = "CenterParent"
$form.FormBorderStyle = "FixedDialog"
$form.TopMost = $true

# Initialize Y position
$yPos = 20

# Create a label
$nameLabel = New-Label -Text "Input the Name: " -X 20 -Y $yPos
$form.Controls.Add($nameLabel)
$yPos += 20

# Create delay input text box
$nameInput = New-Input -X 20 -Y $yPos
$nameInput.Width = 200
$form.Controls.Add($nameInput)
$yPos += 30

# Checks if $StringName exists
if ($StringName){
# Correctly names the window
$form.Text = "$StringName Edit"

# Populates the text box with the name to be edited
$nameInput.Text = $StringName
} else {
# Correctly Names the window
$form.Text = "Add Page"
}

# Creates the OK button
$okButton = New-Button -Text "Ok" -X 20 -Y $yPos
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)


# Selects the input box for the user
$nameInput.Select()

# Assign Dialog to a variable to reference
$result = $form.ShowDialog()

# Validates the texxt input and checks if user presses OK
if ($nameInput.Text.Trim() -ne "" -and $result -eq [System.Windows.Forms.DialogResult]::OK){
# Returns the text in the input box
return [string]$nameInput.Text
}
# if the previous return is never reached, returns null
return $null

}

# Function: Provides a GUI that prompts the user to decide whether to delete the selected string
# Parameters: $StringName is the String that is being deleted
# Output: Returns a boolean as to whether the user has decided to delete the string
function Delete-String {

param( [string]$StringName )

$form = New-Object System.Windows.Forms.Form
$form.Text = "Delete $StringName"
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = "CenterParent"
$form.FormBorderStyle = "FixedDialog"
$form.TopMost = $true

$yPos = 20

$nameLabel = New-Label -Text "Are you sure you want to delete $StringName ?" -X 20 -Y $yPos
$form.Controls.Add($nameLabel)
$yPos += 35

$yesButton = New-Button -Text "Yes" -X 200 -Y $yPos
$yesButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$yesButton.BackColor = [System.Drawing.Color]::PaleGreen
$form.AcceptButton = $yesButton
$form.Controls.Add($yesButton)

$noButton = New-Button -Text "No" -X 20 -Y $yPos
$noButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$noButton.BackColor = [System.Drawing.Color]::PaleVioletRed
$form.CancelButton = $noButton
$form.Controls.Add($noButton)

$result = $form.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK){
return $true
} elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
return $false
}

}

# Function: Updates the noteList listBox depending on the radio button checked
# Parameters: None, but does use variables from Render-TabEdit
# Output: Clears and replaces the noteList listBox
function Update-NoteList {

$noteList.Items.Clear()

if($radioA.Checked){

$noteList.Items.AddRange(@($editApproveNotes.ToArray()))

} elseif ($radioD.Checked) {

$noteList.Items.AddRange(@($editDenyNotes.ToArray()))
}
}

# Function: Provides a GUI that prompts the user to take action on tab elements
# Parameters: [string]$Tab is the tab that the user has selected to work on
# Output: Presents the GUI and handles some decision logic
function Render-TabEdit {

param( [string]$Tab )

# Creates the form
$teForm = New-Object System.Windows.Forms.Form
$teForm.Text = "$Tab Edit"
$teForm.Size = New-Object System.Drawing.Size(535,470)
$teForm.StartPosition = "CenterParent"
$teForm.FormBorderStyle = "FixedDialog"
$teForm.TopMost = $true

# Creates a scrollable panel
    $tePanel = New-Object System.Windows.Forms.Panel
    $tePanel.Dock = 'Fill'
    $tePanel.AutoScroll = $true

# Instantiates $tabsDictionary as an ordered dictionary
$tabsDictionary = [ordered]@{}

# Defines the contents of $tabsDictionary with Get-TabDict
$tabsDictionary = Get-TabDict

# if $Tab does not exist, throw an error message and exit out of the function
if(-not $tabsDictionary[$Tab]){
[System.Windows.Forms.MessageBox]::Show("Please apply changes before editing. ")
return
}

# define arraylists for approve and deny notes
$editApproveNotes = [System.Collections.ArrayList]::new()
$editDenyNotes = [System.Collections.ArrayList]::new()

# populate the approve and deny notes with the selected tab's contents
if ($tabsDictionary[$Tab].approvedNotes) {
$editApproveNotes.AddRange($tabsDictionary[$Tab].approvedNotes)
}

if ($tabsDictionary[$Tab].denyNotes) {
$editDenyNotes.AddRange($tabsDictionary[$Tab].denyNotes)
}

# instantiates Y Position
$yPos = 20

# Creates a Tab Name label
$nameLabel = New-Label -Text "Tab Name" -X 20 -Y $yPos
$tePanel.Controls.Add($nameLabel)
$yPos += 20

# Creates a tab name input, and auto populate it with the selected tab's name
$nameInput = New-Input -X 20 -Y $yPos
$nameInput.Width = 100
$nameInput.Text = $Tab
$tePanel.Controls.Add($nameInput)
$yPos += 30

# Label prompting user to choose a filter
$filterLabel = New-Label -Text "Select a Work Note Type: " -X 20 -Y $yPos
$tePanel.Controls.Add($filterLabel)

# Label prompting user to choose a work note
$noteLabel = New-Label -Text "Select the Work Note you wish to edit: " -X 20 -Y 115
$tePanel.Controls.Add($noteLabel)

# Create a listbox that auto populates with the approved notes
$noteList = New-Object System.Windows.Forms.ListBox
$noteList.Location = New-Object System.Drawing.Point(20, 135)
$noteList.Size = New-Object System.Drawing.Size(260,198)
$noteList.Items.AddRange(@($editApproveNotes.toArray()))
$tePanel.Controls.Add($noteList)
# Creates approved filter radio button
$radioA = New-Object System.Windows.Forms.RadioButton
$radioA.Text = "Approved Notes"
$radioA.Location = New-Object System.Drawing.Point(20, 85)
$radioA.AutoSize = $true
$radioA.Checked = $true # Automatically checks on form open, to be consistent with listbox
# Updates listbox with approvedNotes when clicked
$radioA.Add_Click({
    $noteList.Items.Clear()
    $noteList.Items.AddRange(@($editApproveNotes.toArray()))
})
$tePanel.Controls.Add($radioA)

# Creates deny filter radio button
# Create RadioButton2
$radioD = New-Object System.Windows.Forms.RadioButton
$radioD.Text = "Deny Notes"
$radioD.Location = New-Object System.Drawing.Point(175, 85)
$radioD.AutoSize = $true
# Updates listbox with denyNotes when clicked
$radioD.Add_Click({
    $noteList.Items.Clear()
    $noteList.Items.AddRange(@($editDenyNotes.toArray()))
})
$tePanel.Controls.Add($radioD)

# Reset Y Pos
$yPos = 135

# Creates the Add button
$addButton = New-Button -Text "Add" -X 350 -Y $yPos
$addButton.BackColor = [System.Drawing.Color]::PaleGreen
$addButton.Add_Click({

# Calls AddEdit-String with no parameters to get Add GUI
$addNote = AddEdit-String

if ($addNote){

# Casts string to the result to avoid compatibility errors
$addNote = [string]$addNote

# Updates the corresponding arraylist with the note
if ( $radioA.Checked ) {

$editApproveNotes.Add($addNote)

} elseif ( $radioD.Checked ) {

$editDenyNotes.Add($addNote)
}

Update-NoteList
}

})

$yPos += 65

# Creates the edit button
$editButton = New-Button -Text "Edit" -X 350 -Y $yPos
$editButton.BackColor = [System.Drawing.Color]::LightYellow
$editButton.Add_Click({
# Checks the user selected note from the noteList
$selectedNote = $noteList.SelectedItem

# only takes action if there is a note selected
    if ($selectedNote) {

# Calls AddEdit-String with a parameter to get EDIT gui
        $editedNote = AddEdit-String -String $selectedNote

# Only takes action if editedNote exists
if ($editedNote) {

# casts string on result to avoid compatibility errors
$editedNote = [string]$editedNote
# Updates appropriate arraylist
if ( $radioA.Checked ) {

# Index of the selected Note
$index = $editApproveNotes.IndexOf($selectedNote)
# if the index exists, replaces it with the edited note
if ($index -ne -1) { $editApproveNotes[$index] = $editedNote }

} elseif ( $radioD.Checked ) {
# Index of the selected Note
$index = $editDenyNotes.IndexOf($selectedNote)
# if the index exists, replaces it with the edited note
if ($index -ne -1) {$editDenyNotes[$index] = $editedNote }
}

Update-NoteList
# Selects the editedNote in the noteList
$noteList.SelectedIndex = $index
}

# throws an error if no note is selected
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a note first.")
    }
})
$yPos += 65

# Creates the delete button
$deleteButton = New-Button -Text "Delete" -X 350 -Y $yPos
$deleteButton.BackColor = [System.Drawing.Color]::PaleVioletRed
$deleteButton.Add_Click({

# Checks the user selected note from the noteList
$selectedNote = $noteList.SelectedItem

# only takes action if there is a note selected
if ($selectedNote) {
# Calls Delete string to get user's decision as a boolean
$willDelete = Delete-String -String $selectedNote

# takes action only if they pressed OK
if ($willDelete){

# updates appropriate arraylists
if ( $radioA.Checked ) {

$editApproveNotes.Remove($selectedNote)

} elseif ( $radioD.Checked ) {

$editDenyNotes.Remove($selectedNote)
}

Update-NoteList
}

# throws an error if no arraylist is selected
} else {
        [System.Windows.Forms.MessageBox]::Show("Please select a note first.")
    }
})
$yPos = 135

# Creates the up button
$upButton = New-Button -Text "Move Up" -X 290 -Y $yPos
$upButton.BackColor = [System.Drawing.Color]::LightBlue
$upButton.Size = New-Object System.Drawing.Size(50, 95)
$upButton.Add_Click({

# gets the selected item and its index
$selectedNote = $noteList.SelectedItem
$selectedIndex = $noteList.SelectedIndex

# takes action if the note exists and the index is greater than 0
# no action is taken on the first item in the list
if($selectedNote -and $selectedIndex -gt 0){
# swaps the current selection and the note above it in the appropriate arraylist
if ( $radioA.Checked ) {
$temp = $editApproveNotes[$selectedIndex]
$editApproveNotes[$selectedIndex] = $editApproveNotes[$selectedIndex -1]
$editApproveNotes[$selectedIndex - 1] = $temp 

} elseif ( $radioD.Checked ) {

$temp = $editDenyNotes[$selectedIndex]
$editDenyNotes[$selectedIndex] = $editDenyNotes[$selectedIndex -1]
$editDenyNotes[$selectedIndex - 1] = $temp
}

Update-NoteList

# selects the new position, which now contains the correct work note
$noteList.SelectedIndex = $selectedIndex - 1

}
})
$yPos += 95

# Creates the down button
$downButton = New-Button -Text "Move Down" -X 290 -Y $yPos
$downButton.BackColor = [System.Drawing.Color]::LightBlue
$downButton.Size = New-Object System.Drawing.Size(50, 96)
$downButton.Add_Click({

# gets the selected item and its index
$selectedTab = $noteList.SelectedItem
$selectedIndex = $noteList.SelectedIndex

# takes action if the note exists and the index is less than the length of the list - 1
# no action is taken on the last item in the list
if($selectedTab -and $selectedIndex -lt ($noteList.Items.Count - 1)){
# swaps the current selection adn the note below it in the appropriate arrayList
if ( $radioA.Checked ) {
$temp = $editApproveNotes[$selectedIndex]
$editApproveNotes[$selectedIndex] = $editApproveNotes[$selectedIndex + 1]
$editApproveNotes[$selectedIndex + 1] = $temp 

} elseif ( $radioD.Checked ) {

$temp = $editDenyNotes[$selectedIndex]
$editDenyNotes[$selectedIndex] = $editDenyNotes[$selectedIndex + 1]
$editDenyNotes[$selectedIndex + 1] = $temp
}

Update-NoteList
# selects the new position, which now contains the correct work note
$noteList.SelectedIndex = $selectedIndex + 1
}
})
$yPos += 65

# applies all changes to the JSON file
$applyButton = New-Button -Text "Apply" -X 20 -Y 350
$applyButton.Size = New-Object System.Drawing.Size(482, 50)
$applyButton.Add_Click({

# instantiates a new array which will contain explicitly casted strings
$cleanApprovedNotes = @()

# iterates through the editApproveNotes and ensures that only strings are added to cleanApprovedNotes
foreach ($item in $editApproveNotes) {
if ($item -is [string]){
$cleanApprovedNotes += [string]$item
} elseif ($item -ne $null) {
$cleanApprovedNotes += [string]$item
}
}

# instantiates a new array which will contain explicitly casted strings
$cleanDenyNotes = @()

# iterates through the editDenyNotes and ensures that only strings are added to cleanDenyNotes
foreach ($item in $editDenyNotes) {
if ($item -is [string]){
$cleanDenyNotes += [string]$item
} elseif ($item -ne $null) {
$cleanDenyNotes += [string]$item
}
}
# updates tabsDictionary at the selected tab's approved and deny notes with the clean notes
$tabsDictionary[$Tab].approvedNotes = $cleanApprovedNotes
$tabsDictionary[$Tab].denyNotes = $cleanDenyNotes

# updates the tab name if it exists and does not equal the previous tab name
if ($nameInput.Text -and $nameInput.Text -ne $Tab){
# creates a new tab with the new name and the contents of the old tab
$tabsDictionary[$nameInput.Text] = $tabsDictionary[$Tab]

# removes the old tab
$tabsDictionary.Remove($Tab)
}

# saves the edits to the JSON file
Save-UserWorkNotes -NewTabDict $tabsDictionary

# closes the windows
$teForm.Hide()
$teForm.Dispose()

$tsForm.Hide()
$tsForm.Dispose()

# Resets the GUI
Render-Tabs
Render-TabSelection
})

# Adds all the buttons
$tePanel.Controls.AddRange( @(
$addButton,`
$editButton,`
$deleteButton,`
$upButton,`
$downButton,`
$applyButton`
) )

# shows user the form
$teForm.Controls.Add($tePanel)
$teForm.ShowDialog()

}

# Function: Provides a GUI that prompts the user to take action on tabs
# Parameters: None
# Output: Presents the GUI and handles some decision logic
function Render-TabSelection {

# Create the tab selection Page
$tsForm = New-Object System.Windows.Forms.Form
$tsForm.Text = "Tab Editor"
$tsForm.Size = New-Object System.Drawing.Size(535,400)
$tsForm.StartPosition = "CenterParent"
$tsForm.FormBorderStyle = "FixedDialog"
$tsForm.TopMost = $true

    # Create a scrollable panel
    $tsPanel = New-Object System.Windows.Forms.Panel
    $tsPanel.Dock = 'Fill'
    $tsPanel.AutoScroll = $true

# instantiate tabDict as an ordered dictionary
$tabDict = [ordered]@{}

# define tabDict with Get-TabDict
$tabDict = Get-TabDict

# instantiate Y position
    $yPos = 20

# creates a label prompting user to select a tab
$typeLabel = New-Label -Text "Select the Tab you wish to edit" -X 20 -Y $yPos
$tsPanel.Controls.Add($typeLabel)
$yPos += 20
# creates a listbox populated by tabs
$tabList = New-Object System.Windows.Forms.ListBox
$tabList.Location = New-Object System.Drawing.Point(20, $yPos)
$tabList.Size = New-Object System.Drawing.Size(260,200)
# populates the tablist with the contents of tabDict
foreach ($tabName in $tabDict.Keys){ $tabList.Items.Add($tabName) }
$tsPanel.Controls.Add($tabList)

# creates add button
$addButton = New-Button -Text "Add" -X 350 -Y $yPos
$addButton.BackColor = [System.Drawing.Color]::PaleGreen
$addButton.Add_Click({

# calls addedit string with no parameters to evoke add gui
$addTab = AddEdit-String
if ($addTab){

# casts string to addTab to avoid compatibility issues
$addTab = [string]$addTab

# creates a default state for the new tab
$tabDict[$addTab] = @{

Color = 'Control'
approvedNotes = @( )
denyNotes = @( )
}

# updates the tablist with new changes
$tabList.Items.Clear()
foreach ($tabName in $tabDict.Keys){ $tabList.Items.Add($tabName) }
} 
})
$yPos += 65

# creates the edit button
$editButton = New-Button -Text "Edit" -X 350 -Y $yPos
$editButton.BackColor = [System.Drawing.Color]::LightYellow
$editButton.Add_Click({

# gets the selected tab
$selectedTab = $tabList.SelectedItem

# only acts if a tab is selected
    if ($selectedTab) {
# calls tab edit with selected tab
        Render-TabEdit -Tab $selectedTab
# throws an error if no tab is selected
    } else {
        [System.Windows.Forms.MessageBox]::Show("Please select a tab first.")
    }
})
$yPos += 65

# creates the delete button
$deleteButton = New-Button -Text "Delete" -X 350 -Y $yPos
$deleteButton.BackColor = [System.Drawing.Color]::PaleVioletRed
$deleteButton.Add_Click({

# gets the selected tab
$selectedTab = $tabList.SelectedItem

# only acts if a tab is selected
if ($selectedTab) {
# gets user decision from Delete-String return statement
$willDelete = Delete-String -String $selectedTab

# Acts on user decision
if ($willDelete){
$tabDict.Remove($selectedTab)
}

# Updates the tablist
$tabList.Items.Clear()
foreach ($tabName in $tabDict.Keys){ $tabList.Items.Add($tabName) }

# throws an error if no tab is selected
} else {
[System.Windows.Forms.MessageBox]::Show("Please select a tab first.")
}
})

# instantiates the y position
$yPos = 40

# creates the up button
$upButton = New-Button -Text "Move Up" -X 290 -Y $yPos
$upButton.BackColor = [System.Drawing.Color]::LightBlue
$upButton.Size = New-Object System.Drawing.Size(50, 95)
$upButton.Add_Click({

# gets the selected tab and its index
$selectedTab = $tabList.SelectedItem
$selectedIndex = $tabList.SelectedIndex

# acts only if tab exists and if it is not the first item in the list
if($selectedTab -and $selectedIndex -gt 0){
# Get Keys List
$keysList = [System.Collections.ArrayList]@($tabDict.Keys)

#Swap Positions
$temp = $keysList[$selectedIndex]
$keysList[$selectedIndex] = $keysList[$selectedIndex - 1]
$keysList[$selectedIndex - 1] = $temp

# instantiate ordered dict as an ordered dictionary
$newOrderedDict = [ordered]@{}

# append new order to orderedDict
foreach ($key in $keysList){
$newOrderedDict[$key] = $tabDict[$key]
}

# update the tabDict directly
$tabDict.Clear()
foreach ($kvp in $newOrderedDict.GetEnumerator()){
$tabDict[$kvp.Key] = $kvp.Value
}

# update the tablist
$tabList.Items.Clear()
foreach ($tabName in $tabDict.Keys){ $tabList.Items.Add($tabName) }

# select the new position
$tabList.SelectedIndex = $selectedIndex - 1

}
})
$yPos += 95

# creates the down button
$downButton = New-Button -Text "Move Down" -X 290 -Y $yPos
$downButton.BackColor = [System.Drawing.Color]::LightBlue
$downButton.Size = New-Object System.Drawing.Size(50, 96)
$downButton.Add_Click({

# gets the selected tab and its index
$selectedTab = $tabList.SelectedItem
$selectedIndex = $tabList.SelectedIndex

# only acts if a tab is selected and it is not the last item in the list
if($selectedTab -and $selectedIndex -lt ($tabList.Items.Count - 1)){
# Get Keys List
$keysList = @($tabDict.Keys)

#Swap Positions
$temp = $keysList[$selectedIndex]
$keysList[$selectedIndex] = $keysList[$selectedIndex + 1]
$keysList[$selectedIndex + 1] = $temp

# instantiate ordered dict as an ordered dictionary
$newOrderedDict = [ordered]@{}

# append new order to orderedDict
foreach ($key in $keysList){
$newOrderedDict[$key] = $tabDict[$key]
}

# update the tabDict directly
$tabDict.Clear()
foreach ($kvp in $newOrderedDict.GetEnumerator()){
$tabDict[$kvp.Key] = $kvp.Value
}

# update the tablist
$tabList.Items.Clear()
foreach ($tabName in $tabDict.Keys){ $tabList.Items.Add($tabName) }

# select the new position
$tabList.SelectedIndex = $selectedIndex + 1

}
})
$yPos += 65

# creates the apply button
$applyButton = New-Button -Text "Apply" -X 20 -Y 250
$applyButton.Size = New-Object System.Drawing.Size(482, 50)
$applyButton.Add_Click({

# creates an empty array to hold the new order
$newOrder = @()

# appends tabList order to the order array
foreach ($item in $tabList.Items) { $newOrder += $item }

# saves edits
Save-UserWorkNotes -NewTabDict $tabDict -Order $newOrder

# resets tabs
Render-Tabs
})

# creates the reset button
$resetButton = New-Button -Text "RESET" -X 20 -Y 300
$resetButton.Size = New-Object System.Drawing.Size(482, 50)
$resetButton.BackColor = [System.Drawing.Color]::Black
$resetButton.ForeColor = [System.Drawing.Color]::White
$resetButton.Add_Click({ Reset-UserEdits })

# add buttons to the panel
$tsPanel.Controls.AddRange( @(
$addButton,`
  $editButton, `
$deleteButton, `
$upButton, `
$downButton, `
$applyButton, `
$resetButton `
) )
# add panel to the form
$tsForm.Controls.Add($tsPanel)

# present to user
$tsForm.ShowDialog()

}


##################################################### MAIN FUNCTION #############################################################

function Main {

# Imports the necessary hooks into Windows to enable copying/pasting
Add-Type -AssemblyName System.Windows.Forms

# Creates the window for the GUI
$form = New-Object System.Windows.Forms.Form
$form.Text = "Work Notes Clipboard"
$form.Size = New-Object System.Drawing.Size(400, 300)
$form.MinimumSize = New-Object System.Drawing.Size(300, 200)

# Creates the tabs section of the window
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Dock = 'Fill'

if ($global:userSettings.hoverTabs){

# Simulate hover by tracking mouse movement over the tab headers
$tabControl.Add_MouseMove({
    $mousePos = $tabControl.PointToClient([System.Windows.Forms.Cursor]::Position)
    for ($i = 0; $i -lt $tabControl.TabPages.Count; $i++) {
        $tabRect = $tabControl.GetTabRect($i)
        if ($tabRect.Contains($mousePos)) {
            $tabControl.SelectedIndex = $i
            break
        }
    }
})
}

# Create tabs and assign them to the form
Render-Tabs

Render-SystemButtons

$brosioLabel = New-Label -text "Program Written by Ambrosio Gomez IV, 2025" -X 120 -Y 220
$brosioLabel.Anchor = 'Bottom,Right'
$form.Controls.Add($brosioLabel)

# Add controls to the form
$form.Controls.Add($tabControl)

# Show the form to the user
[void]$form.ShowDialog()
}

Main