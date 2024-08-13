[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$gasUrl = "https://script.google.com/macros/s/AKfycbwSK2urFo3oC8HBeaW_0o2rp-g8OLGWgYCEpuIS5eYKudwPH93fQWpHq5EHQS88rgLXnA/exec"

Add-Type -AssemblyName System.Windows.Forms

$form = New-Object System.Windows.Forms.Form
$form.Text = "GAS Request"
$form.Size = New-Object System.Drawing.Size(300,150)
$form.StartPosition = "CenterScreen"

$inputBox = New-Object System.Windows.Forms.TextBox
$inputBox.Size = New-Object System.Drawing.Size(260,20)
$inputBox.Location = New-Object System.Drawing.Point(10,10)
$form.Controls.Add($inputBox)

$sendButton = New-Object System.Windows.Forms.Button
$sendButton.Text = "Send"
$sendButton.Location = New-Object System.Drawing.Point(10,40)
$form.Controls.Add($sendButton)


$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Size = New-Object System.Drawing.Size(260,20)
$outputBox.Location = New-Object System.Drawing.Point(10,70)
$outputBox.ReadOnly = $true
$form.Controls.Add($outputBox)


$sendButton.Add_Click({
    $content = $inputBox.Text
    if ($content -ne "") {
        try {
            $response = Invoke-RestMethod -Uri $gasUrl -Method POST -Body $content -ContentType "application/json; charset=utf-8"
            if ($response -and $response -ne "ERROR") {
                $outputBox.Text = $response
            } else {
                $outputBox.Text = "Error: Invalid response from server."
            }
        } catch {
            $outputBox.Text = "Error: Failed to connect to server."
        }
    } else {
        $outputBox.Text = "Please enter text to send."
    }
})


$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
