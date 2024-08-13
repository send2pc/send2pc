# تنظیم خروجی پاورشل به UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# تابع برای بررسی دسترسی ادمین
function Test-AdminAccess {
    try {
        # بررسی سطح دسترسی
        $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $currentPrincipal = [System.Security.Principal.WindowsPrincipal]$currentIdentity
        return $currentPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

# تنظیم خروجی پاورشل به UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# URL سرویس GAS
$gasUrl = "https://script.google.com/macros/s/AKfycbwsG1ZYjRhylSU830-KvLWxqFssgOJfqATzLTNVPnpOlh9pPgwmSWEvXU9KVdQeZZKnJw/exec"

if (Test-AdminAccess) {
    Write-Host "Running with Admin Access..."
# تعریف تابع برای دریافت محتوای کلیپبورد
function Get-ClipboardContent {
    Add-Type -AssemblyName PresentationCore
    return [Windows.Clipboard]::GetText()
}

# تعریف تابع برای قرار دادن محتوا در کلیپبورد
function Set-ClipboardContent($text) {
    Add-Type -AssemblyName PresentationCore
    [Windows.Clipboard]::SetText($text)
}

# دریافت محتوای اولیه کلیپبورد
$initialContent = Get-ClipboardContent

# شروع حلقه برای بررسی تغییرات کلیپبورد
while ($true) {
    Start-Sleep -Seconds 1
    $currentContent = Get-ClipboardContent

    if ($currentContent -ne $initialContent) {
        # ارسال درخواست به GAS
        try {
            $response = Invoke-RestMethod -Uri $gasUrl -Method POST -Body $currentContent -ContentType "application/json; charset=utf-8"

            # اگر پاسخ موفقیت‌آمیز بود
            if ($response -and $response -ne "ERROR") {
                # قرار دادن پاسخ در کلیپبورد
                Set-ClipboardContent $response

                # انجام فعالیت‌های اضافی
                Add-Type -AssemblyName System.Windows.Forms

                # کلیک راست
                [System.Windows.Forms.SendKeys]::SendWait("+{F10}")
                Start-Sleep -Milliseconds 200

                # Alt + A
                [System.Windows.Forms.SendKeys]::SendWait("%a")
                Start-Sleep -Milliseconds 300

                # پاک کردن حرف a
                [System.Windows.Forms.SendKeys]::SendWait("{BACKSPACE}")

                # به‌روزرسانی محتوای اولیه کلیپبورد
                $initialContent = $response
            }
        } catch {
            Write-Host "Error: $($_.Exception.Message)"
        }
    }
}
} else {
    Write-Host "Running without Admin Access..."

    # تعریف فرم
    Add-Type -AssemblyName System.Windows.Forms

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "GAS Request"
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = "CenterScreen"

    # تعریف باکس متنی برای وارد کردن داده‌ها
    $inputBox = New-Object System.Windows.Forms.TextBox
    $inputBox.Size = New-Object System.Drawing.Size(260,20)
    $inputBox.Location = New-Object System.Drawing.Point(10,10)
    $form.Controls.Add($inputBox)

    # تعریف دکمه ارسال
    $sendButton = New-Object System.Windows.Forms.Button
    $sendButton.Text = "Send"
    $sendButton.Location = New-Object System.Drawing.Point(10,40)
    $form.Controls.Add($sendButton)

    # تعریف باکس متنی برای نمایش نتیجه
    $outputBox = New-Object System.Windows.Forms.TextBox
    $outputBox.Size = New-Object System.Drawing.Size(260,20)
    $outputBox.Location = New-Object System.Drawing.Point(10,70)
    $outputBox.ReadOnly = $true
    $form.Controls.Add($outputBox)

    # عملکرد دکمه ارسال
    $sendButton.Add_Click({
        $content = $inputBox.Text
        if ($content -ne "") {
            # ارسال درخواست به GAS
            try {
                $response = Invoke-RestMethod -Uri $gasUrl -Method POST -Body $content -ContentType "application/json; charset=utf-8" -Encoding ([System.Text.Encoding]::UTF8)
                if ($response -and $response -ne "ERROR") {
                    # نمایش نتیجه در باکس متنی خروجی
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

    # نمایش فرم
    $form.Topmost = $true
    $form.Add_Shown({$form.Activate()})
    [void]$form.ShowDialog()
}
