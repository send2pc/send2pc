Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class MouseHelper {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern bool GetCursorPos(out POINT lpPoint);

    public struct POINT {
        public int X;
        public int Y;
    }

    public static void MoveCursorSlightly() {
        POINT point;
        GetCursorPos(out point);
        SetCursorPos(point.X + 1, point.Y);  // حرکت مکان‌نما به اندازه 1 پیکسل به سمت راست
    }
}
"@

function Test-Ping {
    param (
        [string[]]$Addresses
    )
    
    $random = New-Object System.Random
    $address = $Addresses[$random.Next($Addresses.Length)]
    
    try {
        $pingResult = Test-Connection -ComputerName $address -Count 1 -Quiet
        return $pingResult
    }
    catch {
        Write-Output "An error occurred during ping: $_"
        return $false
    }
}

$pingAddresses = @("1.1.1.1", "8.8.8.8", "9.9.9.9")  # اضافه کردن آدرس‌های دلخواه

while ($true) {
    if (Test-Ping -Addresses $pingAddresses) {
        [MouseHelper]::MoveCursorSlightly()
    }
    
        Start-Sleep -Milliseconds 500

}
