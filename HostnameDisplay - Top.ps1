# Enable DPI awareness so coordinates match physical pixels regardless of Windows scaling
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class DPI {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
}
"@
[DPI]::SetProcessDPIAware() | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Add Win32 API functions for click-through
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win32 {
    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_TRANSPARENT = 0x20;
    public const int WS_EX_LAYERED = 0x80000;
    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);
    [DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
}
"@

# Add Win32 API for fullscreen detection
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win32Extra {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

# Add Win32 API for getting window class name
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public class Win32Class {
    [DllImport("user32.dll", SetLastError=true)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
}
"@

# Add ForegroundWin type ONCE here
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class ForegroundWin {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", SetLastError=true)]
    public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
}
"@

$script:formHandle = $null

function Get-WindowClassName {
    param([IntPtr]$hWnd)
    $sb = New-Object System.Text.StringBuilder 256
    [Win32Class]::GetClassName($hWnd, $sb, $sb.Capacity) | Out-Null
    $sb.ToString()
}

function Test-FullscreenWindow {
    param($form)
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $margin = 2 # Allow a small margin for borderless windows

    # Get the foreground window
    $fgHandle = [ForegroundWin]::GetForegroundWindow()
    if ($fgHandle -eq [IntPtr]::Zero) { return $false }
    if ($fgHandle.ToInt64() -eq $script:formHandle.ToInt64()) { return $false }

    # Ignore desktop, taskbar, and empty-title windows
    $className = Get-WindowClassName $fgHandle
    $sb = New-Object System.Text.StringBuilder 256
    [ForegroundWin]::GetWindowText($fgHandle, $sb, $sb.Capacity) | Out-Null
    $windowTitle = $sb.ToString()

    if (
        $className -eq "Progman" -or
        $className -eq "WorkerW" -or
        $className -eq "Shell_TrayWnd" -or
        $windowTitle -eq ""
    ) { return $false }

    if (-not [Win32Extra]::IsWindowVisible($fgHandle)) { return $false }
    $rect = New-Object Win32Extra+RECT
    if (-not [Win32Extra]::GetWindowRect($fgHandle, [ref]$rect)) { return $false }

    foreach ($screen in $screens) {
        $screenLeft = $screen.Bounds.X
        $screenTop = $screen.Bounds.Y
        $screenRight = $screen.Bounds.X + $screen.Bounds.Width
        $screenBottom = $screen.Bounds.Y + $screen.Bounds.Height
        if (
            ($rect.Left -le ($screenLeft + $margin)) -and
            ($rect.Top -le ($screenTop + $margin)) -and
            ($rect.Right -ge ($screenRight - $margin)) -and
            ($rect.Bottom -ge ($screenBottom - $margin))
        ) {
            return $true
        }
    }
    return $false
}

$hostname = $env:COMPUTERNAME

# Start with "almost black" for white text
$form = New-Object System.Windows.Forms.Form
$form.Text = ""
$form.FormBorderStyle = 'None'
$form.TopMost = $true
$form.ShowInTaskbar = $true
$form.BackColor = [System.Drawing.Color]::FromArgb(128, 128, 128)
$form.TransparencyKey = [System.Drawing.Color]::FromArgb(128, 128, 128)
$form.Size = New-Object System.Drawing.Size(260, 80)
$form.StartPosition = 'Manual'

# Get primary screen working area
$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$screenBounds = $screen.WorkingArea

# Padding from screen edges (positive values)
$paddingRight = 20
$paddingTop = 10

# Move the form 10 pixels up and 100 pixels to the left from the original top-right position
$x = $screenBounds.Right - $form.Width - $paddingRight - 75
$y = $screenBounds.Top + $paddingTop - 5
$form.Location = New-Object System.Drawing.Point($x, $y)

# Variables for drawing
$script:textColor = [System.Drawing.Color]::White
$script:hostnameText = "Hostname: $hostname"
$script:font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$script:isBright = $false

# Indicator variables for detection
$script:sampleX = 0
$script:sampleY = 0
$script:sampleWidth = 3
$script:sampleHeight = 3
$script:detectScreenX = 0
$script:detectScreenY = 0

# Custom paint event for drawing text directly on the form
$form.Add_Paint({
    param($src, $e)
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAlias
    $size = $e.Graphics.MeasureString($script:hostnameText, $script:font)
    $x = [int](($form.ClientSize.Width - $size.Width) / 2)
    $y = 1 # Draw text 1 pixel from the top
    $brush = New-Object System.Drawing.SolidBrush($script:textColor)
    $e.Graphics.DrawString($script:hostnameText, $script:font, $brush, $x, $y)
    $brush.Dispose()
})

# Timer for updating luminance and text color
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 100 # milliseconds

$timer.Add_Tick({
    if (Test-FullscreenWindow $form) {
        if ($form.Opacity -ne 0) { $form.Opacity = 0 }
        return
    } else {
        if ($form.Opacity -ne 1) { $form.Opacity = 1 }
    }

    # Sample just below the text (since text is at the top)
    $sampleMargin = 10  # Margin below the text
    $size = [System.Windows.Forms.TextRenderer]::MeasureText($script:hostnameText, $script:font)
    $script:sampleX = [int](($form.ClientSize.Width - $script:sampleWidth) / 2)
    $script:sampleY = $size.Height + $sampleMargin

    # Convert to screen coordinates and store for paint event
    $screenPoint = $form.PointToScreen([System.Drawing.Point]::new($script:sampleX, $script:sampleY))
    $script:detectScreenX = $screenPoint.X
    $script:detectScreenY = $screenPoint.Y

    # Sample the screen at this location
    $bmp = New-Object System.Drawing.Bitmap($script:sampleWidth, $script:sampleHeight)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($screenPoint.X, $screenPoint.Y, 0, 0, $bmp.Size)

    $sumR = 0; $sumG = 0; $sumB = 0
    for ($ix = 0; $ix -lt $bmp.Width; $ix++) {
        for ($iy = 0; $iy -lt $bmp.Height; $iy++) {
            $pixel = $bmp.GetPixel($ix, $iy)
            $sumR += $pixel.R
            $sumG += $pixel.G
            $sumB += $pixel.B
        }
    }
    $totalPixels = $bmp.Width * $bmp.Height
    $avgR = [int]($sumR / $totalPixels)
    $avgG = [int]($sumG / $totalPixels)
    $avgB = [int]($sumB / $totalPixels)

    $luminance = (0.299 * $avgR) + (0.587 * $avgG) + (0.114 * $avgB)

    if (-not $script:isBright -and $luminance -ge 180) {
        $script:textColor = [System.Drawing.Color]::Black
        $form.BackColor = [System.Drawing.Color]::FromArgb(128, 128, 128)
        $form.TransparencyKey = [System.Drawing.Color]::FromArgb(128, 128, 128)
        $script:isBright = $true
        $form.Invalidate()
    } elseif ($script:isBright -and $luminance -lt 180) {
        $script:textColor = [System.Drawing.Color]::White
        $form.BackColor = [System.Drawing.Color]::FromArgb(128, 128, 128)
        $form.TransparencyKey = [System.Drawing.Color]::FromArgb(128, 128, 128)
        $script:isBright = $false
        $form.Invalidate()
    }

    $g.Dispose()
    $bmp.Dispose()
})

$form.Add_Shown({
    $script:formHandle = $form.Handle
    $hwnd = $form.Handle
    $exStyle = [Win32]::GetWindowLong($hwnd, [Win32]::GWL_EXSTYLE)
    [Win32]::SetWindowLong($hwnd, [Win32]::GWL_EXSTYLE, $exStyle -bor [Win32]::WS_EX_LAYERED -bor [Win32]::WS_EX_TRANSPARENT) | Out-Null
    $timer.Start()
})

[void]$form.ShowDialog()