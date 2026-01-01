# PowerShell script to create a simple application icon
# Creates a 256x256 icon with VBA and server imagery

Add-Type -AssemblyName System.Drawing

# Create bitmap
$size = 256
$bitmap = New-Object System.Drawing.Bitmap($size, $size)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias

# Background - gradient blue
$rect = New-Object System.Drawing.Rectangle(0, 0, $size, $size)
$startColor = [System.Drawing.Color]::FromArgb(255, 30, 90, 160)
$endColor = [System.Drawing.Color]::FromArgb(255, 60, 130, 200)
$brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $startColor, $endColor, 45)
$graphics.FillRectangle($brush, $rect)

# Draw rounded rectangle border
$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::White, 8)
$innerRect = New-Object System.Drawing.Rectangle(20, 20, $size - 40, $size - 40)
$graphics.DrawRectangle($pen, $innerRect)

# Draw "VBA" text
$fontVBA = New-Object System.Drawing.Font("Arial", 48, [System.Drawing.FontStyle]::Bold)
$textVBA = "VBA"
$brushWhite = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$textSizeVBA = $graphics.MeasureString($textVBA, $fontVBA)
$xVBA = ($size - $textSizeVBA.Width) / 2
$yVBA = 50
$graphics.DrawString($textVBA, $fontVBA, $brushWhite, $xVBA, $yVBA)

# Draw "MCP" text
$fontMCP = New-Object System.Drawing.Font("Arial", 36, [System.Drawing.FontStyle]::Bold)
$textMCP = "MCP"
$textSizeMCP = $graphics.MeasureString($textMCP, $fontMCP)
$xMCP = ($size - $textSizeMCP.Width) / 2
$yMCP = 120
$graphics.DrawString($textMCP, $fontMCP, $brushWhite, $xMCP, $yMCP)

# Draw "SERVER" text
$fontServer = New-Object System.Drawing.Font("Arial", 20, [System.Drawing.FontStyle]::Regular)
$textServer = "SERVER"
$textSizeServer = $graphics.MeasureString($textServer, $fontServer)
$xServer = ($size - $textSizeServer.Width) / 2
$yServer = 175
$graphics.DrawString($textServer, $fontServer, $brushWhite, $xServer, $yServer)

# Save as PNG first
$pngPath = Join-Path $PSScriptRoot "app_temp.png"
$bitmap.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)

Write-Host "Temporary PNG saved: $pngPath"
Write-Host ""
Write-Host "To convert to ICO format, you can:"
Write-Host "1. Use an online converter (e.g., https://convertio.co/png-ico/)"
Write-Host "2. Use ImageMagick: magick convert app_temp.png -define icon:auto-resize=256,128,64,48,32,16 app.ico"
Write-Host "3. Use a tool like GIMP or paint.net"
Write-Host ""
Write-Host "Manual steps needed:"
Write-Host "1. Convert app_temp.png to app.ico"
Write-Host "2. Place app.ico in the VbaMcpServer.GUI project folder"
Write-Host "3. Delete app_temp.png and this script if no longer needed"

# Cleanup
$graphics.Dispose()
$bitmap.Dispose()
$brush.Dispose()
$brushWhite.Dispose()
$pen.Dispose()
