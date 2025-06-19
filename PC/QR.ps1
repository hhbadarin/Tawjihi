# Define the link to convert to a QR code
$link = "https://www.example.com"

# Define customization parameters
$size = 500  # Size of the QR code in pixels (width and height)
$foregroundColor = "000000"  # Foreground color (black)
$backgroundColor = "FFFFFF"  # Background color (white)
$errorCorrection = "H"  # Error correction level (L, M, Q, H)
$margin = 4  # Margin around the QR code

# Define the API URL with customization parameters
$apiUrl = "https://qrcode.tec-it.com/API/QRCode?data=$link&size=$size&color=$foregroundColor-$backgroundColor&ecc=$errorCorrection&margin=$margin"

# Define the output path
$outputPath = "$Home\Desktop\qrcode.png"  # Save to the desktop

# Download the QR code image
Invoke-WebRequest -Uri $apiUrl -OutFile $outputPath

Write-Output "QR code saved to $outputPath"