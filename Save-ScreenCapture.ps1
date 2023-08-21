<#PSScriptInfo

.VERSION 1.0.0.2

.GUID f5783ffa-47a5-49ab-ad1b-d2624d662fd7

.AUTHOR Jeffrey Snover

.COMPANYNAME Microsoft

.COPYRIGHT 

.TAGS 

.LICENSEURI 

.PROJECTURI 

.ICONURI 

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS 

.EXTERNALSCRIPTDEPENDENCIES 

.RELEASENOTES
Version 1.0.0.1 added the -Quality parameter to specify the level of compression: 0 for min size, 100 for max quality
Version 1.0.0.2 Default value for -Quality parameter is 100.  You do something if you want to save memory.


#>

<# 

.DESCRIPTION 
 Capture and save one or more screens to a file as a .JPG.
 You provide a directory or it will use the current directory.
 The files will be named by combining the screen device (e.g. DISPLAY1) and a timestamp.
 Use -Verbose to see the resultant filename.
    Note - this is based upon a code fragment from a piece of malware discovered by John Lambert @JohnLATwC
    It looked pretty useful so I decided to clean it up and make it available to the rest of the world. 
.SYNOPSIS
 Capture and save one or more screens to a file as a .JPG.

#> 
Param(
# Specify the directory to create the files in.
# The files names are a combination of the display name and a timestamp
[Parameter()]
[Alias("Path")]
[string]$Directory = ".",

#The lower the number specified, the higher the compression and therefore the lower the quality of the image. Zero would give you the lowest quality image and 100 the highest.
[Parameter()]
[ValidateRange(0,100)]
[int]$Quality=100,

# By default, only the PRIMARY display is captured
[Parameter()]
[Switch]$AllScreens
)
            
            
Set-StrictMode -Version 2
Add-Type -AssemblyName System.Windows.Forms

if ($AllScreens)
{
    $Capture = [System.Windows.Forms.Screen]::AllScreens 
}
else
{
    $Capture = [System.Windows.Forms.Screen]::PrimaryScreen
}
foreach ($C in $Capture)
{
    $FileName = '{0}-{1}.jpg' -f (Join-Path (Resolve-Path $Directory) ($c.DeviceName -split "\\")[3]),(Get-Date).ToString('yyyyMMdd_HHmmss')
            
            
    $Bitmap = New-Object System.Drawing.Bitmap($C.Bounds.Width, $C.Bounds.Height)
    $G = [System.Drawing.Graphics]::FromImage($Bitmap)
    $G.CopyFromScreen($C.Bounds.Location, (New-Object System.Drawing.Point(0,0)), $C.Bounds.Size)
    $g.Dispose()

    $EncoderParam = [System.Drawing.Imaging.Encoder]::Quality
    $EncoderParamSet = New-Object System.Drawing.Imaging.EncoderParameters(1) 
    $EncoderParamSet.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter($EncoderParam, $Quality) 
    $JPGCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where{$_.MimeType -eq 'image/jpeg'}
    $Bitmap.Save($FileName ,$JPGCodec, $EncoderParamSet)
    $FileSize = [INT]((Get-Childitem $FileName).Length / 1KB)
    Write-Verbose ("Display [$($c.DeviceName)] ScreenCapture saved to File [$FileName] Size [$FileSize] KB")
}
