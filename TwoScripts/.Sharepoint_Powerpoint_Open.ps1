param([string] $SharpointPath)
Add-Type -AssemblyName System.Windows.Forms
$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$width = $screen.Bounds.width
$height = $screen.Bounds.height

$ie = New-Object -ComObject InternetExplorer.Application
$ie.Navigate($SharpointPath)
$ie.Visible = $true
$ie.Width  = $width /2
$ie.Height = $height
$ie.Left  = 0
$ie.Top  = 0

Add-type -AssemblyName office
$ppt = New-Object -ComObject powerpoint.Application
$ppt.visible = $true
$new_presentation = $ppt.Presentations.add()
#$ppt.Width  = $width /2
#$ppt.Height = $height
#$ppt.Left  = $width /2
#$ppt.Top  = 0

#$ie | Get-member





