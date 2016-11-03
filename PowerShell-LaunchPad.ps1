#Requires -Version 3.0

<# 
.SYNOPSIS 
A single point of entry for admins to access a library of PowerShell GUI scripts.
.DESCRIPTION 
The purpose of this PowerShell script is to provide an easy way for team members to bring up one or multiple 
PowerShell scripting tools without having to know the location or parameters for any of them. Also it makes it easier to 
setup one shortcut that can be configured and refreshed without any further changes by anyone whenever scripts are updated 
or added. 
.EXAMPLE 
One way to bring up this tool to avoid having to use the command line is to create a Windows shortcut to PowerShell.exe 
and then call the LaunchPad.ps1 file...

C:\WINDOWS\System32\WindowsPowerShell\v1.0\powershell.exe -File "\\server\share\scripts\PowerShell-LaunchPad.ps1"

*If you need to bypass a local workstation's execution policy, you could include the following in the shortcut target...
-ExecutionPolicy Bypass 

The LaunchPad.ps1 script looks in the $global:lpLaunchFiles directory and pulls up any .ps1 files to populate the drop-down 
with choices.

###
Copyright (c) 2015 Greg Besso

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>


# add assembly used later on for image formatting...
Add-Type -AssemblyName System.Drawing

#
# global variables used later on...
#
# the title of this script...
$global:lpTitle = "PowerShell Tools Launch Pad"

# the UNC path to the network-stored image file to display in the launch pad...
$global:lpLogoURL = "\\server\share\PowerShell\Style\your_logo_page.png" #Powershell-Logo-01.png or your_logo_page.png

# folder where you store the scripts you want to be included in the launch pad drop-down menu...
$global:lpLaunchFiles = '\\server\share\PowerShell\PowerShell-LaunchPad\Scripts'

# link to an imports file that contains any functions that are reusable / referenced by your PowerShell tools that the launch pad links to...
. "\\server\share\PowerShell\PowerShell-Imports\PowerShell-Imports.ps1" 

# setting padding between objects in the form...
$global:lpPaddingH = 30
$global:lpPaddingV = 15

# set max dimension you want to see on the image you choose to use...
$maxDimension = 250

# getting the dimensions of the image file...
$png = New-Object System.Drawing.Bitmap $global:lpLogoURL
$global:lpLogoURLH = $png.Height
$global:lpLogoURLW = $png.Width

# get resized dimensions for the image if needed...
if (($global:lpLogoURLH -gt $maxDimension) -Or ($global:lpLogoURLW -gt $maxDimension)) {

    if ($global:lpLogoURLW -gt $global:lpLogoURLH) {
        $ratio = $maxDimension/$global:lpLogoURLW
        $global:lpLogoURLW = $maxDimension
        $global:lpLogoURLH = $global:lpLogoURLH * $ratio
    }

    if ($global:lpLogoURLH -gt $maxDimension) {
        $ratio = $maxDimension/$global:lpLogoURLH
        $global:lpLogoURLW = $global:lpLogoURLW * $ratio
        $global:lpLogoURLH = $global:lpLogoURLH * $ratio
    }

}


#
# function that brings up the GUI form...
#
function New-LaunchPadStartForm() {
<# 
.SYNOPSIS 
Brings up the GUI form for the launch pad script.
.DESCRIPTION 
The launch pad is just a front door or welcome mat to your library of scripts. An easy to use single point of entry to let someone bring up 
whatever tool is needed at that moment.
.EXAMPLE 
New-LaunchPadStartForm
#>

    #
    # get list of files that the launchpad will list as options...
    #
    $lpFiles = Get-ChildItem "$global:lpLaunchFiles" -Filter *.ps1 -Recurse
    $dropDownChoices = @()
    $lpFiles | ForEach {
        $thisName = $_.BaseName
        $object1 = New-Object PSObject -Property @{
            Name=$thisName         
        }
        $dropDownChoices += $object1
    }
    $dropDownChoices = $dropDownChoices | Sort-Object -Property Name

    # start building the form...
    Write-Output "The Launch Pad form is being prepared, stand by..."
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
    
    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "$global:lpTitle"
    $objForm.Width = 1
    $objForm.Height = 1
    $objForm.AutoSize = $True
    $objForm.StartPosition = "CenterScreen"
    $objForm.BackColor = "#333333"
    $objForm.ForeColor = "#ffffff"
    $Font = New-Object System.Drawing.Font("Lucida Sans Console",10,[System.Drawing.FontStyle]::Regular)
    $objForm.Font = $Font

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {$x=$objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})

    # add an image
    $pictureBox = new-object Windows.Forms.PictureBox
    $pictureBox.Width =  $global:lpLogoURLW
    $pictureBox.Height =  $global:lpLogoURLH   
    $pictureBox.ImageLocation = $global:lpLogoURL
    $pictureBox.SizeMode = 4
    $pictureBox.Location = New-Object Drawing.Point $global:lpPaddingH,$global:lpPaddingV
    $objForm.controls.add($pictureBox)
   

    # first line of the form is the title label...
    $objLabelTitle = New-Object System.Windows.Forms.Label
    $objLabelTitle.Location = New-Object System.Drawing.Size(($pictureBox.Left+$pictureBox.Width+$global:lpPaddingH),$global:lpPaddingV)
    $objLabelTitle.AutoSize = $True 
    $objLabelTitle.Text = "LAUNCH PAD"
    $Font = New-Object System.Drawing.Font("Lucida Sans Console",12,[System.Drawing.FontStyle]::Bold)
    $objLabelTitle.Font = $Font
    $objForm.Controls.Add($objLabelTitle) 

    # second line of the form is an instruction label...
    $objLabelA = New-Object System.Windows.Forms.Label
    $objLabelA.Location = New-Object System.Drawing.Size(($pictureBox.Left+$pictureBox.Width+$global:lpPaddingH),($objLabelTitle.Bottom+$global:lpPaddingV)) 
    $objLabelA.AutoSize = $True 
    $objLabelA.Text = "Pick a tool, any tool..."
    $objForm.Controls.Add($objLabelA) 


    # third line of the form is the drop-down menu with a list of scripts...
    $DropDownChoice = new-object System.Windows.Forms.ComboBox
    $DropDownChoice.Location = new-object System.Drawing.Size(($pictureBox.Left+$pictureBox.Width+$global:lpPaddingH),($objLabelA.Bottom+$global:lpPaddingV))
    $DropDownChoice.Size = new-object System.Drawing.Size(230,20)

    # looping through all ps1 files to populate the drop-down list...    
    $dropDownChoices | ForEach {
        $thisName = $_.Name
        $DropDownChoice.Items.Add("$thisName") | Out-Null
    }
    $objForm.Controls.Add($DropDownChoice)   

    # fourth line of the form is for the OK/cancel buttons...
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(($pictureBox.Left+$pictureBox.Width+$global:lpPaddingH),($DropDownChoice.Bottom+$global:lpPaddingV))
    $OKButton.Size = New-Object System.Drawing.Size(100,23)
    $OKButton.Text = "Launch"

    $OKButton.Add_Click({
        If (($DropDownChoice.Text.Length -gt 1)) {
            # the user selected a script, so it will be loaded...
            $choice = $DropDownChoice.Text
            Write-Host "OK you selected $choice, stand by while that script is loaded"
            $OKButton.Enabled = $False
            $lpFiles | ForEach {
                $checkThis = $_.BaseName
                If ($checkThis -eq $choice) {
                    $useDirectory =  $_.Directory
                    $useFile = $_.Name
                    ."$useDirectory\$useFile"
                }
            }
            $OKButton.Enabled = $True
        }
    })
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(($OKButton.Right + 7),($OKButton.Top)) 
    $CancelButton.Size = New-Object System.Drawing.Size(100,23)
    $CancelButton.Text = "Exit"
    $CancelButton.Add_Click({
        #close any sessions that were opened, then close the form...
        Get-PSSession | Remove-PSSession
        $objForm.Close()
    })
    $objForm.Controls.Add($CancelButton)

    # if true, will ensure the launch pad form hovers in front of other windows, even if not in focus...
    $objForm.Topmost = $True

    # add an icon to the form window...
    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $objForm.Icon = $Icon

    #$objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()
}

#
# call the function that will load the form...
#
New-LaunchPadStartForm