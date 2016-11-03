#Requires -Version 3.0


#
# import the Offboarding global variables, to avoid having to define them 
# in multiple places by the handful of script files. probably the only place with a manually specified path...
#
. "\\server\share\PowerShell\UserManagementAuto\ServerScripts\OffboardingImports.ps1" 





# by default inherits global logo from launch pad, but you can change that if you wanted to...
If ($global:lpLogoURL) {
    $global:usrMgmtLogoURL = $global:lpLogoURL #Powershell-Logo.png
    $global:usrMgmtLogoURLH = $global:lpLogoURLH
    $global:usrMgmtLogoURLW = $global:lpLogoURLW

    # setting padding between objects in the form...
    $global:usrMgmtPaddingH = $global:lpPaddingH
    $global:usrMgmtPaddingV = $global:lpPaddingV
} Else { 
    # if using this script within LaunchPad, you can remove this entire ELSE section :-)
    # add assembly used later on for image formatting...
    Add-Type -AssemblyName System.Drawing
    $global:usrMgmtLogoURL = "\\gbesso-w12r2\SHARE\Style\Powershell-Logo-01.png"

    # set max dimension you want to see on the image you choose to use...
    $maxDimension = 250

    # getting the dimensions of the image file...
    $png = New-Object System.Drawing.Bitmap $global:usrMgmtLogoURL
    $global:usrMgmtLogoURLH = $png.Height
    $global:usrMgmtLogoURLW = $png.Width

    # get resized dimensions for the image if needed...
    if (($global:usrMgmtLogoURLH -gt $maxDimension) -Or ($global:usrMgmtLogoURLW -gt $maxDimension)) {
        if ($global:usrMgmtLogoURLW -gt $global:usrMgmtLogoURLH) {
            $ratio = $maxDimension/$global:usrMgmtLogoURLW
            $global:usrMgmtLogoURLW = $maxDimension
            $global:usrMgmtLogoURLH = $global:usrMgmtLogoURLH * $ratio
        }
        if ($global:usrMgmtLogoURLH -gt $maxDimension) {
            $ratio = $maxDimension/$global:usrMgmtLogoURLH
            $global:usrMgmtLogoURLW = $global:usrMgmtLogoURLW * $ratio
            $global:usrMgmtLogoURLH = $global:usrMgmtLogoURLH * $ratio
        }

    }
    # setting padding between objects in the form...
    $global:usrMgmtPaddingH = 30
    $global:usrMgmtPaddingV = 15
}


# function that gets a list of the computers that are using the PSMANAGE phone home scripts...
function Get-PSManageActiveUsers() {

    $getDC = $global:usrMgmtDomainController
    If (!($sessionDC)) { $sessionDC = New-PSSession -ComputerName $getDC}
    $getInfo = Invoke-Command -Session $sessionDC -ScriptBlock {
        # get input from function calling remote session
        Param ($SamAccountName)

        # do stuff...
        Import-Module ActiveDirectory
        $getInfo = Get-ADUser -Properties * -Filter {(Surname -Like "*") -And (Enabled -eq "True") -And (Mail -like "*")} | Select SamAccountName, Name, Department
        $getInfo

    } -ArgumentList $SamAccountName
    Return $getInfo
    $sessionDC | Remove-PSSession


}

# sends a request to the remote SharePoint server to run the offboarding create task...
function New-OffboardingTasksTrigger() {
<# 
.SYNOPSIS 
Sends remote schtasks request to a computer
.DESCRIPTION 
No more information needed
.PARAMETER computerName
The name of the computer to be reached
.EXAMPLE 
New-OffboardingTasksTrigger -computerName 'gregserver1'
#>
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [string]$computerName
    )

    BEGIN{}
    PROCESS{
        Try {                            
            If (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
                #run the create task, then wait a bit and try to execute the ones that need to be run right away...
                
                Try {
                    $existingQueries = Schtasks.exe /S $computerName /Run /TN "$global:usrMgmtScheduledCreate"
                } Catch {}

                Start-Sleep -s 12

                Try {
                    $existingQueries = Schtasks.exe /S $computerName /Run /TN "$global:usrMgmtScheduledProcess"
                } Catch {}
            }
        } Catch {
            Write-Warning "Error occurred: $_.Exception.Message"
        }
    }
    End {}
}

# creates the GUI form that lets all the stuff happen...
function Get-MassInstallPackageForm() {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 


    $script:preferredDC = "$global:usrMgmtDomainController"

    $getComputers = Get-PSManageActiveUsers
    $getComputers = $getComputers | Sort-Object Name
    $getComputers = $getComputers | Select SamAccountName, Name, Department    


    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "User Off-boarding Form"
    $objForm.AutoSize = $True
    $objForm.StartPosition = "CenterScreen"
    $objForm.BackColor = "#333333"
    $objForm.ForeColor = "#ffffff"
    $Font = New-Object System.Drawing.Font("Lucida Sans Console",10,[System.Drawing.FontStyle]::Regular)
    $objForm.Font = $Font
    $itemY = 0

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {$x=$objTextBox.Text;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})

    # add an image
    $pictureBox = new-object Windows.Forms.PictureBox
    $pictureBox.Width =  $global:usrMgmtLogoURLW
    $pictureBox.Height =  $global:usrMgmtLogoURLH
    $pictureBox.ImageLocation = $global:usrMgmtLogoURL
    $pictureBox.SizeMode = 4
    $pictureBox.Location = New-Object Drawing.Point 10,10
    $objForm.controls.add($pictureBox)

    # add the GUI title
    $objLabelTitle = New-Object System.Windows.Forms.Label
    $objLabelTitle.Location = New-Object System.Drawing.Size(($pictureBox.Left+$pictureBox.Width+$global:usrMgmtPaddingH),$global:usrMgmtPaddingV)
    $objLabelTitle.AutoSize = $True 
    $objLabelTitle.Text = "User Off-boarding Form"
    $Font = New-Object System.Drawing.Font("Lucida Sans Console",12,[System.Drawing.FontStyle]::Bold)
    $objLabelTitle.Font = $Font
    $objForm.Controls.Add($objLabelTitle) 


    $objLabelA.Location = New-Object System.Drawing.Size($objLabelTitle.Left,($objLabelTitle.Bottom+($global:usrMgmtPaddingV*2)))



    # label and datagrid box to list the available computers for installing packages to...
    $labelSelectPackages = New-Object System.Windows.Forms.Label
    $b1 = $pictureBox.Bottom
    $b2 = $DropDownComputer.Bottom
    If ($b1 -gt $b2) {$b3 = $b1} Else { $b3 = $b2} #get bottom of whichever item is lower, to align the grid positioning...
    $labelSelectPackages.Location = New-Object System.Drawing.Size($pictureBox.Left,($b3+($global:usrMgmtPaddingV*2))) 
    $labelSelectPackages.AutoSize = $True 
    $labelSelectPackages.Text = "Select the user you want to off-board... "
    $objForm.Controls.Add($labelSelectPackages)
    
    $dataGrid1 = New-Object System.Windows.Forms.DataGridView
    $dataGrid1.Width = 700
    $dataGrid1.Height = 400
    $dataGrid1.DefaultCellStyle.ForeColor = "#000000"
    $dataGrid1.Name = "dataGrid1"
    $array = New-Object System.Collections.ArrayList
    $array.AddRange($getComputers)
    $dataGrid1.DataSource = $array
    $dataGrid1.ReadOnly = $True
    $dataGrid1.Location = New-Object System.Drawing.Size($labelSelectPackages.Left,($labelSelectPackages.Bottom+($global:usrMgmtPaddingV/2)))
    $dataGrid1.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $objForm.Controls.Add($dataGrid1)


    # label and datagrid box to show chosen computers
    $labelSelectedPackages = New-Object System.Windows.Forms.Label
    $labelSelectedPackages.Location = New-Object System.Drawing.Size(($dataGrid1.Left),($dataGrid1.Bottom+($global:usrMgmtPaddingV*2)))
    $labelSelectedPackages.AutoSize = $True 
    $labelSelectedPackages.Text = "Chosen user..."
    $objForm.Controls.Add($labelSelectedPackages)
    
    $dataGrid2 = New-Object System.Windows.Forms.DataGridView
    $dataGrid2.Width = 700
    $dataGrid2.Height = 50
    $dataGrid2.DefaultCellStyle.ForeColor = "#000000"
    $dataGrid2.Name = "dataGrid2"
    $dataGrid2.ReadOnly = $True
    $array = New-Object System.Collections.ArrayList    
    $dataGrid2.Location = New-Object System.Drawing.Size(($dataGrid1.Left),($labelSelectedPackages.Bottom+($global:usrMgmtPaddingV/2)))
    $dataGrid2.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::AllCells
    $objForm.Controls.Add($dataGrid2)




    # get the email subject info...
    $objLabelSubject = New-Object System.Windows.Forms.Label
    $objLabelSubject.Location = New-Object System.Drawing.Size(($dataGrid2.Left),($dataGrid2.Bottom+($global:usrMgmtPaddingV*2))) 
    $objLabelSubject.AutoSize = $True 
    $objLabelSubject.Text = "Copy/paste the ticket Reply email subject text..."
    $objForm.Controls.Add($objLabelSubject) 


    $objTextBoxSubject = New-Object System.Windows.Forms.TextBox 
    $objTextBoxSubject.Location = New-Object System.Drawing.Size(($objLabelSubject.Left),($objLabelSubject.Bottom+($global:usrMgmtPaddingV/2)))
    $objTextBoxSubject.Size = New-Object System.Drawing.Size(700,20) 
    $objForm.Controls.Add($objTextBoxSubject)



    # buttons to continue/cancel on the form...
    $AddButton = New-Object System.Windows.Forms.Button
    $AddButton.Location = New-Object System.Drawing.Size(($dataGrid2.Left),($objTextBoxSubject.Bottom+$global:usrMgmtPaddingV*2))
    $AddButton.Size = New-Object System.Drawing.Size(100,23)
    $AddButton.Text = "Update List"
    $AddButton.Enabled = $True
    $AddButton.Add_Click({       
        $listUpdate = @()
        $selectedItems = $dataGrid1.SelectedRows

        $thisCount = $selectedItems.Count
        If ($thisCount -gt 0) {
            $selectedItems | ForEach-Object {
                $thisIndex = $_.Index
                $thisPackageName = $dataGrid1.Rows[$thisIndex].Cells[0].Value
                $thisPackageID = $dataGrid1.Rows[$thisIndex].Cells[1].Value
                $this2 = $dataGrid1.Rows[$thisIndex].Cells[2].Value

                $object2 = [pscustomobject]@{
                    SamAccountName = $thisPackageName;
                    Name = $thisPackageID; 
                    Department = $this2;                 
                }
                $listUpdate += $object2  
            }

            $listUpdate = $listUpdate | Sort-Object ChosenPackageName
            $array2 = New-Object System.Collections.ArrayList
            $array2.AddRange(@($listUpdate))
            $dataGrid2.DataSource = $array2

            $VerifyButton.Enabled = $True

            $objLabelResults.Text = 'OK, confirm your selection...'
        }        
    })
    $objForm.Controls.Add($AddButton)


    # button for confirming choice...
    $VerifyButton = New-Object System.Windows.Forms.Button
    $VerifyButton.Location = New-Object System.Drawing.Size(($AddButton.Right+$global:usrMgmtPaddingH),$AddButton.Top)
    $VerifyButton.Size = New-Object System.Drawing.Size(100,23)
    $VerifyButton.Text = "Confirm"
    $VerifyButton.Enabled = $False
    $VerifyButton.Add_Click({        
        $chosenUser = $dataGrid2.SelectedRows

        If ($chosenUser.Length -lt 1) {
            $objLabelResults.Text = 'Select a user to off-board!'
        } Else {
            $ProcessButton.Enabled = $True
            $objLabelResults.Text = 'OK, click the Proceed button...'
        }
    })
    $objForm.Controls.Add($VerifyButton)


    # button for creating tasks...
    $ProcessButton = New-Object System.Windows.Forms.Button
    $ProcessButton.Location = New-Object System.Drawing.Size(($VerifyButton.Right+$global:usrMgmtPaddingH),$AddButton.Top)
    $ProcessButton.Size = New-Object System.Drawing.Size(100,23)
    $ProcessButton.Text = "Proceed"
    $ProcessButton.Enabled = $False
    $ProcessButton.Add_Click({       
    

        $chosenSam = $dataGrid2.Rows[0].Cells[0].Value
        $emailSubject = $objTextBoxSubject.Text
        $thisUser = [Environment]::UserName
        $thisComputer = [Environment]::MachineName
        $thisWhen = Get-Date -Format yyyy-mm-dd-hh-mm-ss-ms

        If ($chosenSam.Length -lt 1) {
            $objLabelResults.Text = 'Select a user to off-board!'
        } ElseIf ($emailSubject.Length -lt 1) {
            $objLabelResults.Text = 'Paste in the email subject!'
        } Else {
            $objLabelResults.Text = 'Standby, your tasks are being created...'  
            
            $VerifyButton.Enabled = $False
            $ProcessButton.Enabled = $False
                                  
            
            $output = @()
            $object1 = [pscustomobject]@{
                SamAccountName = $chosenSam;
                EmailSubject = $emailSubject;
                ThisUser = $thisUser;
                ThisComputer = $thisComputer;
                ThisWhen = $thisWhen;                   
            }
            $output += $object1 

            $fileName = $global:usrMgmtRequestsShare + "\" + ($thisWhen).ToString() + "-" + $chosenSam + ".xml"

            $output | Export-CliXml $fileName -Force
            

            # call the function to create the tasks once the data is all ready...
            #New-MultiplePackageInstallerTasks -chosenSam $chosenSam -emailSubject $emailSubject -spWeb $global:usrMgmtSPWeb
            New-OffboardingTasksTrigger -computerName "$global:usrMgmtProcessServer"

            #let admin know the tasks are created...
            $objLabelResults.Text = 'OK your tasks are now in SharePoint Online...'
        }
    })
    $objForm.Controls.Add($ProcessButton)

    # cancel/exit button...
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(($ProcessButton.Right+$global:usrMgmtPaddingH),$AddButton.Top)
    $CancelButton.Size = New-Object System.Drawing.Size(100,23)
    $CancelButton.Text = "Exit"
    $CancelButton.Add_Click({
        #close any sessions that were opened, then close the form...
        Get-PSSession | Remove-PSSession
        $objForm.Close()
    })
    $objForm.Controls.Add($CancelButton)

    # display the results
    $objLabelResults = New-Object System.Windows.Forms.Label
    $objLabelResults.Location = New-Object System.Drawing.Size(($CancelButton.Right+$global:usrMgmtPaddingH),$AddButton.Top)
    $objLabelResults.AutoSize = $True 
    $objLabelResults.Text = ""
    $objForm.Controls.Add($objLabelResults) 
    $objForm.Topmost = $True

    $Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
    $objForm.Icon = $Icon

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()
}

# bring up the GUI form that will get the process started...
Get-MassInstallPackageForm
