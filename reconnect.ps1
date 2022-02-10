<#  
.SYNOPSIS  
    Utility for PC profile owners, which returns the profile redirect and copies the files back down to the desktop. - Only new files or files that don't exist on the local computer are copied.
.DESCRIPTION  
    ***The folders (and sub-folders) copied back to the local machine are***
        - Desktop,
        - My Documents (including Outlook),
        - Favorites,
        - My Pictures
    Type of Script - Passive (requirements necessitate that the TM manage his/her account by running the script themselves.)
    Note - If files are already present they aren't over writen unless a newer copy exists. 
.NOTES  
    File Name  : Undo-UST.ps1
    Version    : 1.0  
    Caveats    : Used-in Midwest Region (ONLY)
                 Not intended for laptop users!
    Created    : 11/27/2018 by Jim.Berkenbaugh@Wholefoods.com
    Requires   : PowerShell V2 [or greater]
.EXAMPLE 
.LINK

#>

#######################################################################################
# START OF SCRIPT!
# This begining of the script provides variables for unattended running.
#   - The date is selected.
#   - The windows version is determined.
#   - The three letter store abbreviation is determined by the computer name.
#   - The user folder directory is determined.
#     BEFORE EXCUTION OF THE SCRIPT EXISTANCE OF A NETWORK DIRECTORY IS DETERMINED.
#######################################################################################

$erroractionpreference = "SilentlyContinue"

# Set-ExecutionPolicy -Scope CurrentUser Unrestricted -force

Set-Location C:

# Message Bar Variables
$Activity             = "This will automatically re-connect your PST files to your Outlook."
$Task                 = "...connecting:)"
Write-Progress -Activity $Activity -Status $Task

function reconnect {
clear

Write-Host "`n"
Write-Host "`n"
Write-Host "`n"

#######################################################################################
# LOCAL PST RECONNECTION...
# Disconnect ALL PST files from Outlook - These will be network linked
# From the copy function above all PST files have been copied to a local location.
# Reconnect the standard "Local Mail" PST, which is directed to firstname.lastname.pst
#######################################################################################

# Wait Message
$Task                 = "Local PST Reconnection"
Write-Progress -Activity $Activity -Status $Task

# Connecting the new Local Mail PST to the local mail client.

if (!(Test-Path "$env:userprofile\My Documents\Outlook Files"))
	{
	md "$env:userprofile\My Documents\Outlook Files"
	}

    Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
    $outlook = new-object -comobject outlook.application
    $namespace = $outlook.GetNameSpace("MAPI")
    dir “$env:USERPROFILE\Documents\Outlook Files\*.pst” | % { $namespace.AddStore($_.FullName) }

    Write-Host "Locally reconnected all Outlook files..."
    Write-Host "OPERATION COMPLETE!"

# Wait Message
$Task                 = "COMPLETE!"
Write-Progress -Activity $Activity -Status $Task

##########################
# END OF SCRIPT!
##########################

}

############################################################################################################
# The rest of the script handles the GUI

# Load the assembly since it isn't by default 

[void][reflection.assembly]::loadwithpartialname("System.Windows.Forms") 
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

# Create a top-level form 

$form=new-object windows.forms.form 

# Set the text property 

$form.text="Reconnection of Personal Folders to Outlook" 
$form.Size = New-object System.Drawing.Size (480,300)
$form.StartPosition = "CenterScreen"

$button2=new-object windows.forms.button
$button2.text="EMAIL"
$button2.add_click({reconnect})
$button2.Location = New-Object System.Drawing.Size(20,20)
$form.controls.add($button2) 

$Labeltext = @"
Reconnect Personal Folders (PSTs, Archive) folders
to Outlook.

CAUTION! Before running this close your email.

    ============= DIRECTIONS =============
1- Close your email,
2- Run.
3- Click the Close button when done. - The message 
   bar at the top of the results will say Complete!
4- Reopen your Outlook and voila.
5- (Your folders will be collapsed and have an arrow
   next to them; click this to expand the other folders.

"@

$label1 = New-Object windows.forms.label
$label1.height = "200"
$label1.width = "300"
$label1.set_text($Labeltext)
$label1.location = New-Object system.Drawing.Size(130,10)
$form.Controls.add($label1)

$buttonQuit=new-object windows.forms.button 
$buttonQuit.text="Close" 
$buttonQuit.add_click({$form.close()}) 
$buttonQuit.Location = New-Object System.Drawing.Size(20,110)
$form.controls.add($buttonQuit) 

# Make this active when shown 

$form.add_shown({$form.activate()}) 
$form.showdialog() 