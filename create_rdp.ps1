#Add namespace
Add-Type -AssemblyName System.Windows.Forms

#Variables the Scrip uses
$localgroupe = 'Remotedesktopsuer'              #Local Remotedesktopuser group
$AD_PC_GROUP = 'RDS_ComputerGroup'              #ADGroup for the Remote User or $null if you dont have a group
$AD_USER_GROUP = $null #'RDS_Remoteuser'        #ADGroup for the Computer or $null if you dont have a group
$rdpfilepath = "RDP-File-Folder"                #The path where you want to safe the RDP-File
$rdsgateway = "RDPGateway.domain.com"           #The Remotedesktopservergateway IP or 
#Settings for the Mail:
$SMTPServer ="SMTPorExchangeServer"             #Your SMTP-Server
$From = 'mail@mydomain.com'                     #The mailaddress you want to send the mail from
$Subject2 = "Remote Connection"                 #Subject of the mail
$Body = "Hello, <br><br>
        in the attachments you can find the remoteconnectionfile to your workstation $pc.
        To save the file on your desktop, you need the rightclick the attachment and choose save as and then choose your desktop.
        After this you can double click the file and login to your workstation with your account.
        <br><br>
        Greetings
        <br><br>
        Your IT
        <br><br>
        -----------------------------------------
        <br>This message was generated automatically."  #The body of the mail

#Creates the window
$main = New-Object  System.Windows.Forms.Form

#Size + title of the window
$main.ClientSize = '500,300'
$main.text = "Create RDP"
$main.AutoSize = $true
$main.BackColor = "#ffffff"

#Title at the top of the window
$MainText = New-Object System.Windows.Forms.Label
$MainText.Text = "Create RDP-File"  
$MainText.AutoSize = $true
$MainText.Location = New-Object System.Drawing.Point(20,20)
$MainText.Font = 'Microsoft Sans Serif,13'
$MainText.ForeColor = "#2F6AF4" #Ueberschrift in Blau

#Description text under the title
$description = New-Object System.Windows.Forms.Label
$description.Text = "Here you can create an RDP-File for Users in Homeoffice. 
You need the username and the compoutername. 
The RDP-File will be send as Mail to the user."
$description.AutoSize = $true
$description.Font = 'Microsoft Sans Serif,11'
$description.Location = New-Object System.Drawing.Point(20,60)

#Text over textbox for the username 
$txt_username = New-Object System.Windows.Forms.Label
$txt_username.Location = New-Object System.Drawing.Point(20,160)
$txt_username.Text = "Insert the username:"
$txt_username.Font = "Microsoft Sans Serif,13"
$txt_username.AutoSize = $true
$txt_username.ForeColor = "#2F6AF4"

#Textbox for the username
$txtbox_username = New-Object System.Windows.Forms.TextBox
$txtbox_username.Location = New-Object System.Drawing.Point(20,190)
$txtbox_username.Width = 200
$txtbox_username.Height = 50

#Text over the textbox for the username
$txt_computername = New-Object System.Windows.Forms.Label
$txt_computername.Location = New-Object System.Drawing.Point (20,250)
$txt_computername.Text = "Insert the computername:"
$txt_computername.Font = "Microsoft Sans Serif,13"
$txt_computername.AutoSize = $true
$txt_computername.ForeColor = "#2F6AF4"

#Textbox for the computername
$txtbox_computername = New-Object System.Windows.Forms.TextBox
$txtbox_computername.Location = New-Object System.Drawing.Point(20,280)
$txtbox_computername.Width = 200
$txtbox_computername.Height = 50

#Radiobutton send mail
$checkbox1 = New-Object System.Windows.Forms.CheckBox
$checkbox1.Location = New-Object System.Drawing.Point 20, 310
$checkbox1.Text = "Send Mail"

#Ok Button
$OkButton = New-Object System.Windows.Forms.Button
$OkButton.BackColor = "#a4ba67"
$OkButton.Text = "Ok"
$OkButton.Width = 90
$OkButton.Height = 30
$OkButton.Location = New-Object System.Drawing.Point(400,320)
$OkButton.ForeColor = "#ffffff"
#$OkButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

#Cancel Button
$CancleButton = New-Object System.Windows.Forms.Button
$CancleButton.BackColor = "#ffffff"
$CancleButton.Text = "Abbrechen"
$CancleButton.Width = 90
$CancleButton.Height = 30
$CancleButton.Location = New-Object System.Drawing.Point (300,320)
$CancleButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

#Function if OK-Button was clicked
function CreateFile {
    #Logic of Script    
    $pc = $txtbox_computername.Text  #Only the Computername is needed the SAMAccountname for the computer will be generated
    $User = $txtbox_username.Text   #You can insert the UserPrincipalName oder the SAMAccountname
    
    #Compares the userentries with the AD and gets the UPN and SAMAccountname 
    $userobject = $null
    $userobject = Get-ADUser -Filter * | Where-Object userPrincipalName -Match $User
    if(!$userobject)
    {
        $userobject = Get-ADUser -Identity $User
    }    
    $upn = $userobject.UserPrincipalName
    $samaccountname = $userobject.SAMAccountName
    
    #Pasword does not expire anymore
    Set-ADUser -Identity $samaccountname -PasswordNeverExpires $true

    #Makes the SAMAccountname of the computer
    $pc_samaccountname = $pc + "$"

    #Fügt den Benutzer in die Lokale Remotedesktopbenutzer Grupp hinzu
    Invoke-Command -ScriptBlock {
        Add-LocalGroupMember -Group $Using:localgroupe -Member $Using:samaccountname
    } -ComputerName $pc

    #Fügt den Benutzer den AD-Gruppen hinzu
    if($AD_PC_GROUP)
    {
        Add-ADGroupMember -Identity $AD_PC_GROUP -Members $pc_samaccountname
    }
    if($AD_USER_GROUP)
    {
        Add-ADGroupMember -Identity $AD_USER_GROUP -Members $samaccountname
    }    
    

    #Creates the RDP File
    $filename = $pc + "_" + $samaccountname + '.rdp'
    
    New-Item -Path $rdpfilepath -Name $filename -Value "screen mode id:i:2
    use multimon:i:0
    desktopwidth:i:800
    desktopheight:i:600
    session bpp:i:32
    winposstr:s:0,3,0,0,800,600
    compression:i:1
    keyboardhook:i:2
    audiocapturemode:i:0
    videoplaybackmode:i:1
    connection type:i:7
    networkautodetect:i:1
    bandwidthautodetect:i:1
    displayconnectionbar:i:1
    username:s:$upn
    enableworkspacereconnect:i:0
    disable wallpaper:i:0
    allow font smoothing:i:0
    allow desktop composition:i:0
    disable full window drag:i:1
    disable menu anims:i:1
    disable themes:i:0
    disable cursor setting:i:0
    bitmapcachepersistenable:i:1
    full address:s:$pc
    audiomode:i:0
    redirectprinters:i:1
    redirectcomports:i:0
    redirectsmartcards:i:1
    redirectclipboard:i:1
    redirectposdevices:i:0
    autoreconnection enabled:i:1
    authentication level:i:2
    prompt for credentials:i:0
    negotiate security layer:i:1
    remoteapplicationmode:i:0
    alternate shell:s:
    shell working directory:s:
    gatewayhostname:s:$rdsgateway   
    gatewayusagemethod:i:2
    gatewaycredentialssource:i:4
    gatewayprofileusagemethod:i:1
    promptcredentialonce:i:1
    gatewaybrokeringtype:i:0
    use redirection server name:i:0
    rdgiskdcproxy:i:0
    kdcproxyname:s:
    "
    #Creates and sends an Mail to the user
        #Checkbox has to be checked
    if($checkbox1.Checked)
    {
        $attachment = $rdpfilepath + $filename
        $To = Get-ADUser -Identity $samaccountname -Properties mail
        Send-MailMessage -To $To.mail -Subject $Subject2 -Body $Body -SmtpServer $SMTPServer -From $From -Attachments $attachment -BodyAsHtml -Encoding UTF8
    }
    $txtbox_username.Clear()
    $txtbox_computername.Clear()
    $checkbox1.Checked = $false
}

#Aufgabe wenn auf Ok geklickt wird
$OkButton.Add_Click({CreateFile})

#Fügt die Textblöcke hinzu (dadurch sind sie in der GUI sichtbar)
$main.Controls.AddRange(@($MainText,$description,$CancleButton,$OkButton,$txtbox_username,$txt_username,$txt_computername,$txtbox_computername,$checkbox1))

[void]$main.ShowDialog()