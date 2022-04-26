Function Start-IncomingTicket{
#Creating 

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
# Creating the form
$IncomingCall                            = New-Object system.Windows.Forms.Form
$IncomingCall.ClientSize                 = New-Object System.Drawing.Point(600,600)
$IncomingCall.text                       = "Incoming Call Gen"
$IncomingCall.TopMost                    = $false
$IncomingCall.AutoSize                   = $true

############################# Incoming Call Ticket Generator ##################################

$Title                           = New-Object system.Windows.Forms.Label
$Title.text                      = "Incoming call Ticket Generator"
$Title.AutoSize                  = $true
$Title.width                     = 25
$Title.height                    = 10
$Title.location                  = New-Object System.Drawing.Point(90,2)
$Title.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',20)

#Ticket number input box
$ICTicketTitle                   = New-Object system.Windows.Forms.TextBox
$ICTicketTitle.multiline         = $true
$ICTicketTitle.width             = 300
$ICTicketTitle.height            = 20
$ICTicketTitle.location          = New-Object System.Drawing.Point(150,60)
$ICTicketTitle.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Ticket number label
$ICTicketTitleLabel              = New-Object system.Windows.Forms.Label
$ICTicketTitleLabel.text         = "Enter Problem Title"
$ICTicketTitleLabel.AutoSize     = $true
$ICTicketTitleLabel.width        = 25
$ICTicketTitleLabel.height       = 10
$ICTicketTitleLabel.location     = New-Object System.Drawing.Point(10,60)
$ICTicketTitleLabel.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Ticket issue input box
$ICissuebody                     = New-Object system.Windows.Forms.TextBox
$ICissuebody.multiline           = $true
$ICissuebody.AcceptsReturn       = $true
$ICissuebody.ScrollBars          = "Vertical" 
$ICissuebody.width               = 300
$ICissuebody.height              = 150
$ICissuebody.location            = New-Object System.Drawing.Point(150,80)
$ICissuebody.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Ticket issue label
$ICissuebodylabel                = New-Object system.Windows.Forms.Label
$ICissuebodylabel.text           = "Enter Issue description"
$ICissuebodylabel.AutoSize       = $true
$ICissuebodylabel.width          = 25
$ICissuebodylabel.height         = 10
$ICissuebodylabel.location       = New-Object System.Drawing.Point(10,80)
$ICissuebodylabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Steps taken input box
$ICstepstakenbody                 = New-Object system.Windows.Forms.TextBox
$ICstepstakenbody.AcceptsReturn   = $true
$ICstepstakenbody.ScrollBars      = "Vertical"
$ICstepstakenbody.multiline       = $true
$ICstepstakenbody.width           = 300
$ICstepstakenbody.height          = 200
$ICstepstakenbody.location        = New-Object System.Drawing.Point(150,230)
$ICstepstakenbody.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Steps taken Label
$ICstepstakenlabel                = New-Object system.Windows.Forms.Label
$ICstepstakenlabel.text           = "Enter Extra Information"
$ICstepstakenlabel.AutoSize       = $true
$ICstepstakenlabel.width          = 25
$ICstepstakenlabel.height         = 10
$ICstepstakenlabel.location       = New-Object System.Drawing.Point(10,230)
$ICstepstakenlabel.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Asset number input box
$ICAssetBody                      = New-Object system.Windows.Forms.TextBox
$ICAssetBody.multiline            = $false
$ICAssetBody.width                = 100
$ICAssetBody.height               = 20
$ICAssetBody.location             = New-Object System.Drawing.Point(150,430)
$ICAssetBody.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


#Asset number label
$ICAssetLabel                     = New-Object system.Windows.Forms.Label
$ICAssetLabel.text                = "Asset Number"
$ICAssetLabel.AutoSize            = $true
$ICAssetLabel.width               = 25
$ICAssetLabel.height              = 10
$ICAssetLabel.location            = New-Object System.Drawing.Point(10,430)
$ICAssetLabel.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Contact number input box
$ICcontactbody                    = New-Object system.Windows.Forms.TextBox
$ICcontactbody.multiline          = $false
$ICcontactbody.width              = 100
$ICcontactbody.height             = 20
$ICcontactbody.location           = New-Object System.Drawing.Point(150,450)
$ICcontactbody.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Contact number label
$ICcontactlabel                   = New-Object system.Windows.Forms.Label
$ICcontactlabel.text              = "Contact No."
$ICcontactlabel.AutoSize          = $true
$ICcontactlabel.width             = 25
$ICcontactlabel.height            = 10
$ICcontactlabel.location          = New-Object System.Drawing.Point(10,450)
$ICcontactlabel.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#User location input box
$IClocationbody                   = New-Object system.Windows.Forms.TextBox
$IClocationbody.multiline         = $false
$IClocationbody.width             = 100
$IClocationbody.height            = 20
$IClocationbody.location          = New-Object System.Drawing.Point(150,470)
$IClocationbody.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#User location label
$IClocationlabel                  = New-Object system.Windows.Forms.Label
$IClocationlabel.text             = "Users location"
$IClocationlabel.AutoSize         = $true
$IClocationlabel.width            = 25
$IClocationlabel.height           = 10
$IClocationlabel.location         = New-Object System.Drawing.Point(10,470)
$IClocationlabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button to write ticket out
$ICbtnWriteTicket                 = New-Object system.Windows.Forms.Button
$ICbtnWriteTicket.text            = "Run"
$ICbtnWriteTicket.width           = 60
$ICbtnWriteTicket.height          = 30
$ICbtnWriteTicket.location        = New-Object System.Drawing.Point(260,450)
$ICbtnWriteTicket.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#User location label
$ICUserLabel                  = New-Object system.Windows.Forms.Label
$ICUserLabel.text             = "Users Name"
$ICUserLabel.AutoSize         = $true
$ICUserLabel.width            = 25
$ICUserLabel.height           = 10
$ICUserLabel.location         = New-Object System.Drawing.Point(10,40)
$ICUserLabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#User location input box
$ICUserVar                   = New-Object system.Windows.Forms.TextBox
$ICUserVar.multiline         = $false
$ICUserVar.width             = 100
$ICUserVar.height            = 10
$ICUserVar.location          = New-Object System.Drawing.Point(150,40)
$ICUserVar.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)



$IncomingCall.Controls.AddRange(@($title,$ICUserVar,$ICUserLabel,$ICTicketTitleLabel,$ICTicketTitle,$ICissuebody,$ICissuebodylabel,$ICstepstakenbody,$ICstepstakenlabel,$ICAssetBody,$ICAssetLabel,$ICcontactbody,$ICcontactlabel,$IClocationbody,$IClocationlabel,$ICbtnWriteTicket))

$ICbtnWriteTicket.Add_Click({write-IncomingTicket $IncomingCall.Close()})

######################################## Functions #########################################################

function write-IncomingTicket {
#creates the incoming ticket file and 

$ICsubjectline = $ICTicketTitle.Text
$ICBodyInfo = $ICissuebody.Text
$ICExtras = $ICstepstakenbody.Text
$Asset = $ICAssetBody.Text
$Location = $IClocationbody.text
$Contact = $ICcontactbody.Text

$finalICticket = "$ICBodyInfo
$ICExtras

Asset Number: $Asset
Contact: $Contact
Location: $Location
"
$finalICticket | Set-Clipboard

$AddICNote = "$ICsubjectline

$ICBodyInfo
$ICExtras

Asset Number: $Asset
Contact: $Contact
Location: $Location
-----------------------------------------------------------------------------
"

$masterpath = "C:\tickets\MasterRecord.txt"

Add-content $masterpath -Value $AddICNote 
start-sleep -seconds 0.2
Invoke-Item $masterpath

$ICTicketTitle.Text = ""
$ICissuebody.Text = ""
$ICstepstakenbody.Text = ""
$ICAssetBody.Text = ""
$IClocationbody.text = ""
$ICcontactbody.Text = ""

}


[void]$IncomingCall.ShowDialog()

}