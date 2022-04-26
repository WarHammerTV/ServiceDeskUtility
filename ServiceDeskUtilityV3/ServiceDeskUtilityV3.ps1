<#
Service Desk Utility V2.0
Created By: Owen Harris-Evans
\\W4rH4mm3r//

Version Created for Viridor Specifically

#>

#Stops powershell console from opening in the back ground
$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()
# Creating the form
$MainForm                            = New-Object system.Windows.Forms.Form
$MainForm.ClientSize                 = New-Object System.Drawing.Point(1000,600)
$MainForm.text                       = "Service Desk Utility"
$MainForm.TopMost                    = $false
$MainForm.AutoSize                   = $true

#Adding the Title
$SDTitle                             = New-Object System.Windows.Forms.Label
$SDTitle.Text                        = "Service Desk Utility"
$SDTitle.AutoSize                    = $true
$SDTitle.Width                       = 50
$SDTitle.Height                      = 20
$SDTitle.Location                    = New-Object System.Drawing.Point(320,20)
$SDTitle.Font                        = New-Object System.Drawing.Font('Microsoft Sans Serif',30)

############################ Client Specific Generator #########################################

#Adding in the client specific label
$Clientspeclabel                     = New-Object System.Windows.Forms.Label
$Clientspeclabel.Text                = "Viridor Specific Generators"
$Clientspeclabel.AutoSize            = $true
$Clientspeclabel.Width               = 25
$Clientspeclabel.Height              = 10
$Clientspeclabel.Location            = New-Object System.Drawing.Point(50,100)
$Clientspeclabel.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Clientspeclabel.ForeColor           = "Green"

#Button for the Excel Resolve
$btnExcelResolve                     = New-Object system.Windows.Forms.Button
$btnExcelResolve.text                = "Excel Resolve"
$btnExcelResolve.width               = 70
$btnExcelResolve.height              = 50
$btnExcelResolve.location            = New-Object System.Drawing.Point(20,120)
$btnExcelResolve.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button to gather Viridor basic Dataset
$btnDatasetFunction                  = New-Object system.Windows.Forms.Button
$btnDatasetFunction.text             = "Dataset Request"
$btnDatasetFunction.width            = 70
$btnDatasetFunction.height           = 50
$btnDatasetFunction.location         = New-Object System.Drawing.Point(90,120)
$btnDatasetFunction.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button to create citrix resolve
$btnCitrix                          = New-Object system.Windows.Forms.Button
$btnCitrix.text                     = "Citrix Resolve"
$btnCitrix.width                    = 70
$btnCitrix.height                   = 50
$btnCitrix.location                 = New-Object System.Drawing.Point(160,120)
$btnCitrix.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button for Dynamics 365 incoming tickets
$btnD365Gen                 = New-Object system.Windows.Forms.Button
$btnD365Gen.text             = "D365 Incoming"
$btnD365Gen.width            = 70
$btnD365Gen.height           = 50
$btnD365Gen.location         = New-Object System.Drawing.Point(20,170)
$btnD365Gen.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

##################### Generic Generators ####################################################

#Label for the Generic Generator
$genericLabel                       = New-Object System.Windows.Forms.Label
$genericLabel.Text                  = "Generic Generators"
$genericLabel.AutoSize              = $true
$genericLabel.Width                 = 25
$genericLabel.Height                = 10
$genericLabel.Location              = New-Object System.Drawing.Point(320,100)
$genericLabel.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$genericLabel.ForeColor             = "Red"

#Button to create a hello response for Live Chats
$btnHello                  = New-Object system.Windows.Forms.Button
$btnHello.text             = "Hello Gen"
$btnHello.width            = 70
$btnHello.height           = 50
$btnHello.location         = New-Object System.Drawing.Point(280,120)
$btnHello.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button to create a live chat email
$btnLiveChat                  = New-Object system.Windows.Forms.Button
$btnLiveChat.text             = "Live Chat"
$btnLiveChat.width            = 70
$btnLiveChat.height           = 50
$btnLiveChat.location         = New-Object System.Drawing.Point(350,120)
$btnLiveChat.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button for standard resolve
$btnStdResolve                  = New-Object system.Windows.Forms.Button
$btnStdResolve.text             = "Standard Resolve"
$btnStdResolve.width            = 70
$btnStdResolve.height           = 50
$btnStdResolve.location         = New-Object System.Drawing.Point(420,120)
$btnStdResolve.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button to run incoming ticket 
$BtnIncomingCall                   = New-Object System.Windows.Forms.Button
$BtnIncomingCall.text             = "Incoming Call"
$BtnIncomingCall.width            = 70
$BtnIncomingCall.height           = 50
$BtnIncomingCall.location         = New-Object System.Drawing.Point(280,170)
$BtnIncomingCall.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

################################### Input Box ###############################################

#Creating a input Box label
$UsernameInput                        = New-Object system.Windows.Forms.TextBox
$UsernameInput.multiline               = $false
$UsernameInput.width                   = 200
$UsernameInput.height                  = 20
$UsernameInput.location                = New-Object System.Drawing.Point(190,250)
$UsernameInput.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$UsernameLabel                          = New-Object system.Windows.Forms.Label
$UsernameLabel.text                     = "Input Box"
$UsernameLabel.AutoSize                 = $true
$UsernameLabel.width                    = 25
$UsernameLabel.height                   = 10
$UsernameLabel.location                 = New-Object System.Drawing.Point(120,252)
$UsernameLabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


########################### Ticket Generator ###############################################

$Title                           = New-Object system.Windows.Forms.Label
$Title.text                      = "Ticket Generator"
$Title.AutoSize                  = $true
$Title.width                     = 25
$Title.height                    = 10
$Title.location                  = New-Object System.Drawing.Point(700,80)
$Title.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',15)

#Ticket number input box
$TicketTitle                         = New-Object system.Windows.Forms.TextBox
$TicketTitle.multiline               = $true
$TicketTitle.width                   = 150
$TicketTitle.height                  = 20
$TicketTitle.location                = New-Object System.Drawing.Point(700,120)
$TicketTitle.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Ticket number label
$TicketTitleLabel                          = New-Object system.Windows.Forms.Label
$TicketTitleLabel.text                     = "Enter Ticket Number"
$TicketTitleLabel.AutoSize                 = $true
$TicketTitleLabel.width                    = 25
$TicketTitleLabel.height                   = 10
$TicketTitleLabel.location                 = New-Object System.Drawing.Point(575,120)
$TicketTitleLabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Ticket issue input box
$issuebody                        = New-Object system.Windows.Forms.TextBox
$issuebody.multiline               = $true
$issuebody.AcceptsReturn           = $true
$issuebody.ScrollBars              = "Vertical" 
$issuebody.width                   = 300
$issuebody.height                  = 150
$issuebody.location                = New-Object System.Drawing.Point(700,150)
$issuebody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Ticket issue label
$issuebodylabel                          = New-Object system.Windows.Forms.Label
$issuebodylabel.text                     = "Enter Issue description"
$issuebodylabel.AutoSize                 = $true
$issuebodylabel.width                    = 25
$issuebodylabel.height                   = 10
$issuebodylabel.location                 = New-Object System.Drawing.Point(560,150)
$issuebodylabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Steps taken input box
$stepstakenbody                        = New-Object system.Windows.Forms.TextBox
$stepstakenbody.AcceptsReturn           = $true
$stepstakenbody.ScrollBars              = "Vertical"
$stepstakenbody.multiline               = $true
$stepstakenbody.width                   = 300
$stepstakenbody.height                  = 200
$stepstakenbody.location                = New-Object System.Drawing.Point(700,305)
$stepstakenbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Steps taken Label
$stepstakenlabel                          = New-Object system.Windows.Forms.Label
$stepstakenlabel.text                     = "Enter Steps Taken"
$stepstakenlabel.AutoSize                 = $true
$stepstakenlabel.width                    = 25
$stepstakenlabel.height                   = 10
$stepstakenlabel.location                 = New-Object System.Drawing.Point(560,305)
$stepstakenlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Asset number input box
$vwmbody                        = New-Object system.Windows.Forms.TextBox
$vwmbody.multiline               = $false
$vwmbody.width                   = 100
$vwmbody.height                  = 20
$vwmbody.location                = New-Object System.Drawing.Point(700,510)
$vwmbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


#Asset number label
$vwmlabel                          = New-Object system.Windows.Forms.Label
$vwmlabel.text                     = "Please enter Asset Number"
$vwmlabel.AutoSize                 = $true
$vwmlabel.width                    = 25
$vwmlabel.height                   = 10
$vwmlabel.location                 = New-Object System.Drawing.Point(520,510)
$vwmlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Contact number input box
$contactbody                        = New-Object system.Windows.Forms.TextBox
$contactbody.multiline               = $false
$contactbody.width                   = 100
$contactbody.height                  = 20
$contactbody.location                = New-Object System.Drawing.Point(700,530)
$contactbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Contact number label
$contactlabel                          = New-Object system.Windows.Forms.Label
$contactlabel.text                     = "Please enter Contact No."
$contactlabel.AutoSize                 = $true
$contactlabel.width                    = 25
$contactlabel.height                   = 10
$contactlabel.location                 = New-Object System.Drawing.Point(520,530)
$contactlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#User location input box
$locationbody                        = New-Object system.Windows.Forms.TextBox
$locationbody.multiline               = $false
$locationbody.width                   = 100
$locationbody.height                  = 20
$locationbody.location                = New-Object System.Drawing.Point(700,550)
$locationbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#User location label
$locationlabel                          = New-Object system.Windows.Forms.Label
$locationlabel.text                     = "Please enter users location"
$locationlabel.AutoSize                 = $true
$locationlabel.width                    = 25
$locationlabel.height                   = 10
$locationlabel.location                 = New-Object System.Drawing.Point(520,550)
$locationlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Button to write ticket out
$btnWriteTicket                  = New-Object system.Windows.Forms.Button
$btnWriteTicket.text             = "Run"
$btnWriteTicket.width            = 60
$btnWriteTicket.height           = 30
$btnWriteTicket.location         = New-Object System.Drawing.Point(850,525)
$btnWriteTicket.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

##################################Password Generator##########################################################

#Title for password gen ticket
$PassgenLabel                          = New-Object system.Windows.Forms.Label
$PassgenLabel.text                     = "Password Generator"
$PassgenLabel.AutoSize                 = $true
$PassgenLabel.width                    = 25
$PassgenLabel.height                   = 10
$PassgenLabel.location                 = New-Object System.Drawing.Point(10,450)
$PassgenLabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',18)

#Button to generate new Strong password
$BtnGenpass                     = New-Object System.Windows.Forms.Button
$BtnGenpass.Text                = "Strong"
$BtnGenpass.Width               = 60
$BtnGenpass.Height              = 40
$BtnGenpass.Location            = new-object System.Drawing.point(10,490)
$BtnGenpass.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Output box for Strong password
$PassGenOutput                         = New-Object system.Windows.Forms.TextBox
$PassGenOutput.multiline               = $false
$PassGenOutput.width                   = 150
$PassGenOutput.height                  = 50
$PassGenOutput.location                = New-Object System.Drawing.Point(90,490)
$PassGenOutput.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',15)

#Button to generate new Simple password
$BtnSmplGenpass                     = New-Object System.Windows.Forms.Button
$BtnSmplGenpass.Text                = "Simple"
$BtnSmplGenpass.Width               = 60
$BtnSmplGenpass.Height              = 40
$BtnSmplGenpass.Location            = new-object System.Drawing.point(10,525)
$BtnSmplGenpass.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#Output box for Simple password
$PassSmplGenOutput                         = New-Object system.Windows.Forms.TextBox
$PassSmplGenOutput.multiline               = $false
$PassSmplGenOutput.width                   = 150
$PassSmplGenOutput.height                  = 50
$PassSmplGenOutput.location                = New-Object System.Drawing.Point(90,535)
$PassSmplGenOutput.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',15)

############################### Control form ###################################################
$MainForm.controls.AddRange(@($SDTitle,$Clientspeclabel,$genericLabel,$btnExcelResolve,$btnDatasetFunction,$btnCitrix,$UsernameInput,$UsernameLabel,$btnHello,$btnLiveChat,$btnStdResolve,$Title,$TicketTitle,$TicketTitleLabel,$issuebody,$issuebodylabel,$stepstakenbody,$stepstakenlabel,$vwmbody,$vwmlabel,$contactbody,$contactlabel,$locationbody,$locationlabel,$btnWriteTicket,$btnD365Gen,$BtnGenpass,$PassgenLabel,$PassGenOutput,$BtnSmplGenpass,$PassSmplGenOutput,$BtnIncomingCall))

#Adding Functions to buttons
$btnExcelResolve.Add_Click({Write-ExcelResponse})
$btnDatasetFunction.Add_Click({Write-Dataset})
$btnStdResolve.Add_click({write-stdresponse})
$btnCitrix.Add_click({write-citrixsolve})
$btnHello.Add_Click({write-hello})
$btnLiveChat.Add_click({write-livechat})
$btnWriteTicket.Add_Click({test-user})
$btnD365Gen.Add_Click({Write-D365Gen})
$BtnGenpass.Add_Click({Show-Pasword})
$BtnSmplGenpass.Add_Click({Show-SmplPass})
$BtnIncomingCall.Add_Click({Start-IncomingTicket})

## Functions ##

function Write-Dataset { 
    #Function to generate standard viridor dataset request
    $datasetfinal = "Can you please just confirm a few details for me: 1)The VWM number of your device 2) The best contact number for yourself 3) Your current location"

    $datasetfinal | Set-Clipboard
    
}

function Show-Pasword {

    $NewPasswordGenerated = Invoke-WebRequest -Uri https://www.dinopass.com/password/strong | Select-Object -ExpandProperty content
   
   $PassGenOutput.Text = $NewPasswordGenerated
   
   $NewPasswordGenerated | Set-Clipboard

   }

Function Show-SmplPass{

    $NewSmplPassword = Invoke-WebRequest -Uri https://www.dinopass.com/password/simple | Select-Object -ExpandProperty content 

    $PassSmplGenOutput.Text = $NewSmplPassword

    $NewSmplPassword | Set-Clipboard

}

function Write-ExcelResponse {
    #Function to generate resolve for excel issue
    $excelresponseuser = $UsernameInput.Text
    $excelresponsefinal = "Signed $excelresponseuser out of excel and back in. Reopening the report from gatehouse and error resolved"

    $excelresponsefinal | Set-Clipboard 

}

function write-stdresponse {
    #Function to generate a generic resolve for tickets
    $stdresponseuser = $UsernameInput.Text

    $stdresponsefinal = "Hello $stdresponseuser,
Any further issues please don't hesitate  to get back in contact and one of our engineers will be happy to assist you further!
    
Many Thanks
-Littlefish"

    $stdresponsefinal | Set-Clipboard 
} 

function write-citrixsolve {
    #Function to generate citrix resolve for ticket
    $citrixfinal = "Ended all citrix tasks in task manager and restarted citrix workspace. User signed back in and restarted gatehouse program"
    $citrixfinal | Set-Clipboard
    
    }

    function write-hello {
        #Function to create hello text for live chats
        $getuserhello = $UsernameInput.Text
        $morninghello = "Hello $getuserhello, you are through to Littlefish. How can I help today?"
    
        $morninghello | Set-Clipboard

    }

    function write-livechat {
        #function to create email request to start a live chat
        $usersr = $UsernameInput.Text 
        $startlivechat = "I hope you are well today! Can you please start a Littlefish Live Chat (bright orange icon on your desktop called 'Littlefish service desk') and one of our engineers will be happy to assist you! Please quote your ticket number '$usersr' to the engineer when you connect.

Many Thanks
        


-Littlefish"

        $startlivechat | Set-Clipboard 

    }

    function Write-D365Gen {
        #function to generate D365 incoming tickets
        $d35ticket = "I hope you are well today! Can you please just confirm the answers to these questions for me so we can get this resolved for you! 

1) What Legal Entity are you working on 2) What workspace or Module are you attempting to sign into 3) What is the ID of the D365 item you are working on 4) Can you please provide me a screenshot of the error you are getting
        
Many Thanks
-Littlefish"

            $d35ticket | Set-Clipboard
    }

    function write-fullticket {
    
        $issue = $issuebody.text
        $stepstaken = $stepstakenbody.text
        $vwm = $vwmbody.text
        $contact = $contactbody.text
        $location = $locationbody.text
        $Ticket = $TicketTitle.text
    
        $finalticket = "Problem:
$issue
    
Steps taken:
$stepstaken

Asset Number: $vwm
Contact: $contact
Location: $location
"
    
    $finalticket | Set-Clipboard
    new-item -path C:\tickets\ -name "$ticket.txt" -ItemType "File" -Value "$finalticket"
    
    $issuebody.text = ""
    $stepstakenbody.text = ""
    $locationbody.text = ""
    $vwmbody.text = ""
    $contactbody.text = ""
    $TicketTitle.text = ""
    
    
    }
    
    function test-user {
    
        $a = new-object -comobject wscript.shell
    $intAnswer = $a.popup("Do you need to add extra information?", `
    0,"Extra Information",4)
    If ($intAnswer -eq 6) {
      write-extra
    } else {
      write-fullticket
    }
    }
    function write-extraticket {
        
        $issue = $issuebody.text
        $stepstaken = $stepstakenbody.text
        $vwm = $vwmbody.text
        $contact = $contactbody.text
        $location = $locationbody.text
        $Ticket = $TicketTitle.text
        $extrainfo = $extrabody.text 
    
        $finalticket = "Problem:
$issue
    
Steps taken:
$stepstaken
    
Extra Information:
$extrainfo
    
VWM: $vwm
Contact: $contact
Location: $location
"
    
    $finalticket | Set-Clipboard
    new-item -path C:\tickets\ -name "$ticket.txt" -ItemType "File" -Value "$finalticket"
    
    $issuebody.text = ""
    $stepstakenbody.text = ""
    $locationbody.text = ""
    $vwmbody.text = ""
    $contactbody.text = ""
    $TicketTitle.text = ""
    $extrabody.text = ""
    
    
    }
    Function write-extra {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.Application]::EnableVisualStyles()
        
        $ExtraForm                            = New-Object system.Windows.Forms.Form
        $ExtraForm.ClientSize                 = New-Object System.Drawing.Point(400,400)
        $ExtraForm.text                       = "Extra information"
        $ExtraForm.TopMost                    = $false
        $ExtraForm.AutoSize                   = $true
        
        $extrabody                            = New-Object system.Windows.Forms.TextBox
        $extrabody.multiline                  = $True
        $extrabody.width                      = 400
        $extrabody.height                     = 400
        $extrabody.location                   = New-Object System.Drawing.Point(1,40)
        $extrabody.Font                       = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    
        $extrabodylabel                       = New-Object system.Windows.Forms.Label
        $extrabodylabel.text                  = "Enter Extra Info Below"
        $extrabodylabel.AutoSize              = $true
        $extrabodylabel.width                 = 25
        $extrabodylabel.height                = 10
        $extrabodylabel.location              = New-Object System.Drawing.Point(150,20)
        $extrabodylabel.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
        
        $btnRunFunction                       = New-Object system.Windows.Forms.Button
        $btnRunFunction.text                  = "Run"
        $btnRunFunction.width                 = 60
        $btnRunFunction.height                = 30
        $btnRunFunction.location              = New-Object System.Drawing.Point(160,471)
        $btnRunFunction.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    

        
        
        $ExtraForm.controls.AddRange(@($extrabody,$extrabodylabel,$btnRunFunction,$btnExit))
        
        $btnRunFunction.Add_Click({write-extraticket $ExtraForm.Close()})
    
        [void]$ExtraForm.ShowDialog()
    
        
    
    }
    

########################Locations selection dropdown#############################################
$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(10, 350)
$comboBox1.Size = New-Object System.Drawing.Size(350, 310)
$comboBox1.Text = "Select Locations"
$comboBox1.Items.add("Ardley EFW")
$comboBox1.Items.add("Ardley-Landfill")
$comboBox1.Items.add("Atherton")
$comboBox1.Items.add("Avonmouth ERF")
$comboBox1.Items.add("Bargeddie Logistics")
$comboBox1.Items.add("Bargeddie MPT")
$comboBox1.Items.add("Bargeddie Weighbridge FTTC")
$comboBox1.Items.add("Barnstaple")
$comboBox1.Items.add("Beddingham-Power-Plant")
$comboBox1.Items.add("Beddington-ERF")
$comboBox1.Items.add("Beddington-Landfill")
$comboBox1.Items.add("Billingshurst - HWRC")
$comboBox1.Items.add("Blantyre")
$comboBox1.Items.add("Bognor Regis - HWRC")
$comboBox1.Items.add("Bonnyrigg")
$comboBox1.Items.add("Broadpath")
$comboBox1.Items.add("Burgess-Hill")
$comboBox1.Items.add("Calne Landfill")
$comboBox1.Items.add("Chard-HWRC")
$comboBox1.Items.add("Cheddar-HWRC")
$comboBox1.Items.add("Chelson-Meadow-MRF")
$comboBox1.Items.add("Corby")
$comboBox1.Items.add("Crawley")
$comboBox1.Items.add("Crayford")
$comboBox1.Items.add("Crewkerne HWRC")
$comboBox1.Items.add("Crossness")
$comboBox1.Items.add("Dimmer")
$comboBox1.Items.add("Dorset Service Centre (Parkstone)")
$comboBox1.Items.add("Dulcote-HWRC")
$comboBox1.Items.add("Dulverton HWRC")
$comboBox1.Items.add("Dunbar - Landfill")
$comboBox1.Items.add("Dunbar-ERF")
$comboBox1.Items.add("Earls-Barton")
$comboBox1.Items.add("East Kilbride HWRC")
$comboBox1.Items.add("East-Grinstead")
$comboBox1.Items.add("Elsenham LF")
$comboBox1.Items.add("Elsenham-Power-Plant")
$comboBox1.Items.add("Erin-Landfill")
$comboBox1.Items.add("Exeter-Logistics-Unit")
$comboBox1.Items.add("Filton")
$comboBox1.Items.add("Five-Lanes")
$comboBox1.Items.add("Ford-MRF")
$comboBox1.Items.add("Foxhall-Landfill")
$comboBox1.Items.add("Frome-HWRC")
$comboBox1.Items.add("Glasgow-RREC Stores")
$comboBox1.Items.add("Greenhags Transfer Station")
$comboBox1.Items.add("Heathfield-Landfill")
$comboBox1.Items.add("Hersden-Depot")
$comboBox1.Items.add("Highbridge-HWRC")
$comboBox1.Items.add("Hinkley Point C")
$comboBox1.Items.add("Horsham - HWRC")
$comboBox1.Items.add("Horton-Landfill")
$comboBox1.Items.add("Iver South")
$comboBox1.Items.add("Lackford-Landfill")
$comboBox1.Items.add("Lakeside-EFW")
$comboBox1.Items.add("Lancing RTS")
$comboBox1.Items.add("Larkhall HWRC")
$comboBox1.Items.add("Lean Quarry Landfill")
$comboBox1.Items.add("Linwood Transfer Station")
$comboBox1.Items.add("Littlehampton - HWRC")
$comboBox1.Items.add("Llanfoist-Transfer-Station")
$comboBox1.Items.add("Maple Lodge")
$comboBox1.Items.add("Masons Landfill and MRF")
$comboBox1.Items.add("Mavis Valley ATS")
$comboBox1.Items.add("Midhurst - HWRC")
$comboBox1.Items.add("Milton-Keynes")
$comboBox1.Items.add("Minehead-HWRC")
$comboBox1.Items.add("Newhouse")
$comboBox1.Items.add("Parkwood Landfill")
$comboBox1.Items.add("Peninsula House")
$comboBox1.Items.add("Perth")
$comboBox1.Items.add("Peterborough-ERF")
$comboBox1.Items.add("Peterborough-Glass-Weighbridge")
$comboBox1.Items.add("Pilsworth-Power-Plant")
$comboBox1.Items.add("Pilsworth-South")
$comboBox1.Items.add("Plymouth-Depot (Plympton)")
$comboBox1.Items.add("Polmadie Glasgow-RREC")
$comboBox1.Items.add("Priorswood")
$comboBox1.Items.add("Rigmuir (LF &amp; PPlant)")
$comboBox1.Items.add("Riverside-House")
$comboBox1.Items.add("Rochester")
$comboBox1.Items.add("Rochester PRF")
$comboBox1.Items.add("Runcorn EFW")
$comboBox1.Items.add("Rutherglen")
$comboBox1.Items.add("Salmon-Pastures")
$comboBox1.Items.add("Saltlands-HWRC")
$comboBox1.Items.add("Shelford-Landfill-Offices")
$comboBox1.Items.add("Shelford-Landfill-Weighbridge")
$comboBox1.Items.add("Shoreham - HWRC")
$comboBox1.Items.add("Singlerose (St Austell)")
$comboBox1.Items.add("Skelmersdale-Recycling-Plant")
$comboBox1.Items.add("Slough")
$comboBox1.Items.add("Somerton-HWRC")
$comboBox1.Items.add("South West Communications")
$comboBox1.Items.add("South-Ockendon")
$comboBox1.Items.add("Squabb-Wood")
$comboBox1.Items.add("St Helens Electrical Recycling (WEEE)")
$comboBox1.Items.add("Strathaven - HWRC")
$comboBox1.Items.add("Street HWRC")
$comboBox1.Items.add("Tatchells (Wareham)")
$comboBox1.Items.add("Thetford")
$comboBox1.Items.add("Trafford Park")
$comboBox1.Items.add("Trident ERF (Cardiff ERF)")
$comboBox1.Items.add("Trigon-Landfill")
$comboBox1.Items.add("Villiers-Road (Kingston Transfer Station)")
$comboBox1.Items.add("Viridor-House")
$comboBox1.Items.add("Walpole")
$comboBox1.Items.add("Wangford-Landfill")
$comboBox1.Items.add("Warmwell")
$comboBox1.Items.add("Warth Road (Bury MRF)")
$comboBox1.Items.add("Wellington (Poole Landfill Office)")
$comboBox1.Items.add("Westbury")
$comboBox1.Items.add("Westhampnett")
$comboBox1.Items.add("West-Thurrock (Thurrock)")
$comboBox1.Items.add("Willition-HWRC")
$comboBox1.Items.add("Wootton-Landfill")
$comboBox1.Items.add("Worthing - HWRC")
$comboBox1.Items.add("Yanley-Power-Plant")
$comboBox1.Items.add("Yeovil-HWRC")



$MainForm.Controls.Add($comboBox1)

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Point(10, 380)
$Button.Size = New-Object System.Drawing.Size(98, 23)
$Button.Text = "Output"
$Button.add_Click({test-Location})
$MainForm.Controls.Add($Button)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(10, 410)
$outputLabel.Size = New-Object System.Drawing.Size(98, 23)
$outputLabel.Text = ""

$outputLabeler = New-Object System.Windows.Forms.Label
$outputLabeler.Location = New-Object System.Drawing.Point(10, 410)
$outputLabeler.Size = New-Object System.Drawing.Size(98, 23)
$outputLabeler.Text = "Team:"

$MainForm.Controls.Add($outputLabel)
$mainform.controls.add($outputlabeler)



$labelertitle                           = New-Object system.Windows.Forms.Label
$labelertitle.text                      = "Virior Regional Location Selector"
$labelertitle.AutoSize                  = $true
$labelertitle.width                     = 25
$labelertitle.height                    = 10
$labelertitle.location                  = New-Object System.Drawing.Point(15,310)
$labelertitle.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',15)
$mainform.Controls.add($labelertitle)

#functions

function test-Location {

    if($combobox1.Text -contains 'Ardley EFW') {$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Ardley-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Atherton'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Avonmouth ERF'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Bargeddie Logistics'){$outputLabel.text = 'North'}
    elseif($combobox1.text -contains 'Bargeddie MPT'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Bargeddie Weighbridge FTTC'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Barnstaple'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Beddingham-Power-Plant'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Beddington-ERF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Beddington-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Billingshurst - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Blantyre'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Bognor Regis - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Bonnyrigg'){$outputLabel.text = 'North'}
    elseif($combobox1.text -contains 'Broadpath'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Burgess-Hill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Calne Landfill'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Chard-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Cheddar-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Chelson-Meadow-MRF'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Corby'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Crawley'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Crayford'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Crewkerne HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Crossness'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Dimmer'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Dorset Service Centre (Parkstone)'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Dulcote-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Dulverton HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Dunbar - Landfill'){$outputLabel.text = 'North'}
    elseif($combobox1.text -contains 'Dunbar-ERF'){$outputLabel.text = 'North'}
    elseif($combobox1.text -contains 'Earls-Barton'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'East Kilbride HWRC'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'East-Grinstead'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Elsenham LF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Elsenham-Power-Plant'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Erin-Landfill'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Exeter-Logistics-Unit'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Filton'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Five-Lanes'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Ford-MRF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Foxhall-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Frome-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Glasgow-RREC Stores'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Greenhags Transfer Station'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Heathfield-Landfill'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Hersden-Depot'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Highbridge-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Hinkley Point C'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Horsham - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Horton-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Iver South'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Lackford-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Lakeside-EFW'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Lancing RTS'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Larkhall HWRC'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Lean Quarry Landfill'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Linwood Transfer Station'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Littlehampton - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Llanfoist-Transfer-Station'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Maple Lodge'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Masons Landfill and MRF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Mavis Valley ATS'){$outputLabel.text = 'North'}
    elseif($combobox1.text -contains 'Midhurst - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Milton-Keynes'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Minehead-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Newhouse'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Parkwood Landfill'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Peninsula House'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Perth'){$outputLabel.text = 'North'}
    elseif($combobox1.text -contains 'Peterborough-ERF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Peterborough-Glass-Weighbridge'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Pilsworth-Power-Plant'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Pilsworth-South'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Plymouth-Depot (Plympton)'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Polmadie Glasgow-RREC'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Priorswood'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Rigmuir (LF &amp; PPlant)'){$outputLabel.text = 'North'}
    elseif($combobox1.text -contains 'Riverside-House'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Rochester'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Rochester PRF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Runcorn EFW'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Rutherglen'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Salmon-Pastures'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Saltlands-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Shelford-Landfill-Offices'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Shelford-Landfill-Weighbridge'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Shoreham - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Singlerose (St Austell)'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Skelmersdale-Recycling-Plant'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Slough'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Somerton-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'South West Communications'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'South-Ockendon'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Squabb-Wood'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'St Helens Electrical Recycling (WEEE)'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Strathaven - HWRC'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Street HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Tatchells (Wareham)'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Thetford'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Trafford Park'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Trident ERF (Cardiff ERF)'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Trigon-Landfill'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Villiers-Road (Kingston Transfer Station)'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Viridor-House'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Walpole'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Wangford-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Warmwell'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Warth Road (Bury MRF)'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Wellington (Poole Landfill Office)'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Westbury'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Westhampnett'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'West-Thurrock (Thurrock)'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Willition-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Wootton-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Worthing - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Yanley-Power-Plant'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Yeovil-HWRC'){$outputLabel.text = "South West"}
    


}
    
Function Start-IncomingTicket{
#Opening new window with incoming call ticket input form
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    # Creating the form
    $IncomingCall                            = New-Object system.Windows.Forms.Form
    $IncomingCall.ClientSize                 = New-Object System.Drawing.Point(600,600)
    $IncomingCall.text                       = "Incoming Call Gen"
    $IncomingCall.TopMost                    = $true
    $IncomingCall.AutoSize                   = $true
    
    ############################# Incoming Call Ticket Generator ##################################
    
    $Title                           = New-Object system.Windows.Forms.Label
    $Title.text                      = "Incoming call Ticket Generator"
    $Title.AutoSize                  = $true
    $Title.width                     = 25
    $Title.height                    = 8
    $Title.location                  = New-Object System.Drawing.Point(90,1)
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
    
    #User Name label
    $ICUserLabel                  = New-Object system.Windows.Forms.Label
    $ICUserLabel.text             = "Users Name"
    $ICUserLabel.AutoSize         = $true
    $ICUserLabel.width            = 25
    $ICUserLabel.height           = 10
    $ICUserLabel.location         = New-Object System.Drawing.Point(10,40)
    $ICUserLabel.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    
    #User Name input box
    $ICUserVar                   = New-Object system.Windows.Forms.TextBox
    $ICUserVar.multiline         = $false
    $ICUserVar.width             = 60
    $ICUserVar.height            = 50
    $ICUserVar.location          = New-Object System.Drawing.Point(150,38)
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


    
[void]$MainForm.ShowDialog()