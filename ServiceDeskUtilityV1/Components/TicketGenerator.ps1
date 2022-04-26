$path = "C:\tickets"

if(test-path -path $path){
"Target Path Exists"}
Else{new-item -path C:\ -name "tickets" -ItemType Directory}
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()


$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(400,400)
$Form.text                       = "Viridor Ticket Generator"
$Form.TopMost                    = $false
$form.AutoSize                   = $true

$Title                           = New-Object system.Windows.Forms.Label
$Title.text                      = "Viridor Ticket Generator"
$Title.AutoSize                  = $true
$Title.width                     = 25
$Title.height                    = 10
$Title.location                  = New-Object System.Drawing.Point(200,12)
$Title.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',25)

$TicketTitle                         = New-Object system.Windows.Forms.TextBox
$TicketTitle.multiline               = $true
$TicketTitle.width                   = 150
$TicketTitle.height                  = 20
$TicketTitle.location                = New-Object System.Drawing.Point(200,95)
$TicketTitle.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TicketTitleLabel                          = New-Object system.Windows.Forms.Label
$TicketTitleLabel.text                     = "Enter Ticket Number"
$TicketTitleLabel.AutoSize                 = $true
$TicketTitleLabel.width                    = 25
$TicketTitleLabel.height                   = 10
$TicketTitleLabel.location                 = New-Object System.Drawing.Point(20,95)
$TicketTitleLabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$issuebody                        = New-Object system.Windows.Forms.TextBox
$issuebody.multiline               = $true
$issuebody.AcceptsReturn           = $true
$issuebody.ScrollBars              = "Vertical" 
$issuebody.width                   = 300
$issuebody.height                  = 200
$issuebody.location                = New-Object System.Drawing.Point(200,120)
$issuebody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$issuebodylabel                          = New-Object system.Windows.Forms.Label
$issuebodylabel.text                     = "Enter Issue description"
$issuebodylabel.AutoSize                 = $true
$issuebodylabel.width                    = 25
$issuebodylabel.height                   = 10
$issuebodylabel.location                 = New-Object System.Drawing.Point(20,120)
$issuebodylabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$stepstakenbody                        = New-Object system.Windows.Forms.TextBox
$stepstakenbody.AcceptsReturn           = $true
$stepstakenbody.ScrollBars              = "Vertical"
$stepstakenbody.multiline               = $true
$stepstakenbody.width                   = 300
$stepstakenbody.height                  = 200
$stepstakenbody.location                = New-Object System.Drawing.Point(200,325)
$stepstakenbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$stepstakenlabel                          = New-Object system.Windows.Forms.Label
$stepstakenlabel.text                     = "Enter Steps Taken"
$stepstakenlabel.AutoSize                 = $true
$stepstakenlabel.width                    = 25
$stepstakenlabel.height                   = 10
$stepstakenlabel.location                 = New-Object System.Drawing.Point(20,325)
$stepstakenlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$vwmbody                        = New-Object system.Windows.Forms.TextBox
$vwmbody.multiline               = $false
$vwmbody.width                   = 100
$vwmbody.height                  = 20
$vwmbody.location                = New-Object System.Drawing.Point(680,95)
$vwmbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$vwmlabel                          = New-Object system.Windows.Forms.Label
$vwmlabel.text                     = "Please enter VWM Number"
$vwmlabel.AutoSize                 = $true
$vwmlabel.width                    = 25
$vwmlabel.height                   = 10
$vwmlabel.location                 = New-Object System.Drawing.Point(520,95)
$vwmlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$contactbody                        = New-Object system.Windows.Forms.TextBox
$contactbody.multiline               = $false
$contactbody.width                   = 100
$contactbody.height                  = 20
$contactbody.location                = New-Object System.Drawing.Point(680,115)
$contactbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$contactlabel                          = New-Object system.Windows.Forms.Label
$contactlabel.text                     = "Please enter Contact No."
$contactlabel.AutoSize                 = $true
$contactlabel.width                    = 25
$contactlabel.height                   = 10
$contactlabel.location                 = New-Object System.Drawing.Point(520,115)
$contactlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$locationbody                        = New-Object system.Windows.Forms.TextBox
$locationbody.multiline               = $false
$locationbody.width                   = 100
$locationbody.height                  = 20
$locationbody.location                = New-Object System.Drawing.Point(680,135)
$locationbody.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$locationlabel                          = New-Object system.Windows.Forms.Label
$locationlabel.text                     = "Please enter users location"
$locationlabel.AutoSize                 = $true
$locationlabel.width                    = 25
$locationlabel.height                   = 10
$locationlabel.location                 = New-Object System.Drawing.Point(520,135)
$locationlabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnRunFunction                  = New-Object system.Windows.Forms.Button
$btnRunFunction.text             = "Run"
$btnRunFunction.width            = 60
$btnRunFunction.height           = 30
$btnRunFunction.location         = New-Object System.Drawing.Point(630,180)
$btnRunFunction.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnDoExit                       = New-Object System.Windows.Forms.Button
$btnDoExit.text                  = "Exit"
$btnDoExit.Width                 = 60
$btnDoExit.Height                = 30
$btnDoExit.Location              = New-Object System.Drawing.Point (680,180)


$Form.controls.AddRange(@($Title,$TicketTitle,$TicketTitleLabel,$issuebody,$issuebodylabel,$stepstakenbody,$stepstakenlabel,$vwmbody,$vwmlabel,$contactbody,$contactlabel,$locationbody,$locationlabel,$btnRunFunction, $btnDoExit))

$btnRunFunction.Add_Click({test-user})
$btnDoExit.Add_Click({$Form.Close()})

## Enter functions here ##

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

    $btnExit                              = New-Object system.Windows.Forms.Button
    $btnExit.text                         = "Exit"
    $btnExit.width                        = 60
    $btnExit.height                       = 30
    $btnExit.location                     = New-Object System.Drawing.Point(220,471)
    $btnExit.Font                         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
    
    
    $ExtraForm.controls.AddRange(@($extrabody,$extrabodylabel,$btnRunFunction,$btnExit))
    
    $btnRunFunction.Add_Click({write-extraticket})
    $btnExit.Add_Click({$ExtraForm.Close()})

    [void]$ExtraForm.ShowDialog()

    

}




    


[void]$Form.ShowDialog()