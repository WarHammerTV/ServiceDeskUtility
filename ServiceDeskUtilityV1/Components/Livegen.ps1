<#
Viridor Response Generator v2.0
W4rH4mm3r
#>

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(400,400)
$Form.text                       = "Viridor Response Generator"
$Form.TopMost                    = $false
$Form.AutoSize                   = $true

$Title                           = New-Object system.Windows.Forms.Label
$Title.text                      = "Viridor Response Generator"
$Title.AutoSize                  = $true
$Title.width                     = 25
$Title.height                    = 10
$Title.location                  = New-Object System.Drawing.Point(20,12)
$Title.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',25)

$btnExcelResolve                  = New-Object system.Windows.Forms.Button
$btnExcelResolve.text             = "Excel Resolve"
$btnExcelResolve.width            = 70
$btnExcelResolve.height           = 50
$btnExcelResolve.location         = New-Object System.Drawing.Point(20,100)
$btnExcelResolve.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnDatasetFunction                  = New-Object system.Windows.Forms.Button
$btnDatasetFunction.text             = "Dataset Request"
$btnDatasetFunction.width            = 70
$btnDatasetFunction.height           = 50
$btnDatasetFunction.location         = New-Object System.Drawing.Point(100,100)
$btnDatasetFunction.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnStdResolve                  = New-Object system.Windows.Forms.Button
$btnStdResolve.text             = "Standard Resolve"
$btnStdResolve.width            = 70
$btnStdResolve.height           = 50
$btnStdResolve.location         = New-Object System.Drawing.Point(180,100)
$btnStdResolve.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnCitrix                  = New-Object system.Windows.Forms.Button
$btnCitrix.text             = "Citrix Resolve"
$btnCitrix.width            = 70
$btnCitrix.height           = 50
$btnCitrix.location         = New-Object System.Drawing.Point(260,100)
$btnCitrix.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnHello                  = New-Object system.Windows.Forms.Button
$btnHello.text             = "Hello Gen"
$btnHello.width            = 70
$btnHello.height           = 50
$btnHello.location         = New-Object System.Drawing.Point(340,100)
$btnHello.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btnLiveChat                  = New-Object system.Windows.Forms.Button
$btnLiveChat.text             = "Live Chat"
$btnLiveChat.width            = 70
$btnLiveChat.height           = 50
$btnLiveChat.location         = New-Object System.Drawing.Point(420,100)
$btnLiveChat.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$UsernameInput                        = New-Object system.Windows.Forms.TextBox
$UsernameInput.multiline               = $false
$UsernameInput.width                   = 100
$UsernameInput.height                  = 20
$UsernameInput.location                = New-Object System.Drawing.Point(200,160)
$UsernameInput.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$UsernameLabel                          = New-Object system.Windows.Forms.Label
$UsernameLabel.text                     = "Enter User's Name"
$UsernameLabel.AutoSize                 = $true
$UsernameLabel.width                    = 25
$UsernameLabel.height                   = 10
$UsernameLabel.location                 = New-Object System.Drawing.Point(80,160)
$UsernameLabel.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)


$Form.controls.AddRange(@($Title,$btnExcelResolve,$btnDatasetFunction,$btnStdResolve,$btnCitrix,$UsernameInput,$UsernameLabel,$btnHello,$btnLiveChat))

$btnExcelResolve.Add_Click({Write-ExcelResponse})
$btnDatasetFunction.Add_Click({Write-Dataset})
$btnStdResolve.Add_click({write-stdresponse})
$btnCitrix.Add_click({write-citrixsolve})
$btnHello.Add_Click({write-hello})
$btnLiveChat.Add_click({write-livechat})

## Functions ##

function Write-Dataset { 
    $datasetfinal = "Can you please just confirm a few details for me: 1)The VWM number of your device 2) The best contact number for yourself 3) Your current location"

    $datasetfinal | Set-Clipboard
    
}

function Write-ExcelResponse {

    $excelresponseuser = $UsernameInput.Text
    $excelresponsefinal = "Signed $excelresponseuser out of excel and back in. Reopening the report from gatehouse and error resolved"

    $excelresponsefinal | Set-Clipboard 

}

function write-stdresponse {

    $stdresponseuser = $UsernameInput.Text

    $stdresponsefinal = "Hello $stdresponseuser,
Any further issues please don't hesitate  to get back in contact and one of our engineers will be happy to assist you further!
    
Many Thanks
-Littlefish"

    $stdresponsefinal | Set-Clipboard 
} 

function write-citrixsolve {
    $citrixfinal = "Ended all citrix tasks in task manager and restarted citrix workspace. User signed back in and restarted gatehouse program"
    $citrixfinal | Set-Clipboard
    
    }

    function write-hello {

        $getuserhello = $UsernameInput.Text
        $morninghello = "Hello $getuserhello, you are through to Littlefish. How can I help today?"
    
        $morninghello | Set-Clipboard

    }

    function write-livechat {
        $usersr = $UsernameInput.Text 
        $startlivechat = "I hope you are well today! Can you please start a Littlefish Live Chat (bright orange icon on your desktop called 'Littlefish service desk') and one of our engineers will be happy to assist you! Please quote your ticket number '$usersr' to the engineer when you connect.

Many Thanks
        
-Littlefish"

        $startlivechat | Set-Clipboard 

    }





[void]$Form.ShowDialog()