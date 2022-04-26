Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$SrGen                            = New-Object system.Windows.Forms.Form
$SrGen.ClientSize                 = New-Object System.Drawing.Point(400,400)
$SrGen.text                       = "Form"
$SrGen.TopMost                    = $false

$TTLJanesGen                           = New-Object system.Windows.Forms.Label
$TTLJanesGen.text                      = "Janes Generator "
$TTLJanesGen.AutoSize                  = $true
$TTLJanesGen.width                     = 25
$TTLJanesGen.height                    = 10
$TTLJanesGen.location                  = New-Object System.Drawing.Point(60,10)
$TTLJanesGen.Font                      = New-Object System.Drawing.Font('Microsoft Sans Serif',25)


$BtnCstmSoftware                  = New-Object system.Windows.Forms.Button
$BtnCstmSoftware.text             = "Run"
$BtnCstmSoftware.width            = 60
$BtnCstmSoftware.height           = 30
$BtnCstmSoftware.location         = New-Object System.Drawing.Point(10,70)
$BtnCstmSoftware.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BtnStdrdSoftware                  = New-Object system.Windows.Forms.Button
$BtnStdrdSoftware.text             = "Run"
$BtnStdrdSoftware.width            = 60
$BtnStdrdSoftware.height           = 30
$BtnStdrdSoftware.location         = New-Object System.Drawing.Point(157,171)
$BtnStdrdSoftware.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BtnDistribution                  = New-Object system.Windows.Forms.Button
$BtnDistribution.text             = "Run"
$BtnDistribution.width            = 60
$BtnDistribution.height           = 30
$BtnDistribution.location         = New-Object System.Drawing.Point(157,171)
$BtnDistribution.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BtnIboss                  = New-Object system.Windows.Forms.Button
$BtnIboss.text             = "Run"
$BtnIboss.width            = 60
$BtnIboss.height           = 30
$BtnIboss.location         = New-Object System.Drawing.Point(157,171)
$BtnIboss.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BtnLocalAdmin                  = New-Object system.Windows.Forms.Button
$BtnLocalAdmin.text             = "Run"
$BtnLocalAdmin.width            = 60
$BtnLocalAdmin.height           = 30
$BtnLocalAdmin.location         = New-Object System.Drawing.Point(157,171)
$BtnLocalAdmin.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$BtnNewStart                  = New-Object system.Windows.Forms.Button
$BtnNewStart.text             = "Run"
$BtnNewStart.width            = 60
$BtnNewStart.height           = 30
$BtnNewStart.location         = New-Object System.Drawing.Point(157,171)
$BtnNewStart.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Output                         = New-Object system.Windows.Forms.TextBox
$Output.multiline               = $true
$Output.width                   = 350
$Output.height                  = 80
$Output.location                = New-Object System.Drawing.Point(30,250)
$Output.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$output.Text                    = $noteoutput.value

$SrGen.controls.AddRange(@($TTLJanesGen,$BtnCstmSoftware,$BtnStdrdSoftware,$BtnDistribution,$BtnIboss,$BtnLocalAdmin,$BtnNewStart))

$BtnCstmSoftware.Add_Click({ <# ButtonFunction #> })

## Enter functions here ##

[void]$SrGen.ShowDialog()