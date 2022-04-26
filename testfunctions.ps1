[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$Form1 = New-Object System.Windows.Forms.Form
$Form1.ClientSize = New-Object System.Drawing.Size(407, 390)
$form1.topmost = $true



$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(25, 55)
$comboBox1.Size = New-Object System.Drawing.Size(350, 310)
$comboBox1.Text = "Select Locations"
$comboBox1.Items.add("Ardley EFW") 
$comboBox1.Items.add("Ardley-Landfill")
$comboBox1.Items.add("Atherton")
$comboBox1.Items.Add("Avonmouth ERF")
$comboBox1.Items.Add("Bargeddie Logistics")
$comboBox1.Items.Add("Bargeddie MPT")
$comboBox1.Items.Add("Bargeddie Weighbridge FTTC")
$comboBox1.Items.Add("Barnstaple")
$comboBox1.Items.Add("Beddingham-Power-Plant")
$comboBox1.Items.Add("Beddington-ERF")
$comboBox1.Items.Add("Beddington-Landfill")
$comboBox1.Items.Add("Billingshurst - HWRC")
$comboBox1.Items.Add("Blantyre")
$comboBox1.Items.Add("Bognor Regis - HWRC")
$comboBox1.Items.Add("Bonnyrigg")
$comboBox1.Items.Add("Broadpath")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")
$comboBox1.Items.Add("")

$Form1.Controls.Add($comboBox1)

$Button = New-Object System.Windows.Forms.Button
$Button.Location = New-Object System.Drawing.Point(25, 20)
$Button.Size = New-Object System.Drawing.Size(98, 23)
$Button.Text = "Output"
$Button.add_Click({Test-Location})
$Form1.Controls.Add($Button)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(70, 90)
$label.Size = New-Object System.Drawing.Size(98, 23)
$label.Text = "1"
$Form1.Controls.Add($label)

#functions

function Test-Location {

    if($combobox1.Text -contains 'Ardley EFW') {$label.text = "South East"}
    elseif($combobox1.text -contains 'Ardley-Landfill'){$label.Text = "South East"}
    elseif($combobox1.text -contains 'Atherton'){$label.text = "North"}
    elseif($combobox1.text -contains 'Avonmouth ERF'){$label.text = "South West"}
    elseif($combobox1.text -contains 'Bargeddie Logistics'){$label.text = "North"}
    elseif($combobox1.text -contains 'Bargeddie MPT'){$label.text = "North"}
    elseif($combobox1.text -contains 'Bargeddie Weighbridge FTTC'){$label.text = "North"}
    elseif($combobox1.text -contains 'Barnstaple'){$label.text = "South West"}
    elseif($combobox1.text -contains 'Beddingham-Power-Plant'){$label.text = "South East"}
    elseif($combobox1.text -contains 'Beddington-ERF'){$label.text = "South East"}
    elseif($combobox1.text -contains 'Beddington-Landfill'){$label.text = "South East"}
    elseif($combobox1.text -contains 'Billingshurst - HWRC'){$label.text = "South East"}
    elseif($combobox1.text -contains 'Blantyre'){$label.text = "North"}
    elseif($combobox1.text -contains 'Bognor Regis - HWRC'){$label.text = "South East"}
    elseif($combobox1.text -contains 'Bonnyrigg'){$label.text = "North"}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.text -contains ''){$label.text = ""}
    elseif($combobox1.test -contains ''){$label.text = ""}
}


[void]$form1.showdialog()