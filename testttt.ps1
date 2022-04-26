$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
add-type -name win -member $t -namespace native
[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$MainForm= New-Object System.Windows.Forms.Form
$MainForm.ClientSize = New-Object System.Drawing.Size(407, 390)
$MainForm.topmost = $true



$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(25, 55)
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
$Button.Location = New-Object System.Drawing.Point(25, 20)
$Button.Size = New-Object System.Drawing.Size(98, 23)
$Button.Text = "Output"
$Button.add_Click({test-Location})
$MainForm.Controls.Add($Button)

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(25, 90)
$outputLabel.Size = New-Object System.Drawing.Size(98, 23)
$outputLabel.Text = ""
$MainForm.Controls.Add($outputLabel)

#functions

function test-Location {

    if($combobox1.Text -contains 'Ardley EFW') {$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Ardley-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Atherton'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Avonmouth ERF'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Bargeddie Logistics'){$outputLabel.text = 'Scotland'}
    elseif($combobox1.text -contains 'Bargeddie MPT'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Bargeddie Weighbridge FTTC'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Barnstaple'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Beddingham-Power-Plant'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Beddington-ERF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Beddington-Landfill'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Billingshurst - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Blantyre'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Bognor Regis - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Bonnyrigg'){$outputLabel.text = 'Scotland'}
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
    elseif($combobox1.text -contains 'Dunbar - Landfill'){$outputLabel.text = 'Scotland'}
    elseif($combobox1.text -contains 'Dunbar-ERF'){$outputLabel.text = 'Scotland'}
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
    elseif($combobox1.text -contains 'Mavis Valley ATS'){$outputLabel.text = 'Scotland'}
    elseif($combobox1.text -contains 'Midhurst - HWRC'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Milton-Keynes'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Minehead-HWRC'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Newhouse'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Parkwood Landfill'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Peninsula House'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Perth'){$outputLabel.text = 'Scotland'}
    elseif($combobox1.text -contains 'Peterborough-ERF'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Peterborough-Glass-Weighbridge'){$outputLabel.text = "South East"}
    elseif($combobox1.text -contains 'Pilsworth-Power-Plant'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Pilsworth-South'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Plymouth-Depot (Plympton)'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Polmadie Glasgow-RREC'){$outputLabel.text = "North"}
    elseif($combobox1.text -contains 'Priorswood'){$outputLabel.text = "South West"}
    elseif($combobox1.text -contains 'Rigmuir (LF &amp; PPlant)'){$outputLabel.text = 'Scotland'}
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


[void]$MainForm.showdialog()