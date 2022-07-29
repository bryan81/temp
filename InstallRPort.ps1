#import needed framework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


#create form
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Client Info'
$form.Size = New-Object System.Drawing.Size(300,330)
$form.StartPosition = 'CenterScreen'

#create OK button
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,250)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

#create cancel button
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,250)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = 'Cancel'
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

#text input setup

$textLabel = New-Object System.Windows.Forms.Label
$textLabel.Location = New-Object System.Drawing.Point(10,20)
$textLabel.Size = New-Object System.Drawing.Size(280,20)
$textLabel.Text = 'Please enter the client key:'
$form.Controls.Add($textLabel)

$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,40)
$textBox.Size = New-Object System.Drawing.Size(260,20)
$form.Controls.Add($textBox)


#Multi Select List Setup

$listLabel = New-Object System.Windows.Forms.Label
$listLabel.Location = New-Object System.Drawing.Point(10,70)
$listLabel.Size = New-Object System.Drawing.Size(280,20)
$listLabel.Text = 'Please select tags from the list below:'
$form.Controls.Add($listLabel)

$listBox = New-Object System.Windows.Forms.Listbox
$listBox.Location = New-Object System.Drawing.Point(10,90)
$listBox.Size = New-Object System.Drawing.Size(260,20)

$listBox.SelectionMode = 'MultiExtended'

#setup list options
[void] $listBox.Items.Add('Office')
[void] $listBox.Items.Add('Mosaic')
[void] $listBox.Items.Add('Shelter')
[void] $listBox.Items.Add('Desktop')
[void] $listBox.Items.Add('Laptop')
[void] $listBox.Items.Add('Tablet')
[void] $listBox.Items.Add('Client')
[void] $listBox.Items.Add('Staff')
[void] $listBox.Items.Add('IT')
[void] $listBox.Items.Add('Finance')
[void] $listBox.Items.Add('Management')
[void] $listBox.Items.Add('Leadership')

#add list to dialog box
$listBox.Height = 150
$form.Controls.Add($listBox)
$form.Topmost = $true

#show the dialog box
$result = $form.ShowDialog()

#start tag string
$key = 'tags = ['

#assign dialog box results to variable and build the tag string
if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    #text capture
    $id = $textBox.Text

    #list selections capture
    $items = $listBox.SelectedItems
    foreach ($item in $items) {
        $key += ",'$item'"
    }
}

#finish the needed string
$key+= "]"

#remove the extraneous comma
[regex]$pattern = ","
$key=$pattern.replace($key, "", 1) 

#download client software
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$url="https://pairing.rport.io/$id"
Invoke-WebRequest -Uri $url -OutFile "rport-installer.ps1"
powershell -ExecutionPolicy Bypass -File .\rport-installer.ps1 -x -r

#add the correct tags to the Client
((Get-Content -path "C:\Program Files\rport\rport.conf" -Raw) -replace "tags = \[.*\]", $key) | Set-Content -path "C:\Program Files\rport\rport.conf"

#Restart the rport service to reflect changes
restart-service rport