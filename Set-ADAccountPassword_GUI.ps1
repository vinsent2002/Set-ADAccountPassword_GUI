Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#MessageBox
Add-Type -AssemblyName PresentationCore,PresentationFramework

$Title = "Set-ADAccountPassword GUI"

if (!(Get-Module -ListAvailable -Name ActiveDirectory)) {
[System.Windows.MessageBox]::Show("Active Directory Module is not found.",$Title,0,16)
[Environment]::Exit(1)
}

$Domains = "Domain1",
           "Domain2"

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(310,250)
$Form.text                       = $Title
$Form.TopMost                    = $false
$Form.ShowIcon                   = $false
$Form.MaximizeBox                = $false
$Form.MinimizeBox                = $false
$Form.TopMost                    = $True
$Form.FormBorderStyle            = "FixedSingle"
$Form.StartPosition              = "CenterScreen"

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Username:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(5,10)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 150
$TextBox1.height                 = 20
$TextBox1.location               = New-Object System.Drawing.Point(5,30)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Old password:"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(5,60)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$MaskedTextBox1                       = New-Object system.Windows.Forms.MaskedTextBox
$MaskedTextBox1.multiline             = $false
#$MaskedTextBox1.PasswordChar         = '*'
$MaskedTextBox1.UseSystemPasswordChar = $true
$MaskedTextBox1.width                 = 150
$MaskedTextBox1.height                = 20
$MaskedTextBox1.location              = New-Object System.Drawing.Point(5,80)
$MaskedTextBox1.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button2                  = New-Object System.Windows.Forms.Button
$Button2.Text             = "Show"
#$Button2.Top             = 70
#$Button2.Left            = 20
$Button2.AutoSize         = $true
$Button2.Location         = New-Object System.Drawing.Point(160,80)
$Button2.Add_Click({
    if ($MaskedTextBox1.UseSystemPasswordChar) {
        $MaskedTextBox1.UseSystemPasswordChar = $false
        $Button2.Text = "Hide"
    } else {
        $MaskedTextBox1.UseSystemPasswordChar = $true
        $Button2.Text = "Show"
    }
})


$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "New password:"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(5,110)
$Label3.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$MaskedTextBox2                       = New-Object system.Windows.Forms.MaskedTextBox
$MaskedTextBox2.multiline             = $false
#$MaskedTextBox2.PasswordChar         = '*'
$MaskedTextBox2.UseSystemPasswordChar = $true
$MaskedTextBox2.width                 = 150
$MaskedTextBox2.height                = 20
$MaskedTextBox2.location              = New-Object System.Drawing.Point(5,130)
$MaskedTextBox2.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button3                  = New-Object System.Windows.Forms.Button
$Button3.Text             = "Show"
#$Button3.Top             = 70
#$Button3.Left            = 20
$Button3.AutoSize         = $true
$Button3.Location         = New-Object System.Drawing.Point(160,130)
$Button3.Add_Click({
    if ($MaskedTextBox2.UseSystemPasswordChar) {
        $MaskedTextBox2.UseSystemPasswordChar = $false
        $Button3.Text = "Hide"
    } else {
        $MaskedTextBox2.UseSystemPasswordChar = $true
        $Button3.Text = "Show"
    }
})

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Confirm password:"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(5,160)
$Label4.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$MaskedTextBox3                       = New-Object system.Windows.Forms.MaskedTextBox
$MaskedTextBox3.multiline             = $false
#$MaskedTextBox3.PasswordChar         = '*'
$MaskedTextBox3.UseSystemPasswordChar = $true
$MaskedTextBox3.width                 = 150
$MaskedTextBox3.height                = 20
$MaskedTextBox3.location              = New-Object System.Drawing.Point(5,180)
$MaskedTextBox3.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Button4                  = New-Object System.Windows.Forms.Button
$Button4.Text             = "Show"
#$Button4.Top             = 70
#$Button4.Left            = 20
$Button4.AutoSize         = $true
$Button4.Location         = New-Object System.Drawing.Point(160,180)
$Button4.Add_Click({
    if ($MaskedTextBox3.UseSystemPasswordChar) {
        $MaskedTextBox3.UseSystemPasswordChar = $false
        $Button4.Text = "Hide"
    } else {
        $MaskedTextBox3.UseSystemPasswordChar = $true
        $Button4.Text = "Show"
    }
})

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "OK"
$Button1.width                   = 60
$Button1.height                  = 30
$Button1.location                = New-Object System.Drawing.Point(5,210)
$Button1.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Button1.Add_Click({

    if($MaskedTextBox2.Text -eq $MaskedTextBox3.Text){
        $MaskedTextBox2.ReadOnly = $true

        $OldPassSec = ConvertTo-SecureString $MaskedTextBox1.Text -AsPlainText -Force

        $NewPassSec = ConvertTo-SecureString $MaskedTextBox2.Text -AsPlainText -Force

        try {
            Set-ADAccountPassword -Identity $TextBox1.Text -OldPassword $OldPassSec -NewPassword $NewPassSec -Server $ComboBox1.SelectedItem
        } catch {
            [System.Windows.MessageBox]::Show($_,$Title,0,16)
        }
        
        $OldPassSec.Dispose()

        $NewPassSec.Dispose()

        $TextBox1.Text = $null

        $MaskedTextBox2 = $null

        $MaskedTextBox1 = $null

        $MaskedTextBox2 = $null

        $MaskedTextBox2.ReadOnly = $false
    } else {
        $MessageBody = "The password and confirmation password did not match. Please retype the password and confirmation password."
        [System.Windows.MessageBox]::Show($MessageBody,$Title,0,16)
    }

})

$Button5                         = New-Object system.Windows.Forms.Button
$Button5.text                    = "Cancel"
$Button5.width                   = 60
$Button5.height                  = 30
$Button5.location                = New-Object System.Drawing.Point(100,210)
$Button5.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Button5.Add_Click({
    $Form.Close()
})


$ComboBox1                       = New-Object system.Windows.Forms.ComboBox
$ComboBox1.text                  = "comboBox"
$ComboBox1.width                 = 130
$ComboBox1.height                = 20
$ComboBox1.location              = New-Object System.Drawing.Point(160,30)
$ComboBox1.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Domains | foreach {
    $ComboBox1.Items.Add($_)
    $ComboBox1.SelectedIndex = 0
} | Out-Null

#debug
#$ComboBox1.Add_SelectedIndexChanged({
#    $selectedItem = $ComboBox1.SelectedItem
#    Write-Host "Selected item: $selectedItem"
#
#    #[System.Windows.Forms.MessageBox]::Show("Selected item: $selectedItem")
#})


$Form.controls.AddRange(@($TextBox1,$ComboBox1,$MaskedTextBox1,$MaskedTextBox2,$MaskedTextBox3,$Label1,$Label2,$Label3,$Label4,$Button1,$Button2,$Button3,$Button4,$Button5))

#[void]$Form.ShowDialog()
#[System.Windows.Forms.Application]::Run($form)
$form.ShowDialog()

$form.Dispose()
