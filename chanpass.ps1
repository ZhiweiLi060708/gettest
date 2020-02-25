if (-not ([System.Management.Automation.PSTypeName]"TrustAllCertsPolicy").Type)
{
    Add-Type -TypeDefinition  @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem)
    {
        return true;
    }
}
"@
}

if ([System.Net.ServicePointManager]::CertificatePolicy.ToString() -ne "TrustAllCertsPolicy")
{
    [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

while(1){
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Fill in the Box'
    $form.Size = New-Object System.Drawing.Size(300,260)
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point(55,180)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Point(170,180)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(400,20)
    $label.Text = 'IP Address for PVWA server:'
    $form.Controls.Add($label)

    $textBox_ip = New-Object System.Windows.Forms.TextBox
    $textBox_ip.Location = New-Object System.Drawing.Point(10,40)
    $textBox_ip.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($textBox_ip)
    $form.Add_Shown({$textBox_ip.Select()})

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,70)
    $label.Size = New-Object System.Drawing.Size(400,20)
    $label.Text = 'Username for logon:'
    $form.Controls.Add($label)

    $textBox_username = New-Object System.Windows.Forms.TextBox
    $textBox_username.Location = New-Object System.Drawing.Point(10,90)
    $textBox_username.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($textBox_username)
    $form.Add_Shown({$textBox_username.Select()})

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,120)
    $label.Size = New-Object System.Drawing.Size(400,20)
    $label.Text = 'Password for logon:'
    $form.Controls.Add($label)

    $textBox_password = New-Object System.Windows.Forms.TextBox
    $textBox_password.Location = New-Object System.Drawing.Point(10,140)
    $textBox_password.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($textBox_password)
    $form.Add_Shown({$textBox_password.Select()})


    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $ip = $textBox_ip.Text
        $username4logon = $textBox_username.Text
        $password = $textBox_password.Text
    }

    if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
    {
        exit
    }

    Write-Host Testing network connection IP: $ip
    $pingtest = (Test-NetConnection $ip) | ForEach-Object {$_.PingSucceeded}

    if($pingtest){
        $chkpoit = 0
        $headers = @{"accept" = "application/json"}
        $Url = "https://$ip/PasswordVault/API/auth/Cyberark/Logon"
        $Body = @{username = $username4logon;password = $password}
        try{
            $token = Invoke-WebRequest -Method 'Post' -Uri $Url -Body $body -Headers $headers
        }
        catch{
            $err = $_
            $code = $err.Exception.Message
            $message = $err.ErrorDetails.Message
            $ws = New-Object -ComObject WScript.Shell
            $wsr = $ws.popup("Incorrect user or password",-1,"Username or Password Error!",1)
            $chkpoit = 3
        }
        if($chkpoit -eq 0){
            break
        }
    }
    else{
        $ws = New-Object -ComObject WScript.Shell
        $wsr = $ws.popup("$ip Unreachable!",-1,"IP Error!",1)
    }
}


function OpenFile-Dialog($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv) | *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$headers = @{"accept" = "application/json"}
$Url = "https://$ip/PasswordVault/API/auth/Cyberark/Logon"
$Body = @{username = $username4logon;password = $password}
try{
    $token = (Invoke-WebRequest -Method 'Post' -Uri $Url -Body $body -Headers $headers) -Split '"'
}
catch{
    $err = $_
    $code = $err.Exception.Message
    $message = $err.ErrorDetails.Message
    $ws = New-Object -ComObject WScript.Shell
    $wsr = $ws.popup("$message",-1,"$code",1)
    exit
}

$headers = @{"Authorization" = $token[1]}
$Url = "https://$ip/PasswordVault/api/Accounts"

try{
    $test =  Invoke-WebRequest -Method 'Get' -Uri $Url -Headers $headers
    $detail = Invoke-RestMethod -Method 'Get' -Uri $Url -Headers $headers
}
catch{
    $err = $_
    $code = $err.Exception.Message
    $message = $err.ErrorDetails.Message
    $ws = New-Object -ComObject WScript.Shell
    $wsr = $ws.popup("$message",-1,"$code",1)
    exit
}

# Define three variables for the CSV file.
$name = New-Object 'System.Collections.Generic.List[System.Object]'
$username = New-Object 'System.Collections.Generic.List[System.Object]'
$AssetName = New-Object 'System.Collections.Generic.List[System.Object]'
$password = New-Object 'System.Collections.Generic.List[System.Object]'

$filepath = OpenFile-Dialog($Env:CSIDL_DEFAULT_DOWNLOADS)
$csvfile = Import-CSV -Path $filepath

foreach($el in $csvfile){
    $name.Add($el.Password_name)
    $username.Add($el.username)
    $AssetName.Add($el.AssetName)
    $password.Add($el.password)
}

# Query user's ID.
$id = New-Object 'System.Collections.Generic.List[System.Object]'

for($i=0;$i -lt $name.Count;$i++){
    $tmp = $detail | ForEach-Object {$_.value} | ForEach-Object {if($_.platformAccountProperties.AssetName -eq $AssetName[$i] -and $_.name -eq $name[$i]){$_.id}}
    $id.Add($tmp)
}

# ÐÞ¸ÄÃÜÂë
$nochangeid = New-Object 'System.Collections.Generic.List[System.Object]'
$logpath = $filepath + ".log"
(Get-Date) | Out-File -Append $logpath
"============================" | Out-File -Append $logpath

if(($csvfile.password | Measure-Object -Line).Lines -eq 0){
        Write-Host The password attribute was not detected and a random password is being used. -Foreground "Yellow"
    }
    else{
        Write-Host The password attribute was detected and the specified password is being used. -Foreground "Yellow"
    }

for($i=0;$i -lt $id.Count;$i++){
    $headers = @{"Authorization" = $token[1]}
    
    if(($csvfile.password | Measure-Object -Line).Lines -eq 0){
        $Url = "https://$ip/passwordvault/api/Accounts/" + $id[$i] + "/Change"
        $Body = @{ChangeEntireGroup = "false"}
    }
    else{
        $Url = "https://$ip/passwordvault/api/Accounts/" + $id[$i] + "/SetNextPassword"
        $Body = @{ChangeImmediately = "true";NewCredentials = $password[$i]}
    }
    
    Write-Host Changing password Name: $name[$i] UserName: $username[$i] ID: $id[$i]

    try{
        $getchangeresult = Invoke-WebRequest -Method 'Post' -Uri $Url -Body $body -Headers $headers
    }
    catch{
        ("Name: " + $name[$i] + " UserName: " + $username[$i] + " ID: " + $id[$i]) | Out-File -Append $logpath
        $err = $_
        ($err.Exception.Message) | Out-File -Append $logpath
        ($err.ErrorDetails.Message) | Out-File -Append $logpath
        "" | Out-File -Append $logpath
        $nochangeid.Add($name[$i])
    }
}

$numnochangeid = $nochangeid.Count

if($numnochangeid -eq 0){
    $ws = New-Object -ComObject WScript.Shell
    $wsr = $ws.popup("All successfully! For more details, reading $logpath.",-1,"All successfully!",1)
}
else{
    $ws = New-Object -ComObject WScript.Shell
    $wsr = $ws.popup("$numnochangeid users have errors! For more details, reading $logpath.",-1,"Have some Error!",1)
}