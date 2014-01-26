Function welcome{

        Clear
        Write-Host "******************************************************" -foreground Red
        Write-Host "     Welcome to Office365 Administrator Tool        " -foreground Red 
        Write-Host "******************************************************" -foreground Red
}
Function Check_users_status{

        Write-Host "Enter user account you want to check : " -nonewline
        $careaccount = Read-Host
        Get-MsolUser -UserPrincipalName $careaccount | fl  #(format list)
}
Function Change_Password{

        $careaccount = Read-Host "Enter user account to change password "
        $NewPass = Read-Host "Enter New Password "
        Set-MsolUserPassword -userPrincipalName $careaccount -NewPassword $NewPass -ForceChangePassword $false
        Set-MsolUser -UserPrincipalName $careaccount -PasswordNeverExpires $true
        Write-Host "Change Password Done!";
}

Function Create_user{

        $UserPrincipalName = Read-Host "Account ";
        $NewPassword = Read-Host "Password";
        $FirstName = Read-Host "FirstName ";
        $LastName = Read-Host "LastName ";
        $DisplayName = Read-Host "DisplayName ";
        New-MsolUser -FirstName $FirstName -LastName $LastName -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName -Password $NewPassword -ForceChangePassword $false
        Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation TW
        Get-MsolAccountSku |ft AccountSkuId,ActiveUnits,ConsumedUnits
        Write-Host "Choose the Licenses to Active " -nonewline
        Write-Host "1.STUDENT 2.FACULTY : " -nonewline -foreground Red
        $Licenses_no = Read-Host ;
        if($Licenses_no -eq "1"){
                $ServicePlans = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq "ENTERPRISEPACK_STUDENT"};
        }
        elseif($Licenses_no -eq "2"){
                 $ServicePlans = Get-MsolAccountSku | Where {$_.SkuPartNumber -eq "ENTERPRISEPACK_FACULTY"} ;                
        }
        ForEach ($item in $ServicePlans){                         #ForEach ($<個別的項目或物件> in $<集合物件>)
                 Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $item.AccountSkuId
        }
        Write-Host "Create User Done!";
}

Function Remove_user{
        Get-MsolUser -all | Out-GridView -PassThru | ForEach-Object { Remove-MsolUser -UserPrincipalName  $_.UserPrincipalName -force}
        Write-Host "Remove User Done!";
}
Function Recover_user{
         Get-MsolUser -ReturnDeletedUsers | Out-GridView -PassThru | ForEach-Object {Restore-MsolUser -UserPrincipalName  $_.UserPrincipalName -AutoReconcileProxyConflicts}    
         Write-Host "Recover User Done!";
}

Function  Send_mail
{
        $SMTPServer = "smtp.office365.com";
        $SMTPPort = "587";
        $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
        $smtp.EnableSSL = $true;
        $smtp.Credentials = New-Object System.Net.NetworkCredential($admaccount, $admpassword_plain);

        $From = $admaccount;
        Write-Host "From : $From"
        $To = Read-Host "To ";
        $subject = Read-Host "Subject ";
        $body = Read-Host "Content  ";
        $smtp.Send($From, $To, $subject, $body); 
}

####Main Begin####
(Get-Host).UI.RawUI.BackgroundColor = "black"
welcome;
Import-Module MsOnline
Write-Host "Enter Office365 account : " -nonewline -foreground "yello"
$admaccount = Read-Host 
#$admaccount = "test@domain.com"	#enter account
Write-Host "Enter Password : " -nonewline -foreground "yello"
$admpassword_plain = Read-Host
#$admpassword_plain = "P@ssw0rd"	#enter password
$admpassword_encrypt = convertto-securestring  $admpassword_plain -asplaintext -force

$cred=New-Object System.Management.Automation.PSCredential($admaccount,$admpassword_encrypt)
Connect-MsolService -credential $cred

$i = 1;
while ($i) {
        welcome;
        Write-Host "1.Check users status"  -foreground Yellow 
        Write-Host "2.Change password" -foreground Yellow 
        Write-Host "3.Send mail"-foreground Yellow 
        Write-Host "5.Create user"-foreground Yellow 
        Write-Host "7.Remove user"-foreground Yellow
        Write-Host "8.Recover user"-foreground Yellow  
        Write-Host "9.Exit" -foreground Yellow 

        $choose = Read-Host "Please choose the number "

        switch ($choose) 
        { 
                1{ Check_users_status } 
                2{ Change_Password   }
                3{ Send_mail }
                5{ Create_user }
                7{ Remove_user}
                8{ Recover_user}
                default { $i = 0; }
        }
        Pause
}
Remove-Module MSOnline
####Main  End####