Function welcome{

        Clear;
        Write-Host "******************************************************"  -foreground Red 
        Write-Host "     Welcome to Office365 Administrator Tool          " -foreground Red 
        Write-Host "******************************************************"  -foreground Red 
}

Function Login {

        while(1){
                Welcome;
                Import-Module MsOnline;
                Write-Host "step 1" -ForegroundColor yellow;
                Write-Host " Enter Office365 account : " -nonewline
                $global:adm_account = Read-Host ;
                Write-Host "--------------------------------------------------"-ForegroundColor yellow;

                Write-Host "step 2" -ForegroundColor yellow;
                Write-Host " Please enter your password : " -nonewline;
                $global:adm_password_plain = Read-Host -assecurestring;
                $global:adm_password_plain = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($adm_password_plain));

                $global:adm_password_encrypt = convertto-securestring  $adm_password_plain -asplaintext -force

                $global:adm_cred=New-Object System.Management.Automation.PSCredential($adm_account,$adm_password_encrypt);
                $report = Connect-MsolService -credential $adm_cred 2>&1;
                $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
				                if($err){
					                Write-Host " $report" -background Black -foreground Red	
					                Read-Host;
                                    Clear;
				                }
				                else{
					                Write-Host "Login Success!" -background Black -foreground Magenta	
                                    Write-Host "--------------------------------------------------"-ForegroundColor yellow;
					                Break;
				                }
        }    
} 

Function DirectLogin {
        Welcome;		
        Import-Module MsOnline
        $global:adm_account = "test@domain.com"
        $global:admpassword_plain = "1234567890"
        $global:admpassword_encrypt = convertto-securestring  $admpassword_plain -asplaintext -force
        $global:adm_cred=New-Object System.Management.Automation.PSCredential($adm_account,$admpassword_encrypt)
        Connect-MsolService -credential $adm_cred
}

#1
Function Check_users_status{

			Get-MsolUser -All | Out-GridView -Title "Choose the User account you want to check" -PassThru | fl  #(format list) 
}

#2
Function Create_user{

        $UserAccount = Read-Host "Account ";
        $split = $adm_account.split("@");

        $UserPrincipalName = $UserAccount + "@" + $split[1];
        $NewPassword = Read-Host "Password ";
        $FirstName = Read-Host "FirstName ";
        $LastName = Read-Host "LastName ";
        $DisplayName = Read-Host "DisplayName ";
        $report = New-MsolUser -FirstName $FirstName -LastName $LastName -UserPrincipalName $UserPrincipalName -DisplayName $DisplayName -Password $NewPassword -ForceChangePassword $false 2>&1;
        $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
		if($err){
            Write-Host $err -BackgroundColor Black -ForegroundColor Red;
        }
        else{
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
            ForEach ($item in $ServicePlans){                         #ForEach ($<each item> in $<group object>)
                     Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $item.AccountSkuId
            }
            Write-Host "Create User Done!" -BackgroundColor Black -ForegroundColor Green;
        }
}

#3
Function Create_multi_user{
    
    Get-MsolUser -UserPrincipalName $adm_account | Select-Object UserPrincipalName,Password,FirstName,LastName,DisplayName | Export-Csv sample_new_users.csv;
    $CurrentPath = $(Get-Location).ToString();
    Write-Host "We create a reference sample at $CurrentPath\sample_new_users.csv"  -background Black -foreground Magenta
    while(1){
        $Importfile = Read-Host "Please enter the csv filename ";
        $report =  Get-Content -path $CurrentPath"\"$Importfile  2>&1;
        $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
		if($err){
			Write-Host "Can't open file." -background Black -foreground Red	
		}
		else{
            Write-Host "Reading $Importfile ..." -background Black -foreground Magenta	
            Break;
		}
    }

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
    ForEach ($item in $ServicePlans){                        #ForEach ($<each item> in $<group object>)
        $Licenses = $item.AccountSkuId;
    }

    Import-Csv $Importfile | ForEach-Object {
        $report = New-MsolUser -FirstName $_.FirstName -LastName $_.LastName -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName -Password $_.Password -ForceChangePassword $false 2>&1;
        $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
		if($err){
            Write-Host $err -BackgroundColor Black -ForegroundColor Red;
        }
        else{
            $UserPrincipalName = $_.UserPrincipalName;
            Set-MsolUser -UserPrincipalName $_.UserPrincipalName -PasswordNeverExpires $true
            Set-MsolUser -UserPrincipalName $UserPrincipalName -UsageLocation TW
            Set-MsolUserLicense -UserPrincipalName $UserPrincipalName -AddLicenses $Licenses
            Write-Host "Create $UserPrincipalName Success" -BackgroundColor Black -ForegroundColor Green;
        }
    
    }
    Write-Host "Create All User Done!" -BackgroundColor Black -ForegroundColor Green;

}

#4
Function Remove_user{

        Get-MsolUser -all | Out-GridView -Title "Choose the users you want to Delete" -PassThru | ForEach-Object { Remove-MsolUser -UserPrincipalName  $_.UserPrincipalName -force}
        Write-Host "Remove User Done!" -BackgroundColor Black -ForegroundColor Green;
}

#5
Function Recover_user{

         Get-MsolUser -ReturnDeletedUsers | Out-GridView -Title "Choose the users you want to Recover" -PassThru | ForEach-Object {Restore-MsolUser -UserPrincipalName  $_.UserPrincipalName -AutoReconcileProxyConflicts}    
         Write-Host "Recover User Done!" -BackgroundColor Black -ForegroundColor Green
}

#6
Function Change_Password{
        Write-Host "1. Table mode 2. csv mode : " -NoNewline -ForegroundColor Yellow ;
        $choose = Read-Host
        if($choose -eq 1){
                $NewPass = Read-Host "Enter New Password "
                Get-MsolUser -all | Out-GridView -Title "Choose the users you want to Change Password" -PassThru | ForEach-Object {
                        Set-MsolUserPassword -UserPrincipalName $_.UserPrincipalName -NewPassword $NewPass -ForceChangePassword $false
                        Set-MsolUser -UserPrincipalName $_.UserPrincipalName -PasswordNeverExpires $true
                }
                Write-Host "Change Password Done!" -BackgroundColor Black -ForegroundColor Green
        }
        elseif ($choose -eq 2){
                Get-MsolUser -UserPrincipalName $adm_account | Select-Object UserPrincipalName,NewPassword | export-csv sample_change_pwd.csv -encoding "utf8"
                $CurrentPath = $(Get-Location).ToString();
                Write-Host "We create a reference sample at $CurrentPath\sample_change_pwd.csv"  -background Black -foreground Magenta
                while(1){
                        $Importfile = Read-Host "Please enter the csv filename ";
                        $report =  Get-Content -path $CurrentPath"\"$Importfile  2>&1;
                        $err = $report | ?{$_.gettype().Name -eq "ErrorRecord"}
			            if($err){
					               Write-Host "Can't open file." -background Black -foreground Red	
			            }
			            else{
                                   Write-Host "Reading $Importfile ..." -background Black -foreground Magenta	
                                   Import-Csv $Importfile | ForEach-Object {
                                            Set-MsolUserPassword ¡VUserPrincipalName $_.UserPrincipalName -NewPassword $_.NewPassword ¡VForceChangePassword $false
                                            Set-MsolUser -UserPrincipalName $_.UserPrincipalName -PasswordNeverExpires $true
                                   }
                                   Write-Host "Change Password Done!" -BackgroundColor Black -ForegroundColor Green
                                   Break;
			            }
                }
        }
}

#7
Function Send_mail{

        $SMTPServer = "smtp.office365.com";
        $SMTPPort = "587";
        $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
        $smtp.EnableSSL = $true;
        $smtp.Credentials = New-Object System.Net.NetworkCredential($adm_account, $admpassword_plain);

        $From = $adm_account;
        Write-Host "From : $From"
        $To = Read-Host "To ";
        $subject = Read-Host "Subject ";
        $body = Read-Host "Content  ";
        $smtp.Send($From, $To, $subject, $body); 
}

#8
Function See_mail_log{

        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $adm_cred -Authentication Basic ¡VAllowRedirection
        Import-PSSession $Session
        Get-StaleMailboxReport | Out-GridView -Title "Main Log"	

}


<# Main #>
	
		Login
        #DirectLogin;

		while (1) {
				welcome;
				$ident = Get-MsolUser -UserPrincipalName $adm_account
				Write-Host "$ident "   -BackgroundColor Black -ForegroundColor Magenta
				Write-Host "Login : $adm_account "   -BackgroundColor Black  -ForegroundColor Magenta
				Write-Host "  1.Check users status"  -foreground Yellow 
				Write-Host "  2.Create user"-foreground Yellow
                Write-Host "  3.Create Muti-user"-foreground Yellow
				Write-Host "  4.Remove user"-foreground Yellow
				Write-Host "  5.Recover user"-foreground Yellow 
				Write-Host "  6.Change password" -foreground Yellow 
				Write-Host "  7.Send mail"-foreground Yellow 
				Write-Host "  8.See mail log"-foreground Yellow  
				Write-Host "  9.Logout" -foreground Yellow
				Write-Host "  0.Exit" -foreground Yellow
				 
				$choose = Read-Host "Please choose the number "

				switch ($choose) 
				{ 
						1{ Check_users_status } 
						2{ Create_user }
                        3{ Create_multi_user }
						4{ Remove_user}
						5{ Recover_user}
						6{ Change_Password}
						7{ Send_mail }
						8{ See_mail_log }
						9{ Remove-Module MSOnline;
							Login 
						}
						default { ; }
				}
			    if($choose -ne -0){
                        Write-Host "Press any key to continue" -ForegroundColor Red;
					    Read-Host;
			    }
                else{
                        break;
                } 
		}
		Remove-Module MSOnline

<# End Main #>