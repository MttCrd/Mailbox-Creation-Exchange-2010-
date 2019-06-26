param (

    [string]$Domain = $(throw "-Domain is required.")
	
)

Import-Module ActiveDirectory

# ======================================================
#Database
# ======================================================

$Database = ""

Write-Host ""

# ======================================================
#LogonDomain
# ======================================================

    if ($Domain -eq 1){ 

        $cred = 

        $Server = ""

        $DomainController = ""

        $MasterAccount = ""

    }

    if ($Domain -eq 2){ #

        $cred = 

        $Server = ""

        $DomainController = ""

        $MasterAccount = ""
        
    }

		# ======================================================
        # Checking available space on Database
        # ======================================================
		
		Write-Verbose "Checking available space on Database" -Verbose
		
		$Count = (Get-Mailbox -Database $Database).Count
		
		$Free = 250-$Count
		
		Write-Host ""
		
		Write-Host "Free Space on" $Database :$Free
		
		if ($Free -le 20){
		
		Write-Host ""
		
		Write-Host "Free Space Runinng Out" -foregroundcolor red
		
		}
		
        # ======================================================
        # Import CSV (User)
        # ======================================================
        
        import-csv UserCreation.csv | Foreach-Object {

        $Alias = $_.Alias

        # ======================================================
        # Checking if the mailbox already exist
        # ======================================================        
		
        $Check = Get-Mailbox $Alias -erroraction 'silentlycontinue'

        if ($Check)  {

            Write-Host "Mailbox already Exist" -foregroundcolor green

            Write-Host ""

            Write-Host "#===================================================#" -foregroundcolor blue

            Write-Host ""

            Write-Host "DisplayName     : " $Check.DisplayName            -foregroundcolor yellow
            Write-Host "Alias           : " $Check.Alias                  -foregroundcolor yellow
            Write-Host "Creation Date  : " $Check.WhenCreated            -foregroundcolor yellow
            Write-Host "SMTP            : " $Check.WindowsEmailAddress    -foregroundcolor yellow

            Write-Host ""

            Write-Host "#===================================================#" -foregroundcolor blue

            Write-Host ""

            Sleep -s 2

            $Continue = Read-Host "You want to continue with the overwriting of the policies? : (yes) - (no)"

            if ($Continue -eq 'no') {

                Exit

            }

        }

        # ======================================================
        # Getting Users info from AD
        # ======================================================

		# ======================================================
        # Name
		# ======================================================
		
        $Name = Get-ADUser -Server $Server -Credential $cred $Alias -Properties * | select GivenName

        $Name1 = $Name.GivenName
		
		# ======================================================
		# Name conversion keeping the first capital letter
		# ======================================================

        $LowerCaseName = (Get-Culture).textinfo.totitlecase("$Name1".tolower())
		
		# ======================================================
        # Surname
		# ======================================================

        $Surname = Get-ADUser -Server $Server -Credential $cred $Alias -Properties * | select sn

        $Surname1 = $Surname.sn
		
		# ======================================================
		# Surname conversion keeping the first capital letter
		# ======================================================
    
        $LowerCaseSurname = (Get-Culture).textinfo.totitlecase("$Surname1".tolower())

		# ======================================================
        # Company
		# ======================================================
		
        $CompanyUser = Get-ADUser -Server $Server -Credential $cred $Alias -Properties * | select Company
		
        # ======================================================
        # User Info (Video)
        # ======================================================

        Write-Host ""

        Write-Host "User Info:" -foregroundcolor white        

        Write-Host ""

        Write-Host "NAME        : " $LowerCaseName 			-foregroundcolor yellow
        Write-Host "SURNAME     : " $LowerCaseSurname 		-foregroundcolor yellow
        Write-Host "ALIAS	    : " $Alias 				-foregroundcolor yellow
        Write-Host "COMPANY     : " $CompanyUser.Company 	-foregroundcolor yellow
		
        Write-Host ""

        #Write-Host " - Press a Key to Continue..."  -foregroundcolor red

        #$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        Write-Host ""

        # ======================================================
        # Company selection from ActiveDirectory
        # ======================================================

		# ======================================================
        # Company List
		# ======================================================

        if ($CompanyUser.Company -eq ''){

            $OU = ""

            $CA9 = ""

            $Value = 1

            }

        if ($CompanyUser.Company -eq ''){

            $OU = ""

            $CA9 = ""

            $Value = 1

            }

        if ($Value -ne 1){

            Write-Host "Field Empty" -foregroundcolor red
			
			Write-Host ""
			 
			$CompanyUser.Company = Read-Host "Insert Company (Full Name)"

            Write-Host ""
			 
			Exit

        }

        # ======================================================
	
            New-Mailbox -Database "$Database" –FirstName "$LowerCaseName" –LastName "$LowerCaseSurname" –DisplayName "$LowerCaseSurname $LowerCaseName" –Name "$Alias" –Alias "$Alias" -LinkedDomainController $DomainController -LinkedMasterAccount "$MasterAccount\$Alias" –OrganizationalUnit $OU -UserPrincipalName "$Alias@enirf.res.prirf" -LinkedCredential $cred –ResetPasswordOnNextLogon $false -Confirm:$false
        
			}
			
            Sleep -s 5

            Set-Mailbox –identity "$Alias" -UseDatabaseQuotaDefaults $true –SingleItemRecoveryEnabled $true -Verbose -Confirm:$false
        
            Set-Mailbox –identity "$Alias" -ManagedFolderMailboxPolicy '' -ManagedFolderMailboxPolicyAllowed -Confirm:$false
        
            Set-User -identity "$Alias" -Company $CompanyUser.Company -Confirm:$false
                    
            Set-CASMailbox –Identity "$Alias" -POPEnabled:$false -ImapEnabled:$false -Confirm:$false
            
            Write-Host ""
            
            $employee = Get-ADUser -Server $Server -Credential $cred $Alias -Properties * | select employeeType
            
            Sleep -s 2

            # ======================================================
            # Getting Employee type from AD
            # ====================================================== 
            
            if ($employee.employeeType -eq 'D'){
        
                $EmployeeID = Get-ADUser -Server $Server -Credential $cred $Alias -Properties * | select employeeID

                $EID = $EmployeeID.employeeID

                Write-Host ""

                Write-Host "#=====================================#"

                Write-Host "Employee - EmployeeID: " $EID -foregroundcolor yellow

                Write-Host "#=====================================#"

                Sleep -s 2
				
				Write-Host ""
				
				Write-Verbose "Setting CA9" -Verbose

				Write-Verbose "" -Verbose
                                
                Set-Mailbox -identity "$Alias" -CustomAttribute8 "D" -CustomAttribute9 $CA9

                sleep -s 5

                Set-Mailbox -identity "$Alias" -EmailAddressPolicyEnabled $false -Confirm:$false

                sleep -s 5

                Set-ADUser -Identity "$Alias" -Replace @{employeeID="$EID"}

                # ======================================================
                # Output video of some useful Info
                # ====================================================== 

                $Smtp = Get-Mailbox -identity "$Alias" | select PrimarySmtpAddress

                $LinkedAccount = Get-Mailbox -identity "$Alias" | select LinkedMasterAccount

                $UserName = Get-Mailbox -identity "$Alias" | select UserPrincipalName

                $SmtpTxt = $Smtp.PrimarySmtpAddress

                $LinkedMasterAccountTxt = $LinkedAccount.LinkedMasterAccount

                $UserNameTxt = $UserName.UserPrincipalName

                $WC = Get-Mailbox $Alias | select WhenCreated

                $Date = $WC.WhenCreated

                Write-Host ""

                Write-Host "Smtp Primario       : " $SmtpTxt

                Write-Host "LinkedMasterAccount : " $LinkedMasterAccountTxt

                Write-Host "UserPrincipalName   : " $UserNameTxt

                Write-Host "OrganizationalUnit  : " $OU

                Write-Host "CustomAttribute9    : " $CA9
				
				Write-Host ""
				
                # ======================================================
				# Sending Email with the Log
                # ======================================================
	
                .\2-MailLog.ps1
                
            } else {

                Write-Host ""

                Write-Host "#========#"
            
                Write-Host "Consultant" -foregroundcolor yellow

                Write-Host "#========#"

                Write-Host ""
            
                Set-Mailbox -identity "$Alias" -CustomAttribute8 "C" -CustomAttribute9 $CA9
                
                Sleep -s 5
            
                Set-Mailbox -identity "$Alias" -EmailAddressPolicyEnabled $false -Confirm:$false
                
                Sleep -s 5

                # ======================================================
                # Output video of some useful Info
                # ======================================================
            
                $Smtp = Get-Mailbox -identity "$Alias" | select PrimarySmtpAddress

                $LinkedAccount = Get-Mailbox -identity "$Alias" | select LinkedMasterAccount

                $UserName = Get-Mailbox -identity "$Alias" | select UserPrincipalName

                $SmtpTxt = $Smtp.PrimarySmtpAddress

                $LinkedMasterAccountTxt = $LinkedAccount.LinkedMasterAccount

                $UserNameTxt = $UserName.UserPrincipalName

                $WC = Get-Mailbox $Alias | select WhenCreated

                $Date = $WC.WhenCreated

                Write-Host ""

                Write-Host "Primary Smtp        : " $SmtpTxt

                Write-Host "LinkedMasterAccount : " $LinkedMasterAccountTxt

                Write-Host "UserPrincipalName   : " $UserNameTxt

                Write-Host "OrganizationalUnit  : " $OU

                Write-Host "CustomAttribute9    : " $CA9
				
				Write-Host ""

                # ======================================================
				# Sending Email with the Log
                # ======================================================

                .\2-MailLog.ps1
        
            }
