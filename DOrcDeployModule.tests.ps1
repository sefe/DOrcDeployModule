
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Import-Module "$here\DOrcDeployModule.psm1" -ErrorAction Stop

Describe "Enhanced Parameter Security Tests" {
    Context "Basic functionality" {
        It "Hides sensitive parameter names" {
            Format-ParameterForLogging "PASSWORD=secret123" | Should Be "PASSWORD=***HIDDEN***"
        }
        
        It "Preserves non-sensitive parameters" {
            Format-ParameterForLogging "SERVER=localhost" | Should Be "SERVER=localhost"
        }
        
        It "Is case insensitive" {
            Format-ParameterForLogging "mypassword=test" | Should Be "mypassword=***HIDDEN***"
        }
    }
    
    Context "Connection string security (Azure SignalR, SQL, etc.)" {
        It "Masks AccessKey in Azure SignalR connection string" {
            $testInput = "SIGNALR_CONN=Endpoint=https://app.service.signalr.net;AccessKey=secret123;Version=1.0"
            $result = Format-ParameterForLogging $testInput
            $result | Should Match "AccessKey=\*\*\*HIDDEN\*\*\*"
            $result | Should Match "Endpoint=https://app.service.signalr.net"
        }
        
        It "Masks Password in SQL connection string" {
            $testInput = "DB_CONN=Server=srv;Database=db;Password=secret123;Timeout=30"
            $result = Format-ParameterForLogging $testInput
            $result | Should Match "Password=\*\*\*HIDDEN\*\*\*"
            $result | Should Match "Server=srv"
        }
        
        It "Handles multiple secrets in one connection string" {
            $testInput = "CONN=Server=test;Password=secret;AccessKey=key123;Database=mydb"
            $result = Format-ParameterForLogging $testInput
            $result | Should Match "Password=\*\*\*HIDDEN\*\*\*"
            $result | Should Match "AccessKey=\*\*\*HIDDEN\*\*\*" 
            $result | Should Match "Server=test"
        }
    }
}

Describe "Get-DorcCredSSPStatus tests" {
    Context "Computer reachable"{
        Context "Returns an object with all the expected properties"{
            $username = "nwtraders\administrator" 
            $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
            $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
            #Client CredSSP enabled
            Mock -CommandName Get-WSManCredSSP -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."}
            #Server CredSSP enabled, Hotfix vulnerability workaround not in place, e.g. AllowEncryptionOracle value is different to 2
            Mock -CommandName Get-WmiObject -MockWith {return (New-Object psobject -Property @{"Version"="10.0.17134"; "AllowEncryptionOracle"= 0})}
            Mock -CommandName Invoke-Command -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1
            This computer is configured to receive credentials from a remote client computer."}
            #Return hotfix not relevant to CredSSP vulnerability
            Mock Get-MSHotfix {Return (New-Object psobject -Property @{"HotFixID"="KB12345678"})}
            It "Returns an object with all the expected properties" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $TestCred -Test
                $result.LocalComputerName | Should -Be $Env:COMPUTERNAME
                $result.RemoteComputerName | Should -Be "SERVER1"
                $result.LocalOS | Should -Be "10.0.17134"
                $result.RemoteOS | Should -Be "10.0.17134"
                $result.LocalCredSSPEnabled | Should -Match "true|false"
                $result.RemoteCredSSPEnabled | Should -Match "true|false"
                $result.LocalPatchInstalled | Should -Match "true|false"
                $result.RemotePatchInstalled | Should -Match "true|false"
                $result.LocalHotFixWorkaroundInPlace | Should -Match "true|false" #can be either depending on local machine it's executed on
                $result.RemoteHotFixWorkaroundInPlace | Should -Match "true|false"
                $result.CredSSPWorks | Should -Match "true|false"
            }
        }

        Context "Credential Delegation Enabled on Client and on Server"{
            $username = "nwtraders\administrator" 
            $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
            $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
            #Client CredSSP enabled
            Mock -CommandName Get-WSManCredSSP -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."}
            #Server CredSSP enabled, Hotfix vulnerability workaround not in place, e.g. AllowEncryptionOracle value is different to 2
            Mock -CommandName Get-WmiObject -MockWith {return (New-Object psobject -Property @{"Version"="10.0.17134"; "AllowEncryptionOracle"= 0})}
            Mock -CommandName Invoke-Command -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1
            This computer is configured to receive credentials from a remote client computer."}
            #Return hotfix not relevant to CredSSP vulnerability
            Mock Get-MSHotfix {Return (New-Object psobject -Property @{"HotFixID"="KB12345678"})}
            
            It "Returns LocalCredSSPEnabled and RemoteCredSSPEnabled as $true " {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $TestCred
                $result.LocalCredSSPEnabled | Should -Be $True
                $result.RemoteCredSSPEnabled | Should -Be $True
            }
        }
        Context "Credential Delegation NOT enabled on Client"{
            $username = "nwtraders\administrator" 
            $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
            $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
            #Client CredSSP enabled
            Mock -CommandName Get-WSManCredSSP -MockWith {return "The machine is noy configured to allow delegating fresh credentials.`nThis computer is not configured to receive credentials from a remote client computer."}
            #Server CredSSP enabled, Hotfix vulnerability workaround not in place, e.g. AllowEncryptionOracle value is different to 2
            Mock -CommandName Get-WmiObject -MockWith {return (New-Object psobject -Property @{"Version"="10.0.17134"; "AllowEncryptionOracle"= 0})}
            Mock -CommandName Invoke-Command -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1
            This computer is configured to receive credentials from a remote client computer."}
            #Return hotfix not relevant to CredSSP vulnerability
            Mock -CommandName Get-MSHotfix -MockWith {Return (New-Object psobject -Property @{"HotFixID"="KB12345678"})}
            
            It "Returns LocalCredSSPEnabled as $false" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $TestCred
                $result.LocalCredSSPEnabled | Should -Be $false
            }
        }

        Context "Credential Delegation enabled on Client but delegated computer doesn't match remote computer's name"{
            $username = "nwtraders\administrator" 
            $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
            $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
            #Client CredSSP enabled
            Mock -CommandName Get-WSManCredSSP -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER2`nThis computer is not configured to receive credentials from a remote client computer."}
            #Server CredSSP enabled, Hotfix vulnerability workaround not in place, e.g. AllowEncryptionOracle value is different to 2
            Mock -CommandName Get-WmiObject -MockWith {return (New-Object psobject -Property @{"Version"="10.0.17134"; "AllowEncryptionOracle"= 0})}
            Mock -CommandName Invoke-Command -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1
            This computer is configured to receive credentials from a remote client computer."}
            #Return hotfix not relevant to CredSSP vulnerability
            Mock -CommandName Get-MSHotfix -MockWith {Return (New-Object psobject -Property @{"HotFixID"="KB12345678"})}
            
            It "Returns LocalCredSSPEnabled as $false" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $TestCred
                $result.LocalCredSSPEnabled | Should -Be $false
            }
        }
        Context "Credential Delegation not enabled on Server"{
            $username = "nwtraders\administrator" 
            $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
            $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
            #Client CredSSP enabled
            Mock -CommandName Get-WSManCredSSP -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."}
            #Server CredSSP enabled, Hotfix vulnerability workaround not in place, e.g. AllowEncryptionOracle value is different to 2
            Mock -CommandName  Get-WmiObject -MockWith {return (New-Object psobject -Property @{"Version"="10.0.17134"; "AllowEncryptionOracle"= 0})}
            Mock -CommandName Invoke-Command -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."}
            #Return hotfix not relevant to CredSSP vulnerability
            Mock -CommandName Get-MSHotfix -MockWith {Return (New-Object psobject -Property @{"HotFixID"="KB12345678"})}
            
            It "Returns an object with all the expected properties" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $TestCred
                $result.RemoteCredSSPEnabled | Should -Be $false
            }
        }
    }
    Context "Computer not reachable"{ 
        $username = "nwtraders\administrator" 
        $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
        $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
        Mock -CommandName Get-WmiObject -MockWith {return (New-Object psobject -Property @{"Version"="10.0.17134"; "AllowEncryptionOracle"= 0})}
        It "Fails" {
            {Get-DorcCredSSPStatus -ComputerName Server1 -Credential $TestCred} | should -Throw "Failed to connect"
        }
    }
}

Describe "Enable-DorcCredSSP tests" {
    Context "Computer reachable" {
        Context "Already enabled" {
            $username = "nwtraders\administrator" 
            $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
            $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
            Mock -CommandName Get-DorcCredSSPStatus -MockWith {return ((New-Object psobject -Property `
                @{
                    "LocalComputerName"             = $Env:COMPUTERNAME
                    "RemoteComputerName"            = "SERVER1"
                    "LocalOS"                       = "10.0.17134"
                    "RemoteOS"                      = "10.0.14393"
                    "LocalCredSSPEnabled"           = "True"
                    "RemoteCredSSPEnabled"          = "True"
                    "LocalPatchInstalled"           = "False"
                    "RemotePatchInstalled"          = "False"
                    "LocalHotFixWorkaroundInPlace"  = "False"
                    "RemoteHotFixWorkaroundInPlace" = "False"
                    "CredSSPWorks"                  = "True"
                }))}
            It "Nothing to do" {
                $result = Enable-DorcCredSSP -ComputerName SERVER1 -Credential $TestCred
                $result | Should -Be "[INFO] CredSSP has already been enabled between [$Env:COMPUTERNAME] and [SERVER1]"
            }

        }
        Context "Not yet enabled" {
            Context "Return $false and then $true for CredSSPWorks property with each execution of Get-DorcCredSSPStatus"{
                $username = "nwtraders\administrator" 
                $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
                $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
                #.CredSSPWorks needs to first return $false and then $true. This little loop allows for this.
                $script:mockCalled = 0
                $MockDorcCredSSPStatus = {
                    $script:mockCalled++
                    if ($script:mockCalled %2 -eq 1) { #each odd run is $false 
                        return ((New-Object psobject -Property `
                            @{
                                "LocalComputerName"             = $Env:COMPUTERNAME
                                "RemoteComputerName"            = "SERVER1"
                                "LocalOS"                       = "10.0.17134"
                                "RemoteOS"                      = "10.0.14393"
                                "LocalCredSSPEnabled"           = "False"
                                "RemoteCredSSPEnabled"          = "False"
                                "LocalPatchInstalled"           = "False"
                                "RemotePatchInstalled"          = "False"
                                "LocalHotFixWorkaroundInPlace"  = "False"
                                "RemoteHotFixWorkaroundInPlace" = "False"
                                "CredSSPWorks"                  = "False"
                            }))
                    }
                    else {
                        return ((New-Object psobject -Property `
                            @{
                                "LocalComputerName"             = $Env:COMPUTERNAME
                                "RemoteComputerName"            = "SERVER1"
                                "LocalOS"                       = "10.0.17134"
                                "RemoteOS"                      = "10.0.14393"
                                "LocalCredSSPEnabled"           = "True"
                                "RemoteCredSSPEnabled"          = "True"
                                "LocalPatchInstalled"           = "True"
                                "RemotePatchInstalled"          = "True"
                                "LocalHotFixWorkaroundInPlace"  = "True"
                                "RemoteHotFixWorkaroundInPlace" = "True"
                                "CredSSPWorks"                  = "True"
                            }))
                    }
                    
                }
                Mock -CommandName Invoke-Command -MockWith {return "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1
                This computer is configured to receive credentials from a remote client computer."}
                Mock -CommandName Get-DorcCredSSPStatus -MockWith $MockDorcCredSSPStatus
                Mock -CommandName Get-MSHotfix -MockWith {Return (New-Object psobject -Property @{"HotFixID"="KB12345678"})}
                It "Enables CredSSP successfully" {
                    $result = Enable-DorcCredSSP -ComputerName SERVER1 -Credential $TestCred
                    $result | Should -Be "CredSSP Successfully enabled between [$ENV:ComputerName] and [SERVER1]"
                }
            }
        }
    }
    Context "Computer not reachable" {
        $username = "nwtraders\administrator" 
        $password = "mypassword" | ConvertTo-SecureString -asPlainText -Force
        $TestCred = New-Object System.Management.Automation.PSCredential($username,$password)
        Mock Get-DorcCredSSPStatus {return ((New-Object psobject -Property `
            @{
                "LocalComputerName"             = $Env:COMPUTERNAME
                "RemoteComputerName"            = "SERVER1"
                "LocalOS"                       = "10.0.17134"
                "RemoteOS"                      = "10.0.14393"
                "LocalCredSSPEnabled"           = "True"
                "RemoteCredSSPEnabled"          = "False"
                "LocalPatchInstalled"           = "False"
                "RemotePatchInstalled"          = "False"
                "LocalHotFixWorkaroundInPlace"  = "False"
                "RemoteHotFixWorkaroundInPlace" = "False"
                "CredSSPWorks"                  = "False"
            }))}
        It "Throws" {
            {Enable-DorcCredSSP -ComputerName nonexistentcomputer -Credential $TestCred } | Should -Throw 'Failed'
        }
    }
}
