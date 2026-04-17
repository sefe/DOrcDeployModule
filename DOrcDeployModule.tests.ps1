BeforeAll {
    $here = Split-Path -Parent $PSCommandPath
    Import-Module "$here\DOrcDeployModule.psm1" -Force -ErrorAction Stop
}

Describe "Get-DorcCredSSPStatus tests" {
    Context "Computer reachable" {

        Context "Returns an object with all the expected properties" {
            BeforeAll {
                $script:TestCred = New-Object System.Management.Automation.PSCredential(
                    'nwtraders\administrator',
                    ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

                # Bypass the admin-role guard (Get-DorcCredSSPStatus enforces
                # it at call time; all remote operations are mocked anyway).
                Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }

                # Client CredSSP enabled
                Mock -ModuleName DOrcDeployModule Get-WSManCredSSP {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."
                }
                # Get-ServerOSVersion is the internal function the code actually
                # uses for OS lookup (it wraps Get-CimInstance + a BuildNumber
                # switch map). Mock it directly rather than trying to fake the
                # CimInstance path — the test only cares that LocalOS/RemoteOS
                # surface the string in the resulting object.
                Mock -ModuleName DOrcDeployModule Get-ServerOSVersion { '10.0.17134' }
                Mock -ModuleName DOrcDeployModule Invoke-Command {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-MSHotfix {
                    New-Object psobject -Property @{ HotFixID = 'KB12345678' }
                }
            }

            It "Returns an object with all the expected properties" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $script:TestCred -Test
                $result.LocalComputerName | Should -Be $Env:COMPUTERNAME
                $result.RemoteComputerName | Should -Be "SERVER1"
                $result.LocalOS | Should -Be "10.0.17134"
                $result.RemoteOS | Should -Be "10.0.17134"
                $result.LocalCredSSPEnabled | Should -Match "true|false"
                $result.RemoteCredSSPEnabled | Should -Match "true|false"
                $result.LocalPatchInstalled | Should -Match "true|false"
                $result.RemotePatchInstalled | Should -Match "true|false"
                $result.LocalHotFixWorkaroundInPlace | Should -Match "true|false"
                $result.RemoteHotFixWorkaroundInPlace | Should -Match "true|false"
                $result.CredSSPWorks | Should -Match "true|false"
            }
        }

        Context "Credential Delegation Enabled on Client and on Server" {
            BeforeAll {
                $script:TestCred = New-Object System.Management.Automation.PSCredential(
                    'nwtraders\administrator',
                    ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

                # Bypass the admin-role guard (Get-DorcCredSSPStatus enforces
                # it at call time; all remote operations are mocked anyway).
                Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }

                Mock -ModuleName DOrcDeployModule Get-WSManCredSSP {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-WmiObject {
                    New-Object psobject -Property @{ Version = '10.0.17134'; AllowEncryptionOracle = 0 }
                }
                Mock -ModuleName DOrcDeployModule Invoke-Command {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-MSHotfix {
                    New-Object psobject -Property @{ HotFixID = 'KB12345678' }
                }
            }

            It "Returns LocalCredSSPEnabled and RemoteCredSSPEnabled as True" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $script:TestCred
                $result.LocalCredSSPEnabled | Should -Be $true
                $result.RemoteCredSSPEnabled | Should -Be $true
            }
        }

        Context "Credential Delegation NOT enabled on Client" {
            BeforeAll {
                $script:TestCred = New-Object System.Management.Automation.PSCredential(
                    'nwtraders\administrator',
                    ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

                # Bypass the admin-role guard (Get-DorcCredSSPStatus enforces
                # it at call time; all remote operations are mocked anyway).
                Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }

                Mock -ModuleName DOrcDeployModule Get-WSManCredSSP {
                    "The machine is not configured to allow delegating fresh credentials.`nThis computer is not configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-WmiObject {
                    New-Object psobject -Property @{ Version = '10.0.17134'; AllowEncryptionOracle = 0 }
                }
                Mock -ModuleName DOrcDeployModule Invoke-Command {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-MSHotfix {
                    New-Object psobject -Property @{ HotFixID = 'KB12345678' }
                }
            }

            It "Returns LocalCredSSPEnabled as False" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $script:TestCred
                $result.LocalCredSSPEnabled | Should -Be $false
            }
        }

        Context "Credential Delegation enabled on Client but delegated computer doesn't match remote computer's name" {
            BeforeAll {
                $script:TestCred = New-Object System.Management.Automation.PSCredential(
                    'nwtraders\administrator',
                    ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

                # Bypass the admin-role guard (Get-DorcCredSSPStatus enforces
                # it at call time; all remote operations are mocked anyway).
                Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }

                Mock -ModuleName DOrcDeployModule Get-WSManCredSSP {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER2`nThis computer is not configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-WmiObject {
                    New-Object psobject -Property @{ Version = '10.0.17134'; AllowEncryptionOracle = 0 }
                }
                Mock -ModuleName DOrcDeployModule Invoke-Command {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-MSHotfix {
                    New-Object psobject -Property @{ HotFixID = 'KB12345678' }
                }
            }

            It "Returns LocalCredSSPEnabled as False (delegated computer mismatch)" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $script:TestCred
                $result.LocalCredSSPEnabled | Should -Be $false
            }
        }

        Context "Credential Delegation not enabled on Server" {
            BeforeAll {
                $script:TestCred = New-Object System.Management.Automation.PSCredential(
                    'nwtraders\administrator',
                    ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

                # Bypass the admin-role guard (Get-DorcCredSSPStatus enforces
                # it at call time; all remote operations are mocked anyway).
                Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }

                Mock -ModuleName DOrcDeployModule Get-WSManCredSSP {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-WmiObject {
                    New-Object psobject -Property @{ Version = '10.0.17134'; AllowEncryptionOracle = 0 }
                }
                Mock -ModuleName DOrcDeployModule Invoke-Command {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-MSHotfix {
                    New-Object psobject -Property @{ HotFixID = 'KB12345678' }
                }
            }

            It "Returns RemoteCredSSPEnabled as False" {
                $result = Get-DorcCredSSPStatus -ComputerName Server1 -Credential $script:TestCred
                $result.RemoteCredSSPEnabled | Should -Be $false
            }
        }
    }

    Context "Computer not reachable" {
        BeforeAll {
            $script:TestCred = New-Object System.Management.Automation.PSCredential(
                'nwtraders\administrator',
                ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

            Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }
            # Local WSMan / OS succeed, then the remote Invoke-Command (which
            # the function uses to probe the target's WSMan state) fails —
            # that's the path that produces the "Failed to connect..." throw.
            Mock -ModuleName DOrcDeployModule Get-WSManCredSSP {
                "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is not configured to receive credentials from a remote client computer."
            }
            Mock -ModuleName DOrcDeployModule Get-ServerOSVersion { '10.0.17134' }
            Mock -ModuleName DOrcDeployModule Invoke-Command {
                throw [System.Management.Automation.RemoteException]::new('Connection refused')
            }
        }

        It "Fails" {
            { Get-DorcCredSSPStatus -ComputerName Server1 -Credential $script:TestCred } | Should -Throw "*Failed to connect*"
        }
    }
}

Describe "Enable-DorcCredSSP tests" {
    Context "Computer reachable" {

        Context "Already enabled" {
            BeforeAll {
                $script:TestCred = New-Object System.Management.Automation.PSCredential(
                    'nwtraders\administrator',
                    ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

                # Bypass the admin-role guard (Get-DorcCredSSPStatus enforces
                # it at call time; all remote operations are mocked anyway).
                Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }

                Mock -ModuleName DOrcDeployModule Get-DorcCredSSPStatus {
                    New-Object psobject -Property @{
                        LocalComputerName             = $Env:COMPUTERNAME
                        RemoteComputerName            = "SERVER1"
                        LocalOS                       = "10.0.17134"
                        RemoteOS                      = "10.0.14393"
                        LocalCredSSPEnabled           = "True"
                        RemoteCredSSPEnabled          = "True"
                        LocalPatchInstalled           = "False"
                        RemotePatchInstalled          = "False"
                        LocalHotFixWorkaroundInPlace  = "False"
                        RemoteHotFixWorkaroundInPlace = "False"
                        CredSSPWorks                  = "True"
                    }
                }
            }

            It "Nothing to do" {
                $result = Enable-DorcCredSSP -ComputerName SERVER1 -Credential $script:TestCred
                $result | Should -Be "[INFO] CredSSP has already been enabled between [$Env:COMPUTERNAME] and [SERVER1]"
            }
        }

        Context "Not yet enabled - Get-DorcCredSSPStatus returns False then True across successive calls" {
            BeforeAll {
                $script:TestCred = New-Object System.Management.Automation.PSCredential(
                    'nwtraders\administrator',
                    ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

                Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }
                # Enable-DorcCredSSP calls Enable-WSManCredSSP -Role Client
                # on the local machine, which requires elevation — mock it
                # away so the test runs on a non-elevated CI agent.
                Mock -ModuleName DOrcDeployModule Enable-WSManCredSSP { }

                # .CredSSPWorks needs to return False on the first call (before
                # Enable-DorcCredSSP does its work) and True on the second
                # (afterwards). A stateful counter on script scope alternates.
                $script:mockCalled = 0
                Mock -ModuleName DOrcDeployModule Get-DorcCredSSPStatus {
                    $script:mockCalled++
                    if ($script:mockCalled % 2 -eq 1) {
                        # Odd run = False
                        return (New-Object psobject -Property @{
                            LocalComputerName             = $Env:COMPUTERNAME
                            RemoteComputerName            = "SERVER1"
                            LocalOS                       = "10.0.17134"
                            RemoteOS                      = "10.0.14393"
                            LocalCredSSPEnabled           = "False"
                            RemoteCredSSPEnabled          = "False"
                            LocalPatchInstalled           = "False"
                            RemotePatchInstalled          = "False"
                            LocalHotFixWorkaroundInPlace  = "False"
                            RemoteHotFixWorkaroundInPlace = "False"
                            CredSSPWorks                  = "False"
                        })
                    }
                    # Even run = True
                    return (New-Object psobject -Property @{
                        LocalComputerName             = $Env:COMPUTERNAME
                        RemoteComputerName            = "SERVER1"
                        LocalOS                       = "10.0.17134"
                        RemoteOS                      = "10.0.14393"
                        LocalCredSSPEnabled           = "True"
                        RemoteCredSSPEnabled          = "True"
                        LocalPatchInstalled           = "True"
                        RemotePatchInstalled          = "True"
                        LocalHotFixWorkaroundInPlace  = "True"
                        RemoteHotFixWorkaroundInPlace = "True"
                        CredSSPWorks                  = "True"
                    })
                }
                # Invoke-Command is called twice: once to probe the remote
                # machine's WSMan config (expects the "configured to receive"
                # string), and once as the wrapper around Enable-WSManCredSSP
                # -Role Server. The Enable path mustn't return any value —
                # Enable-DorcCredSSP pipelines Invoke-Command's output as
                # its return when there's no assignment, so leaking the probe
                # string there pollutes the function's return.
                Mock -ModuleName DOrcDeployModule Invoke-Command `
                    -ParameterFilter { $ScriptBlock -and $ScriptBlock.ToString() -match 'Enable-WSManCredSSP' } `
                    -MockWith { }  # emit nothing — real Enable-WSManCredSSP returns no output
                Mock -ModuleName DOrcDeployModule Invoke-Command {
                    "The machine is configured to allow delegating fresh credentials to the following target(s): SERVER1`nThis computer is configured to receive credentials from a remote client computer."
                }
                Mock -ModuleName DOrcDeployModule Get-MSHotfix {
                    New-Object psobject -Property @{ HotFixID = 'KB12345678' }
                }
            }

            It "Enables CredSSP successfully" {
                $result = Enable-DorcCredSSP -ComputerName SERVER1 -Credential $script:TestCred
                # The function prefixes the success message with "[INFO] " —
                # the original test was written against an earlier version
                # that didn't prefix it.
                $result | Should -Be "[INFO] CredSSP Successfully enabled between [$Env:COMPUTERNAME] and [SERVER1]"
            }
        }
    }

    Context "Computer not reachable" {
        BeforeAll {
            $script:TestCred = New-Object System.Management.Automation.PSCredential(
                'nwtraders\administrator',
                ('mypassword' | ConvertTo-SecureString -AsPlainText -Force))

            Mock -ModuleName DOrcDeployModule Test-IsRunningAsAdministrator { return $true }
            Mock -ModuleName DOrcDeployModule Get-DorcCredSSPStatus {
                New-Object psobject -Property @{
                    LocalComputerName             = $Env:COMPUTERNAME
                    RemoteComputerName            = "SERVER1"
                    LocalOS                       = "10.0.17134"
                    RemoteOS                      = "10.0.14393"
                    LocalCredSSPEnabled           = "True"
                    RemoteCredSSPEnabled          = "False"
                    LocalPatchInstalled           = "False"
                    RemotePatchInstalled          = "False"
                    LocalHotFixWorkaroundInPlace  = "False"
                    RemoteHotFixWorkaroundInPlace = "False"
                    CredSSPWorks                  = "False"
                }
            }
        }

        It "Throws" {
            # The function wraps the underlying Invoke-Command WinRM failure
            # as "[ERROR] Failed to enable CredSSP on remote machine ...".
            # Use a wildcard pattern — Pester's Should -Throw without
            # wildcards does exact-match, not substring.
            { Enable-DorcCredSSP -ComputerName nonexistentcomputer -Credential $script:TestCred } |
                Should -Throw '*Failed*'
        }
    }
}
