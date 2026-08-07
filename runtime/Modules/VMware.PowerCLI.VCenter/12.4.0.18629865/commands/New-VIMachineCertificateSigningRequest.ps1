using namespace VMware.VimAutomation.Sdk.Util10Ps.BaseCmdlet
using namespace VMware.VimAutomation.ViCore.Types.V1
using namespace VMware.VimAutomation.ViCore.Types.V1.Inventory

using module VMware.PowerCLI.Sdk.Types

. (Join-Path $PSScriptRoot "../utils/Connection.ps1")
. (Join-Path $PSScriptRoot "../types/builders/New-ViMachineCertificateRequestInfo.ps1")

<#
.SYNOPSIS

This cmdlet generates a new Certificate Signing Request (CSR) for a vCenter Server system or an ESXi host.

.DESCRIPTION

This cmdlet generates a new Certificate Signing Request (CSR) for a vCenter Server system or an ESXi host.

By default, this cmdlet generates a CSR for a vCenter Server system. If you want to generate a CSR for an ESXi host, you can use the VMHost parameter.

This cmdlet requires a connection to a vCenter Server system through the Connect-VIServer cmdlet.

.PARAMETER VMHost

Specifies an ESXi host for which to generate a CSR.

.PARAMETER Country

Provides input for the country field.

.PARAMETER Email

Provides input for the email field.

.PARAMETER Locality

Provides input for the locality field.

.PARAMETER Organization

Provides input for the organization field.

.PARAMETER OrganizationUnit

Provides input for the organization unit field.

.PARAMETER StateOrProvince

Provides input for the state or province field.

.PARAMETER CommonName

Specifies a common name for the certificate. If not used, the PNID of the vCenter Server system or the FQDN of the ESXi host is the common name.

.PARAMETER KeySize

Specifies the size of the certificate private key. If not used, the default is 2048 bits.

.PARAMETER SubjectAltName

Specifies domain names and IP addresses that you want to secure with the certificate.

.EXAMPLE
PS C:\> $csr = New-VIMachineCertificateSigningRequest `
   -Country 'US' `
   -Email 'example@dummy.com' `
   -Locality 'Example Town' `
   -Organization 'Example organization' `
   -OrganizationUnit 'Example organization unit' `
   -StateOrProvince 'Example state'

Create a certificate signing request for the vCenter Server system and saves it in the csr variable.

.EXAMPLE
PS C:\> $csr = New-VIMachineCertificateSigningRequest 'MyHost'

Create a certificate signing request for the 'MyHost' ESXi host and saves it in the csr variable.

.OUTPUTS

[ViMachineCertificateRequestInfo]

.LINK

https://developer.vmware.com/docs/powercli/latest/vmware.powercli.vcenter/commands/new-vimachinecertificatesigningrequest


#>
function New-VIMachineCertificateSigningRequest {
   [CmdletBinding(
      ConfirmImpact = "Low",
      DefaultParameterSetName = "VCenter",
      SupportsShouldProcess = $False)]
   [OutputType([string])]
   Param (
      [Parameter(
         Mandatory = $true,
         ValueFromPipeline = $true,
         Position = 0,
         ParameterSetName = 'VmHost')]
      [Parameter(
         Mandatory = $true,
         ValueFromPipeline = $true,
         Position = 0,
         ParameterSetName = 'VmHostWithFQDNCommonName')]
      [ObnArgumentTransformation([VMHost])]
      [VMHost]
      $VMHost,

      [Parameter(Mandatory = $true, ParameterSetName = "VCenter")]
      [Parameter(Mandatory = $true, ParameterSetName = 'VmHost')]
      [String]
      $Country,

      [Parameter(Mandatory = $true, ParameterSetName = "VCenter")]
      [String]
      $Email,

      [Parameter(Mandatory = $true, ParameterSetName = "VCenter")]
      [Parameter(Mandatory = $true, ParameterSetName = 'VmHost')]
      [String]
      $Locality,

      [Parameter(Mandatory = $true, ParameterSetName = "VCenter")]
      [Parameter(Mandatory = $true, ParameterSetName = 'VmHost')]
      [String]
      $Organization,

      [Parameter(Mandatory = $true, ParameterSetName = "VCenter")]
      [Parameter(Mandatory = $true, ParameterSetName = 'VmHost')]
      [String]
      $OrganizationUnit,

      [Parameter(Mandatory = $true, ParameterSetName = 'VCenter')]
      [Parameter(Mandatory = $true, ParameterSetName = 'VmHost')]
      [String]
      $StateOrProvince,

      [Parameter(ParameterSetName = 'VCenter')]
      [Parameter(Mandatory = $true, ParameterSetName = 'VmHost')]
      [ValidateNotNullOrEmpty()]
      [String]
      $CommonName,

      [Parameter(ParameterSetName = 'VCenter')]
      [ValidateRange(1, [UInt32]::MaxValue)]
      [UInt32]
      $KeySize,

      [Parameter(ParameterSetName = 'VCenter')]
      [ValidateNotNull()]
      [String[]]
      $SubjectAltName,

      [Parameter()]
      [ObnArgumentTransformation([VIServer], Critical = $true)]
      [VIServer]
      $Server
   )

   Begin {
      # Handle Server obn first
      if($Server) {
         $resolvedServer = Resolve-ObjectByName -Object $Server `
             -Type ([VIServer]) `
             -OneObjectExpected

         $Server = [VIServer] $resolvedServer
      }

      $activeServer = GetActiveServer($Server)
      if (-Not $VMHost) {
         ValidateApiVersionSupported -server $activeServer -major 7 -minor 0
         $apiServer = GetApiServer($activeServer)
      }

      # Collect OBN for parameter 'VMHost'
      if($VMHost) {
         $resolvedVMHost = Resolve-ObjectByName -Object $VMHost `
            -Type ([VMHost]) `
            -CollectorCmdlet 'Get-VMHost' `
            -OneOrMoreObjectsExpected `
            -Server $activeServer

         $VMHost = [VMHost] $resolvedVMHost
      }
   }

   Process {
      # Validate all objects are from the same server
      if($VMHost) {
         $VMHost | ValidateSameServer -ExpectedServer $activeServer
      }

      if ($VMHost) {
         try {
            # VmHost parameter set
            $certificateManager = Get-View $VMHost.ExtensionData.ConfigManager.CertificateManager -Server $activeServer

            if ($CommonName) {
               # Common name becomes mandatory if any part of the distinguished name is specified.
               $dn = "CN = $CommonName, C = $Country, L = $Locality, ST = $StateOrProvince, O = $Organization, OU = $OrganizationUnit"

               $certificateManager.GenerateCertificateSigningRequestByDn($dn) | `
                  New-ViMachineCertificateRequestInfo -TargetESXi $VMHost | `
                  Write-Output
            } else {
               $certificateManager.GenerateCertificateSigningRequest($false) | `
                  New-ViMachineCertificateRequestInfo -TargetESXi $VMHost | `
                  Write-Output
            }
         } catch {
            Write-PowerCLIError `
               -ErrorObject $_ `
               -ErrorId "PowerCLI_New-VIMachineCertificateSigningRequest_FailedToGenerateVMHostCSR"
         }
      } else {
         # VCenter parameter set
         try {
            $specParams = @{
               Organization = $Organization;
               OrganizationUnit = $OrganizationUnit;
               Locality = $Locality;
               StateOrProvince = $StateOrProvince;
               Country = $Country;
               EmailAddress = $Email
            }
            if ($KeySize -gt 0) {
               $specParams['KeySize'] = $KeySize
            }
            if (-not [string]::IsNullOrEmpty($CommonName)) {
               $specParams['CommonName'] = $CommonName
            }
            if ($SubjectAltName) {
               $specParams['SubjectAltName'] = $SubjectAltName
            }

            $vcCSR = Initialize-CertificateManagementVcenterTlsCsrSpec @specParams | `
            Invoke-CreateCertificateManagementTlsCsr -Server $apiServer -ErrorAction:Stop

            $vcCSR.csr | `
               New-ViMachineCertificateRequestInfo -TargetVC $activeServer | `
               Write-Output
         } catch {
            Write-PowerCLIError `
               -ErrorObject $_ `
               -ErrorId "PowerCLI_New-VIMachineCertificateSigningRequest_FailedToGenerateVCenterCSR"
         }
      }
   }
}

# SIG # Begin signature block
# MIIrIAYJKoZIhvcNAQcCoIIrETCCKw0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBmB5AMd2OuNYS6
# VJ3BQOPBtOykqhzM5/EWigeL67I1cKCCDdowggawMIIEmKADAgECAhAIrUCyYNKc
# TJ9ezam9k67ZMA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0z
# NjA0MjgyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQDVtC9C0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0
# JAfhS0/TeEP0F9ce2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJr
# Q5qZ8sU7H/Lvy0daE6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhF
# LqGfLOEYwhrMxe6TSXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+F
# LEikVoQ11vkunKoAFdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh
# 3K3kGKDYwSNHR7OhD26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJ
# wZPt4bRc4G/rJvmM1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQay
# g9Rc9hUZTO1i4F4z8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbI
# YViY9XwCFjyDKK05huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchAp
# QfDVxW0mdmgRQRNYmtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRro
# OBl8ZhzNeDhFMJlP/2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IB
# WTCCAVUwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+
# YXsIiGX0TkIwHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0P
# AQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAC
# hjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAED
# MAgGBmeBDAEEATANBgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql
# +Eg08yy25nRm95RysQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFF
# UP2cvbaF4HZ+N3HLIvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1h
# mYFW9snjdufE5BtfQ/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3Ryw
# YFzzDaju4ImhvTnhOE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5Ubdld
# AhQfQDN8A+KVssIhdXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw
# 8MzK7/0pNVwfiThV9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnP
# LqR0kq3bPKSchh/jwVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatE
# QOON8BUozu3xGFYHKi8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bn
# KD+sEq6lLyJsQfmCXBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQji
# WQ1tygVQK+pKHJ6l/aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbq
# yK+p/pQd52MbOoZWeE4wggciMIIFCqADAgECAhAOxvKydqFGoH0ObZNXteEIMA0G
# CSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTEwHhcNMjEwODEwMDAwMDAwWhcNMjMwODEw
# MjM1OTU5WjCBhzELMAkGA1UEBhMCVVMxEzARBgNVBAgTCkNhbGlmb3JuaWExEjAQ
# BgNVBAcTCVBhbG8gQWx0bzEVMBMGA1UEChMMVk13YXJlLCBJbmMuMRUwEwYDVQQD
# EwxWTXdhcmUsIEluYy4xITAfBgkqhkiG9w0BCQEWEm5vcmVwbHlAdm13YXJlLmNv
# bTCCAaIwDQYJKoZIhvcNAQEBBQADggGPADCCAYoCggGBAMD6lJG8OWkM12huIQpO
# /q9JnhhhW5UyW9if3/UnoFY3oqmp0JYX/ZrXogUHYXmbt2gk01zz2P5Z89mM4gqR
# bGYC2tx+Lez4GxVkyslVPI3PXYcYSaRp39JsF3yYifnp9R+ON8O3Gf5/4EaFmbeT
# ElDCFBfExPMqtSvPZDqekodzX+4SK1PIZxCyR3gml8R3/wzhb6Li0mG7l0evQUD0
# FQAbKJMlBk863apeX4ALFZtrnCpnMlOjRb85LsjV5Ku4OhxQi1jlf8wR+za9C3DU
# ki60/yiWPu+XXwEUqGInIihECBbp7hfFWrnCCaOgahsVpgz8kKg/XN4OFq7rbh4q
# 5IkTauqFhHaE7HKM5bbIBkZ+YJs2SYvu7aHjw4Z8aRjaIbXhI1G+NtaNY7kSRrE4
# fAyC2X2zV5i4a0AuAMM40C1Wm3gTaNtRTHnka/pbynUlFjP+KqAZhOniJg4AUfjX
# sG+PG1LH2+w/sfDl1A8liXSZU1qJtUs3wBQFoSGEaGBeDQIDAQABo4ICJTCCAiEw
# HwYDVR0jBBgwFoAUaDfg67Y7+F8Rhvv+YXsIiGX0TkIwHQYDVR0OBBYEFIhC+HL9
# QlvsWsztP/I5wYwdfCFNMB0GA1UdEQQWMBSBEm5vcmVwbHlAdm13YXJlLmNvbTAO
# BgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAwwCgYIKwYBBQUHAwMwgbUGA1UdHwSBrTCB
# qjBToFGgT4ZNaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3Rl
# ZEc0Q29kZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5jcmwwU6BRoE+GTWh0
# dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVTaWdu
# aW5nUlNBNDA5NlNIQTM4NDIwMjFDQTEuY3JsMD4GA1UdIAQ3MDUwMwYGZ4EMAQQB
# MCkwJwYIKwYBBQUHAgEWG2h0dHA6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCBlAYI
# KwYBBQUHAQEEgYcwgYQwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0
# LmNvbTBcBggrBgEFBQcwAoZQaHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0MDk2U0hBMzg0MjAyMUNBMS5j
# cnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEACQAYaQI6Nt2KgxdN
# 6qqfcHB33EZRSXkvs8O9iPZkdDjEx+2fgbBPLUvk9A7T8mRw7brbcJv4PLTYJDFo
# c5mlcmG7/5zwTOuIs2nBGXc/uxCnyW8p7kD4Y0JxPKEVQoIQ8lJS9Uy/hBjyakeV
# ef982JyzvDbOlLBy6AS3ZpXVkRY5y3Va+3v0R/0xJ+JRxUicQhiZRidq2TCiWEas
# d+tLL6jrKaBO+rmP52IM4eS9d4Yids7ogKEBAlJi0NbvuKO0CkgOlFjp1tOvD4sQ
# taHIMmqi40p4Tjyf/sY6yGjROXbMeeF1vlwbBAASPWpQuEIxrNHoVN30YfJyuOWj
# zdiJUTpeLn9XdjM3UlhfaHP+oIAKcmkd33c40SFRlQG9+P9Wlm7TcPxGU4wzXI8n
# Cw/h235jFlAAiWq9L2r7Un7YduqsheJVpGoXmRXJH0T2G2eNFS5/+2sLn98kN2Cn
# J7j6C242onjkZuGL2/+gqx8m5Jbpu9P4IAeTC1He/mX9j6XpIu+7uBoRVwuWD1i0
# N5SiUz7Lfnbr6Q1tHMXKDLFdwVKZos2AKEZhv4SU0WvenMJKDgkkhVeHPHbTahQf
# P1MetR8tdRs7uyTWAjPK5xf5DLEkXbMrUkpJ089fPvAGVHBcHRMqFA5egexOb6sj
# tKncUjJ1xAAtAExGdCh6VD2U5iYxghycMIIcmAIBATB9MGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBDb2RlIFNpZ25pbmcgUlNBNDA5NiBTSEEzODQgMjAyMSBDQTECEA7G
# 8rJ2oUagfQ5tk1e14QgwDQYJYIZIAWUDBAIBBQCggZYwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwKgYKKwYB
# BAGCNwIBDDEcMBqhGIAWaHR0cDovL3d3dy52bXdhcmUuY29tLzAvBgkqhkiG9w0B
# CQQxIgQgW6yaX541w+DXHSRYNhJAyW6DeavEgXPFB9IxKXrgUHEwDQYJKoZIhvcN
# AQEBBQAEggGALOJT+omI2PX1Tcjb1vgS30AE9lzi49400SCh37C33N9nYVjfDCgl
# WvXghugiUywsCxWsDMi9ShCM8ojdxZ/9wRnA3GnGw9amH1/3mnt96qLTGwzPmJoK
# TeMlZ0Kddaqx9+n8hB51rf2+ypw9/UyBEIb6TO6UbzRHB4c0c2irG7aQg3M20Ikr
# MutzGCzP01kqwwNoqZcVZNEcYTNawji7eFU/J7gXR5MlLJipCuwGAKFVuWdyuwsf
# GCOkuJNtC0jENVQjMh4Jwqn2tJpEvNQVTJHonaHyYMAxXGh9A0q6/erNnw8xiUNS
# lmX/slEVObwF9je4/8z5AiaU8bu9zTly4BKfy0yHFV1wuWl1j9vmIG1bL8HIg9mv
# 4ibWoWLSWaDAu0NAAPwdlWRmYzWmP7mtjbyoisF1zw0EyLZYBRus6/nY/PBJy9Mz
# DvtoQD9oq7VW8Q57NcOW6Iuyo4ko98yKKD9DRNgWvAwO7l3G1b6mW+LAVF0mU/X2
# JKd6oYhXDx5ZoYIZ1zCCGdMGCisGAQQBgjcDAwExghnDMIIZvwYJKoZIhvcNAQcC
# oIIZsDCCGawCAQMxDzANBglghkgBZQMEAgEFADCB3AYLKoZIhvcNAQkQAQSggcwE
# gckwgcYCAQEGCSsGAQQBoDICAzAxMA0GCWCGSAFlAwQCAQUABCB5uimcfuWJ9aLh
# tlGgEVNdUoeTPQVzc5gMbKdSJp4wyQIUW+90vwbzxKEtjt4QfXysmqzPEUEYDzIw
# MjEwOTE0MjMwNDU3WjADAgEBoFekVTBTMQswCQYDVQQGEwJCRTEZMBcGA1UECgwQ
# R2xvYmFsU2lnbiBudi1zYTEpMCcGA1UEAwwgR2xvYmFsc2lnbiBUU0EgZm9yIEFk
# dmFuY2VkIC0gRzSgghVkMIIGVTCCBD2gAwIBAgIQAQBGaVCmBKnZcOgd0k1BnzAN
# BgkqhkiG9w0BAQsFADBbMQswCQYDVQQGEwJCRTEZMBcGA1UEChMQR2xvYmFsU2ln
# biBudi1zYTExMC8GA1UEAxMoR2xvYmFsU2lnbiBUaW1lc3RhbXBpbmcgQ0EgLSBT
# SEEzODQgLSBHNDAeFw0yMTA1MjcwOTU1MjNaFw0zMjA2MjgwOTU1MjJaMFMxCzAJ
# BgNVBAYTAkJFMRkwFwYDVQQKDBBHbG9iYWxTaWduIG52LXNhMSkwJwYDVQQDDCBH
# bG9iYWxzaWduIFRTQSBmb3IgQWR2YW5jZWQgLSBHNDCCAaIwDQYJKoZIhvcNAQEB
# BQADggGPADCCAYoCggGBAN8waZh7lw1uo1S0OV9kWXUEIv5OaW/oF3R0pX1RGA5G
# IB9oLrrlZdbJ0pGh7KT3Veqq7TvM+2KbhEKSejJM+fTyHPiY0bkxgSWtrVZnMqb+
# hzLzXMMtYvFNiQw3tH/rKuNgi29sPTHy7cldgJspnVXg4sT/6naTGB5sqat7yR4S
# YdXA56Dm+JApMyy4v25ast3HB0PRO9swC7R4w+zq8aJUz2CTOMz3ZEP1zwgEFnDI
# tNsO1AqKCNy7k8EdbvKMnOshNZ7/j7ywfsKEOH7mnWR6JqDxILG84dgqJZ0YUuRt
# 1EwwCnjMLUaO7VcLP3mVUKcDsDODMrdAnvS0kpcTDFC3nqq0QU4LmInM+8QhRJAy
# jkjyLEsMF+SEV1umrPuXg/mNZFTC7GpDHs8KdpKyEL/t1qMgD7XRMI4aQLE259CO
# ePMTwC8LiJA7CGHjD61Hsw5UcJV/oEPUWsbdF5+UywCHaA7hrpPuLHIEGzIXkEvX
# K4AlBR/lM/TowGgqeReg7wIDAQABo4IBmzCCAZcwDgYDVR0PAQH/BAQDAgeAMBYG
# A1UdJQEB/wQMMAoGCCsGAQUFBwMIMB0GA1UdDgQWBBSufnCBeCAUKa3yePhZANnM
# piQCjjBMBgNVHSAERTBDMEEGCSsGAQQBoDIBHjA0MDIGCCsGAQUFBwIBFiZodHRw
# czovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzAJBgNVHRMEAjAAMIGQ
# BggrBgEFBQcBAQSBgzCBgDA5BggrBgEFBQcwAYYtaHR0cDovL29jc3AuZ2xvYmFs
# c2lnbi5jb20vY2EvZ3N0c2FjYXNoYTM4NGc0MEMGCCsGAQUFBzAChjdodHRwOi8v
# c2VjdXJlLmdsb2JhbHNpZ24uY29tL2NhY2VydC9nc3RzYWNhc2hhMzg0ZzQuY3J0
# MB8GA1UdIwQYMBaAFOoWxmnn48tXRTkzpPBAvtDDvWWWMEEGA1UdHwQ6MDgwNqA0
# oDKGMGh0dHA6Ly9jcmwuZ2xvYmFsc2lnbi5jb20vY2EvZ3N0c2FjYXNoYTM4NGc0
# LmNybDANBgkqhkiG9w0BAQsFAAOCAgEAf2Lo+tl3L0Jvaw/X3UVZPPR1egDsvfZv
# DiNtLTNCchRPRJSBveuMAohMrH/HXc23xCSau5kBaApa6kVh07As132gF+5dgPEa
# 4uf8sd8dMgQoDzaE1wlGLbZ+wEAVIhp5YWeXthKP0E9mLC5UKlgGrJlO/XWtVCYK
# aP+SJ/g8uRltMIEmTIUs83Pcj+DlymRKe0cRTNqi1Lfx5FF65jmwIQcZ4PCMuXFw
# fZHtNJ+LMZ4NxMY+Nitm1sBB1bIjSSLTvl+JNoxa1sVQqj8OTlQJtv4Nkdlx2J82
# PDSOiYO35PNmSs43kItdeuo+o+MHBz2UGRSe+rFnS+u2srcIb5KWRV1M7g5ZWotm
# c2FFNkqGzmNDGW4GOglGOZB0xnMLXkLRzS8ibCQnpwICUZKNAbRdhcf4w0F13WSM
# 8vOY7um3hwmnvQoTMDdiH1nnKXJ3aXV4kLDNHDpcahCGcvcAsjKXWXieTvizZv2v
# K/yJtnWilAo3khNBdd31Pzqup6i0QtPZnFES8vJ61ivsnkwl2W2ckfQfAU9Ix+yP
# +Vuq7PpcEXJgruw3cZS+XEmJTClt81c7GgBXvL6QLkJhgtXf/wCBlnwBVZO4YmTo
# BoarVUpvM8Xz2lgFjd0B9TxVIYX+ezV5xX+y+9itvZ35VQokZHRhiiuXNl9WvfLX
# 4Ox8/fnrktQwggZZMIIEQaADAgECAg0B7BySQN79LkBdfEd0MA0GCSqGSIb3DQEB
# DAUAMEwxIDAeBgNVBAsTF0dsb2JhbFNpZ24gUm9vdCBDQSAtIFI2MRMwEQYDVQQK
# EwpHbG9iYWxTaWduMRMwEQYDVQQDEwpHbG9iYWxTaWduMB4XDTE4MDYyMDAwMDAw
# MFoXDTM0MTIxMDAwMDAwMFowWzELMAkGA1UEBhMCQkUxGTAXBgNVBAoTEEdsb2Jh
# bFNpZ24gbnYtc2ExMTAvBgNVBAMTKEdsb2JhbFNpZ24gVGltZXN0YW1waW5nIENB
# IC0gU0hBMzg0IC0gRzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDw
# AuIwI/rgG+GadLOvdYNfqUdSx2E6Y3w5I3ltdPwx5HQSGZb6zidiW64HiifuV6PE
# Ne2zNMeswwzrgGZt0ShKwSy7uXDycq6M95laXXauv0SofEEkjo+6xU//NkGrpy39
# eE5DiP6TGRfZ7jHPvIo7bmrEiPDul/bc8xigS5kcDoenJuGIyaDlmeKe9JxMP11b
# 7Lbv0mXPRQtUPbFUUweLmW64VJmKqDGSO/J6ffwOWN+BauGwbB5lgirUIceU/kKW
# O/ELsX9/RpgOhz16ZevRVqkuvftYPbWF+lOZTVt07XJLog2CNxkM0KvqWsHvD9WZ
# uT/0TzXxnA/TNxNS2SU07Zbv+GfqCL6PSXr/kLHU9ykV1/kNXdaHQx50xHAotIB7
# vSqbu4ThDqxvDbm19m1W/oodCT4kDmcmx/yyDaCUsLKUzHvmZ/6mWLLU2EESwVX9
# bpHFu7FMCEue1EIGbxsY1TbqZK7O/fUF5uJm0A4FIayxEQYjGeT7BTRE6giunUln
# EYuC5a1ahqdm/TMDAd6ZJflxbumcXQJMYDzPAo8B/XLukvGnEt5CEk3sqSbldwKs
# DlcMCdFhniaI/MiyTdtk8EWfusE/VKPYdgKVbGqNyiJc9gwE4yn6S7Ac0zd0hNkd
# Zqs0c48efXxeltY9GbCX6oxQkW2vV4Z+EDcdaxoU3wIDAQABo4IBKTCCASUwDgYD
# VR0PAQH/BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFOoWxmnn
# 48tXRTkzpPBAvtDDvWWWMB8GA1UdIwQYMBaAFK5sBaOTE+Ki5+LXHNbH8H/IZ1Og
# MD4GCCsGAQUFBwEBBDIwMDAuBggrBgEFBQcwAYYiaHR0cDovL29jc3AyLmdsb2Jh
# bHNpZ24uY29tL3Jvb3RyNjA2BgNVHR8ELzAtMCugKaAnhiVodHRwOi8vY3JsLmds
# b2JhbHNpZ24uY29tL3Jvb3QtcjYuY3JsMEcGA1UdIARAMD4wPAYEVR0gADA0MDIG
# CCsGAQUFBwIBFiZodHRwczovL3d3dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5
# LzANBgkqhkiG9w0BAQwFAAOCAgEAf+KI2VdnK0JfgacJC7rEuygYVtZMv9sbB3DG
# +wsJrQA6YDMfOcYWaxlASSUIHuSb99akDY8elvKGohfeQb9P4byrze7AI4zGhf5L
# FST5GETsH8KkrNCyz+zCVmUdvX/23oLIt59h07VGSJiXAmd6FpVK22LG0LMCzDRI
# RVXd7OlKn14U7XIQcXZw0g+W8+o3V5SRGK/cjZk4GVjCqaF+om4VJuq0+X8q5+dI
# ZGkv0pqhcvb3JEt0Wn1yhjWzAlcfi5z8u6xM3vreU0yD/RKxtklVT3WdrG9KyC5q
# ucqIwxIwTrIIc59eodaZzul9S5YszBZrGM3kWTeGCSziRdayzW6CdaXajR63Wy+I
# Lj198fKRMAWcznt8oMWsr1EG8BHHHTDFUVZg6HyVPSLj1QokUyeXgPpIiScseeI8
# 5Zse46qEgok+wEr1If5iEO0dMPz2zOpIJ3yLdUJ/a8vzpWuVHwRYNAqJ7YJQ5NF7
# qMnmvkiqK1XZjbclIA4bUaDUY6qD6mxyYUrJ+kPExlfFnbY8sIuwuRwx773vFNgU
# QGwgHcIt6AvGjW2MtnHtUiH+PvafnzkarqzSL3ogsfSsqh3iLRSd+pZqHcY8yvPZ
# HL9TTaRHWXyVxENB+SXiLBB+gfkNlKd98rUJ9dhgckBQlSDUQ0S++qCV5yBZtnjG
# pGqqIpswggVHMIIEL6ADAgECAg0B8kBCQM79ItvpbHH8MA0GCSqGSIb3DQEBDAUA
# MEwxIDAeBgNVBAsTF0dsb2JhbFNpZ24gUm9vdCBDQSAtIFIzMRMwEQYDVQQKEwpH
# bG9iYWxTaWduMRMwEQYDVQQDEwpHbG9iYWxTaWduMB4XDTE5MDIyMDAwMDAwMFoX
# DTI5MDMxODEwMDAwMFowTDEgMB4GA1UECxMXR2xvYmFsU2lnbiBSb290IENBIC0g
# UjYxEzARBgNVBAoTCkdsb2JhbFNpZ24xEzARBgNVBAMTCkdsb2JhbFNpZ24wggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCVB+hzymb57BTKezz3DQjxtEUL
# LIK0SMbrWzyug7hBkjMUpG9/6SrMxrCIa8W2idHGsv8UzlEUIexK3RtaxtaH7k06
# FQbtZGYLkoDKRN5zlE7zp4l/T3hjCMgSUG1CZi9NuXkoTVIaihqAtxmBDn7Eirxk
# TCEcQ2jXPTyKxbJm1ZCatzEGxb7ibTIGph75ueuqo7i/voJjUNDwGInf5A959eqi
# HyrScC5757yTu21T4kh8jBAHOP9msndhfuDqjDyqtKT285VKEgdt/Yyyic/QoGF3
# yFh0sNQjOvddOsqi250J3l1ELZDxgc1Xkvp+vFAEYzTfa5MYvms2sjnkrCQ2t/Dv
# thwTV5O23rL44oW3c6K4NapF8uCdNqFvVIrxclZuLojFUUJEFZTuo8U4lptOTloL
# R/MGNkl3MLxxN+Wm7CEIdfzmYRY/d9XZkZeECmzUAk10wBTt/Tn7g/JeFKEEsAvp
# /u6P4W4LsgizYWYJarEGOmWWWcDwNf3J2iiNGhGHcIEKqJp1HZ46hgUAntuA1iX5
# 3AWeJ1lMdjlb6vmlodiDD9H/3zAR+YXPM0j1ym1kFCx6WE/TSwhJxZVkGmMOeT31
# s4zKWK2cQkV5bg6HGVxUsWW2v4yb3BPpDW+4LtxnbsmLEbWEFIoAGXCDeZGXkdQa
# J783HjIH2BRjPChMrwIDAQABo4IBJjCCASIwDgYDVR0PAQH/BAQDAgEGMA8GA1Ud
# EwEB/wQFMAMBAf8wHQYDVR0OBBYEFK5sBaOTE+Ki5+LXHNbH8H/IZ1OgMB8GA1Ud
# IwQYMBaAFI/wS3+oLkUkrk1Q+mOai97i3Ru8MD4GCCsGAQUFBwEBBDIwMDAuBggr
# BgEFBQcwAYYiaHR0cDovL29jc3AyLmdsb2JhbHNpZ24uY29tL3Jvb3RyMzA2BgNV
# HR8ELzAtMCugKaAnhiVodHRwOi8vY3JsLmdsb2JhbHNpZ24uY29tL3Jvb3QtcjMu
# Y3JsMEcGA1UdIARAMD4wPAYEVR0gADA0MDIGCCsGAQUFBwIBFiZodHRwczovL3d3
# dy5nbG9iYWxzaWduLmNvbS9yZXBvc2l0b3J5LzANBgkqhkiG9w0BAQwFAAOCAQEA
# SaxexYPzWsthKk2XShUpn+QUkKoJ+cR6nzUYigozFW1yhyJOQT9tCp4YrtviX/yV
# 0SyYFDuOwfA2WXnzjYHPdPYYpOThaM/vf2VZQunKVTm808Um7nE4+tchAw+3Ttlb
# YGpDtH0J0GBh3artAF5OMh7gsmyePLLCu5jTkHZqaa0a3KiJ2lhP0sKLMkrOVPs4
# 6TsHC3UKEdsLfCUn8awmzxFT5tzG4mE1MvTO3YPjGTrrwmijcgDIJDxOuFM8sRer
# 5jUs+dNCKeZfYAOsQmGmsVdqM0LfNTGGyj43K9rE2iT1ThLytrm3R+q7IK1hFreg
# M+Mtiae8szwBfyMagAk06TCCA18wggJHoAMCAQICCwQAAAAAASFYUwiiMA0GCSqG
# SIb3DQEBCwUAMEwxIDAeBgNVBAsTF0dsb2JhbFNpZ24gUm9vdCBDQSAtIFIzMRMw
# EQYDVQQKEwpHbG9iYWxTaWduMRMwEQYDVQQDEwpHbG9iYWxTaWduMB4XDTA5MDMx
# ODEwMDAwMFoXDTI5MDMxODEwMDAwMFowTDEgMB4GA1UECxMXR2xvYmFsU2lnbiBS
# b290IENBIC0gUjMxEzARBgNVBAoTCkdsb2JhbFNpZ24xEzARBgNVBAMTCkdsb2Jh
# bFNpZ24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDMJXaQeQZ4Ihb1
# wIO2hMoonv0FdhHFrYhy/EYCQ8eyip0EXyTLLkvhYIJG4VKrDIFHcGzdZNHr9Syj
# D4I9DCuul9e2FIYQebs7E4B3jAjhSdJqYi8fXvqWaN+JJ5U4nwbXPsnLJlkNc96w
# yOkmDoMVxu9bi9IEYMpJpij2aTv2y8gokeWdimFXN6x0FNx04Druci8unPvQu7/1
# PQDhBjPogiuuU6Y6FnOM3UEOIDrAtKeh6bJPkC4yYOlXy7kEkmho5TgmYHWyn3f/
# kRTvriBJ/K1AFUjRAjFhGV64l++td7dkmnq/X8ET75ti+w1s4FRpFqkD2m7pg5Nx
# dsZphYIXAgMBAAGjQjBAMA4GA1UdDwEB/wQEAwIBBjAPBgNVHRMBAf8EBTADAQH/
# MB0GA1UdDgQWBBSP8Et/qC5FJK5NUPpjmove4t0bvDANBgkqhkiG9w0BAQsFAAOC
# AQEAS0DbwFCq/sgM7/eWVEVJu5YACUGssxOGhigHM8pr5nS5ugAtrqQK0/Xx8Q+K
# v3NnSoPHRHt44K9ubG8DKY4zOUXDjuS5V2yq/BKW7FPGLeQkbLmUY/vcU2hnVj6D
# uM81IcPJaP7O2sJTqsyQiunwXUaMld16WCgaLx3ezQA3QY/tRG3XUyiXfvNnBB4V
# 14qWtNPeTCekTBtzc3b0F5nCH3oO4y0IrQocLP88q1UOD5F+NuvDV0m+4S4tfGCL
# w0FREyOdzvcya5QBqJnnLDMfOjsl0oZAzjsshnjJYS8Uuu7bVW/fhO4FCU29KNhy
# ztNiUGUe65KXgzHZs7XKR1g/XzGCA00wggNJAgEBMG8wWzELMAkGA1UEBhMCQkUx
# GTAXBgNVBAoTEEdsb2JhbFNpZ24gbnYtc2ExMTAvBgNVBAMTKEdsb2JhbFNpZ24g
# VGltZXN0YW1waW5nIENBIC0gU0hBMzg0IC0gRzQCEAEARmlQpgSp2XDoHdJNQZ8w
# DQYJYIZIAWUDBAIBBQCgggEvMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAt
# BgkqhkiG9w0BCTQxIDAeMA0GCWCGSAFlAwQCAQUAoQ0GCSqGSIb3DQEBCwUAMC8G
# CSqGSIb3DQEJBDEiBCCS+hG4fKbMszHZzvucPAHnBDIjI2qQdMBAHMlJzc/ADDCB
# sAYLKoZIhvcNAQkQAi8xgaAwgZ0wgZowgZcEIBPW6cQg/21OJ1RyjGjneIJlZGfb
# mhkPgWWX9n+2zMb5MHMwX6RdMFsxCzAJBgNVBAYTAkJFMRkwFwYDVQQKExBHbG9i
# YWxTaWduIG52LXNhMTEwLwYDVQQDEyhHbG9iYWxTaWduIFRpbWVzdGFtcGluZyBD
# QSAtIFNIQTM4NCAtIEc0AhABAEZpUKYEqdlw6B3STUGfMA0GCSqGSIb3DQEBCwUA
# BIIBgL2sx7Wau9Y0t/ORldPQsX6EHUoQGxajnjwrdvCHeEUcr/0rvlwj5qFfiFMD
# XwGXVWjGXnQ66Rh6NldPxi/qRnzQwByz+2myLG35nwVAZAbj0qIE1IcAlDKKOlN6
# bAAzsCM8Kc8LQ/bjw1QJhXKaRc9qtvACEgC81OR6iPaYfM4O+KwVv5jLx8fKSiwN
# dnJhux8yU0mLObbUwwGWZeHqtEMzesGWvYHOOKjjmZObfEiaTrIvkDLYbGjzcOOz
# BaIXJ/k0KMdVZze1KhrC6JpfcfTf2Y4Qr4ddN4a5ecxrLYaCqA7j9xIh/Mf/Eu5a
# GoPiATW5yhg+t8Ez4MibI/p2EcE0quYrjSW/uUavgoBbKeN1RTM0FrYgr3fV+rQp
# tsYsSbSpPhfSg2inMapoY40OFhxNw1LWEXydCUJg+GfLR2cO6i0n4ULZfPhha47I
# uZ9WMCjxIy6COWSZyqIWIVye5aHIUXXitt7VAN5JgfGSWknpGX6a3+314vnzRuX/
# h8Hxbg==
# SIG # End signature block
