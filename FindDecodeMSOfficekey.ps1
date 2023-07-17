<#
.SYNOPSIS
Get the Microsoft Office product keys.

.DESCRIPTION
This function searches the registry of one or more computers for Microsoft Office product keys and decodes them using a base24 algorithm.

.PARAMETER ComputerName
An array of computer names to search for Office keys. The default value is the local computer.

.PARAMETER Hive
A string that specifies the root key to search for Office keys. The valid values are HKCR, HKCU, HKLM, HKU, and HKCC. The default value is HKLM.

.PARAMETER Path
A string that specifies the registry path to search for Office keys. The default value is "SOFTWARE\Microsoft\Office".

.PARAMETER ValueName
A string that specifies the registry value name that contains the Office key data. The default value is "DigitalProductId".

.PARAMETER OutputFormat
A string that specifies the output format of the function result. The valid values are "Object", "CSV", or "Table". The default value is "Object".

.PARAMETER Verbose
A switch that enables verbose messages.

.PARAMETER Debug
A switch that enables debug messages.

.EXAMPLE
Get-MSOfficeProductKey

This example searches for Office keys in the HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office path on the local computer and returns an array of custom objects.

.EXAMPLE
Get-MSOfficeProductKey -ComputerName "Server01","Server02" -Hive HKCU -Path "Software\MyApp" -OutputFormat Table -Verbose

This example searches for Office keys in the HKEY_CURRENT_USER\Software\MyApp path on two remote computers and returns a formatted table with verbose messages.

.NOTES
This function only works for Office versions that store the product key in the DigitalProductId value in the registry, such as Office 2003, 2007, 2010, and 2013. It does not work for Office versions that use online activation or subscription-based activation, such as Office 2016, 2019, or 365.
This function requires administrative privileges to access the remote registry.
This function uses WMI to access the remote registry, which depends on the network and firewall settings of the remote computers.
#>
function Get-MSOfficeProductKey {
    [CmdletBinding()]
    param(
    [Parameter(ValueFromPipeline=$true)]
    [string[]]$computerName = ".",
    [string]$path = "SOFTWARE\Microsoft\Office",
    [string]$valueName = "DigitalProductId",
    [ValidateSet("Object","CSV","Table")]
    [string]$outputFormat = "Object"
    )

    Begin {
        $hives = @("HKCR","HKCU","HKLM","HKU","HKCC")

        $product = @()

        Function Decode-Key {
            param([byte []] $key)
            $KeyOutput=""
            $KeyOffset = 52
            $IsWin8 = ([System.Math]::Truncate($key[66] / 6)) -band 1
            $key[66] = ($Key[66] -band 0xF7) -bor (($isWin8 -band 2) * 4)
            $i = 24
            $maps = "BCDFGHJKMPQRTVWXY2346789"
            Do {
                $current= 0
                $j = 14
                Do {
                    $current = $current* 256
                    $current = $Key[$j + $KeyOffset] + $Current
                    $Key[$j + $KeyOffset] = [System.Math]::Truncate($Current / 24 )
                    $Current=$Current % 24
                    $j--
                } while ($j -ge 0)
                $i--
                $KeyOutput = $Maps.Substring($Current, 1) + $KeyOutput
                $last = $current
            } while ($i -ge 0)
            If ($isWin8 -eq 1) {
                $keypart1 = $KeyOutput.Substring(1,$last)
                $insert = "N"
                $KeyOutput = $KeyOutput.Replace($keypart1, $keypart1 + $insert)
                if ($Last -eq 0) {
                    $KeyOutput = $insert + $KeyOutput
                }
            }
            if ($keyOutput.Length -eq 26) {
                $result = [String]::Format("{0}- {1}- {2}- {3}- {4}",$KeyOutput.Substring(1,5),$KeyOutput.Substring(6,5),$KeyOutput.Substring(11,5),$KeyOutput.Substring(16,5),$KeyOutput.Substring(21,5))
            }
            else {
                $KeyOutput
            }
            return $result
        }
    }

    Process {
        foreach ($computer in $computerName) {

            try {
                $wmi = [WMIClass]"\\$computer\root\default:stdRegProv"
                Write-Verbose "Connected to registry on ${computer}" -Verbose:$Verbose
            }
            catch {
                Write-Error "Failed to connect to registry on ${computer}: $_"
                continue
            }

            foreach ($hive in $hives) {

                Write-Debug "Searching for Office keys on ${computer} in ${hive}\$path" -Debug:$Debug

                switch ($hive) {
                    "HKCR" {$hiveValue = 2147483648}
                    "HKCU" {$hiveValue = 2147483649}
                    "HKLM" {$hiveValue = 2147483650}
                    "HKU" {$hiveValue = 2147483651}
                    "HKCC" {$hiveValue = 2147483653}
                }

                try {
                    # Use a try-catch block to handle any errors when accessing a hive that does not exist on the remote computer
                    # For example, HKU may not exist on some computers
                    # In that case, skip the hive and continue with the next one
                    $subkeys1 = $wmi.EnumKey($hiveValue,$path)
                    foreach ($subkey1 in $subkeys1.snames) {
                        if ($subkey1 -match "^\d+\.\d+$") {
                            $subkeys2 = $wmi.EnumKey($hiveValue,"$path\$subkey1")
                            foreach ($subkey2 in $subkeys2.snames) {
                                if ($subkey2 -eq "Registration") {
                                    $subkeys3 = $wmi.EnumKey($hiveValue,"$path\$subkey1\$subkey2")
                                    foreach ($subkey3 in $subkeys3.snames) {
                                        $productName = $wmi.GetStringValue($hiveValue,"$path\$subkey1\$subkey2\$subkey3","ProductName")
                                        $productKey = $wmi.GetBinaryValue($hiveValue,"$path\$subkey1\$subkey2\$subkey3","DigitalProductId")
                                        if ($productName.ReturnValue -eq 0 -and $productKey.ReturnValue -eq 0) {
                                            $decodedKey = Decode-Key $productKey.uValue
                                            $product += [PSCustomObject]@{
                                                ComputerName = $computer
                                                Hive = $hive
                                                ProductName = $productName.sValue
                                                ProductKey = $decodedKey
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Failed to access hive ${hive} on computer ${computer}: $_"
                    continue
                }
            }

            switch ($outputFormat) {
                "Object" {$product}
                "CSV" {$product | ConvertTo-Csv -NoTypeInformation}
                "Table" {$product | Format-Table -AutoSize}
            }
        }
    }
}
