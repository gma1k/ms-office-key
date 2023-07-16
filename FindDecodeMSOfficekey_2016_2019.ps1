function Get-MSOfficeProductKey {
    [CmdletBinding()]
    param(
    [Parameter(ValueFromPipeline=$true)]
    [string[]]$computerName = ".",
    [string]$path = "SOFTWARE\Microsoft\Office\16.0\Registration",
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
                Write-Verbose "Connected to registry on ${computer}"
            }
            catch {
                Write-Error "Failed to connect to registry on ${computer}: $_"
                continue
            }

            foreach ($hive in $hives) {

                Write-Debug "Searching for Office keys on ${computer} in ${hive}\$path"

                switch ($hive) {
                    "HKCR" {$hiveValue = 2147483648}
                    "HKCU" {$hiveValue = 2147483649}
                    "HKLM" {$hiveValue = 2147483650}
                    "HKU" {$hiveValue = 2147483651}
                    "HKCC" {$hiveValue = 2147483653}
                }

                try {
                    
                    $subkeys1 = $wmi.EnumKey($hiveValue,$path)
                    foreach ($subkey1 in $subkeys1.snames) {
                        
                        $productName = $wmi.GetStringValue($hiveValue,"$path\$subkey1","ProductName")
                        $productKey = $wmi.GetBinaryValue($hiveValue,"$path\$subkey1","DigitalProductId")
                        
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