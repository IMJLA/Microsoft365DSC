function Remove-Quote {
    param (
        [string]$String
    )
    $first = $String[0]
    $max = $String.Length - 1
    $last = $String[$max]
    if ($first -eq '"' -and $last -eq '"' -or
        $first -eq "'" -and $last -eq "'") {
        return $String.Substring(1, $String.Length - 2)
    }
    else {
        return $String
    }            
}

function ConvertTo-PowerShellHashtableCode {
    param ([hashtable]$Hashtable)
    
    $KeyStrings = ForEach ($Key in $Hashtable.Keys) {
        "'$Key' = '$($Hashtable[$Key])'"
    }
    $KeysString = $KeyStrings -join "`r`n                                        "
    $FullString = "@{$KeysString`r`n                                    }"
    return $FullString
}

function ConvertFrom-PowerShelHashtableCode {
    param ([string]$String)
    $regex = '[^@{]([\s]*[^=]*[\s]*=[\s]*[^;\r\n}]*)[\s;]*'
    $KeyPairs = [regex]::Matches($str, $regex).Value
    $out = @{}
    ForEach ($pair in $KeyPairs) {
        $split = $pair.Split('=')
        $key = Remove-Quote -String $split[0].Trim()
        $value = Remove-Quote -String $split[1].Trim()
        $out[$key] = $value
    }
    return $out
}

function Get-AADHomeRealmDiscoveryPolicyInstance {

    [CmdletBinding()]

    param (
        [string]$DisplayName,
        [string]$Id
    )

    if ($PSBoundParameters.ContainsKey('Id')) {
        $instance = Get-MgBetaPolicyHomeRealmDiscoveryPolicy -HomeRealmDiscoveryPolicyId $Id
    }
    else {
        $instance = Get-MgBetaPolicyHomeRealmDiscoveryPolicy -Filter "DisplayName eq '$DisplayName'" -ErrorAction Stop
    }

    if ($instance.Count -gt 1) {
        Write-Warning -Message "Found multiple instances of a HomeRealmDiscoveryPolicy named {$DisplayName}, which could result in inconsistencies retrieving its values. The instances will be sorted alphabetically by Id and the first one will be returned."
        $instance = $instance | Sort-Object -Property Id | Select-Object -First 1
    }

    return $instance

}

function Get-TargetResource {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        
        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [String]
        $BodyParameter,

        [Parameter()]
        [String]
        $AppliesTo,

        [Parameter()]
        [String]
        $Definition,

        [Parameter()]
        [String]
        $DeletedDateTime,

        [Parameter()]
        [String]
        $Description,

        [Parameter()]
        [String]
        $Id,

        [Parameter()]
        [Boolean]
        $IsOrganizationDefault,

        [Parameter()]
        [System.String]
        $AdditionalProperties,


        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    New-M365DSCConnection -Workload MicrosoftGraph -InboundParameters $PSBoundParameters | Out-Null

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $nullResult = $PSBoundParameters
    $nullResult.Ensure = 'Absent'
    try {
        if ($null -ne $Script:exportedInstances -and $Script:ExportMode) {
            
            $instance = $Script:exportedInstances | Where-Object -FilterScript { $_.DisplayName -eq $DisplayName }
        }
        else {
            
            $instance = Get-AADHomeRealmDiscoveryPolicyInstance -DisplayName $DisplayName -Id $Id -ErrorAction Stop

        }
        if ($null -eq $instance) {
            return $nullResult
        }

        if ($null -eq $instance.AdditionalProperties) {
            Write-Host "instance has null AdditionalProperties"
        }
        else {
            Write-Host "`r`ninstance has some additionalProperties: $($instance.AdditionalProperties.Keys)"
        }
        $AdditionalPropertiesAsString = ConvertTo-PowerShellHashtableCode -Hashtable $instance.AdditionalProperties
        Write-Host $AdditionalPropertiesAsString -ForegroundColor Cyan

        $results = @{
            DisplayName           = $instance.DisplayName
            BodyParameter         = $instance.BodyParameter
            AdditionalProperties  = $AdditionalPropertiesAsString
            AppliesTo             = $instance.AppliesTo
            Definition            = [string]$instance.Definition
            DeletedDateTime       = $instance.DeletedDateTime
            Description           = $instance.Description
            Id                    = $instance.Id
            IsOrganizationDefault = $instance.IsOrganizationDefault
            Ensure                = 'Present'
            Credential            = $Credential
            ApplicationId         = $ApplicationId
            TenantId              = $TenantId
            CertificateThumbprint = $CertificateThumbprint
            ManagedIdentity       = $ManagedIdentity.IsPresent
            AccessTokens          = $AccessTokens
        }
        return [System.Collections.Hashtable] $results
    }
    catch {
        New-M365DSCLogEntry -Message 'Error retrieving data:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        return $nullResult
    }
}

function Set-TargetResource {
    [CmdletBinding()]
    param
    (
        
        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [String]
        $BodyParameter,

        [Parameter()]
        [String]
        $AppliesTo,

        [Parameter()]
        [String]
        $Definition,

        [Parameter()]
        [String]
        $DeletedDateTime,

        [Parameter()]
        [String]
        $Description,

        [Parameter()]
        [String]
        $Id,

        [Parameter()]
        [Boolean]
        $IsOrganizationDefault,

        [Parameter()]
        [System.String]
        $AdditionalProperties,


        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $currentInstance = Get-TargetResource @PSBoundParameters

    $setParameters = Remove-M365DSCAuthenticationParameter -BoundParameters $PSBoundParameters
    $SetParameters['AdditionalProperties'] = ConvertFrom-PowerShelHashtableCode -String $currentInstance.AdditionalProperties

    # CREATE
    if ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Absent') {
        New-MgBetaPolicyHomeRealmDiscoveryPolicy @SetParameters
    }
    # UPDATE
    elseif ($Ensure -eq 'Present' -and $currentInstance.Ensure -eq 'Present') {
            
        $currentInstance

        if ($null -ne $currentInstance) {
            $setParameters.Remove('Id')
            $setParameters.Add('HomeRealmDiscoveryPolicyId', $currentInstance.Id)
            Update-MgBetaPolicyHomeRealmDiscoveryPolicy @SetParameters
        }
        else {
            Write-Warning "Could not find AADHomeRealmDiscoveryPolicy with Displayname '$DisplayName' to update it."
        }

    }
    # REMOVE
    elseif ($Ensure -eq 'Absent' -and $currentInstance.Ensure -eq 'Present') {

        if ($null -ne $currentInstance) {
            $setParameters.Remove('Id')
            $setParameters.Add('HomeRealmDiscoveryPolicyId', $currentInstance.Id)
            Remove-MgBetaPolicyHomeRealmDiscoveryPolicy -HomeRealmDiscoveryPolicyId $currentInstance
        }
        else {
            Write-Warning "Could not find AADHomeRealmDiscoveryPolicy with Displayname '$DisplayName' to remove it."
        }

    }
}

function Test-TargetResource {
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        
        [Parameter(Mandatory = $true)]
        [System.String]
        $DisplayName,

        [Parameter()]
        [String]
        $BodyParameter,

        [Parameter()]
        [String]
        $AppliesTo,

        [Parameter()]
        [String]
        $Definition,

        [Parameter()]
        [String]
        $DeletedDateTime,

        [Parameter()]
        [String]
        $Description,

        [Parameter()]
        [String]
        $Id,

        [Parameter()]
        [Boolean]
        $IsOrganizationDefault,

        [Parameter()]
        [System.String]
        $AdditionalProperties,


        [Parameter()]
        [ValidateSet('Present', 'Absent')]
        [System.String]
        $Ensure = 'Present',

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $CurrentValues = Get-TargetResource @PSBoundParameters
    $ValuesToCheck = ([Hashtable]$PSBoundParameters).Clone()

    Write-Verbose -Message "Current Values: $(Convert-M365DscHashtableToString -Hashtable $CurrentValues)"
    Write-Verbose -Message "Target Values: $(Convert-M365DscHashtableToString -Hashtable $ValuesToCheck)"

    $testResult = Test-M365DSCParameterState -CurrentValues $CurrentValues `
        -Source $($MyInvocation.MyCommand.Source) `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck $ValuesToCheck.Keys

    Write-Verbose -Message "Test-TargetResource returned $testResult"

    return $testResult
}

function Export-TargetResource {
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $ManagedIdentity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    
    $ConnectionMode = New-M365DSCConnection -Workload MicrosoftGraph `
        -InboundParameters $PSBoundParameters

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName.Replace('MSFT_', '')
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    try {
        $Script:ExportMode = $true
        
        [array] $Script:exportedInstances = Get-MgBetaPolicyHomeRealmDiscoveryPolicy -ErrorAction Stop

        $i = 1
        $dscContent = ''
        if ($Script:exportedInstances.Length -eq 0) {
            Write-Host $Global:M365DSCEmojiGreenCheckMark
        }
        else {
            Write-Host "`r`n" -NoNewline
        }
        foreach ($config in $Script:exportedInstances) {
            if ($null -ne $Global:M365DSCExportResourceInstancesCount) {
                $Global:M365DSCExportResourceInstancesCount++
            }

            $displayedKey = $config.Id
            Write-Host "    |---[$i/$($Script:exportedInstances.Count)] $displayedKey" -NoNewline

            $params = @{
                
                DisplayName           = $config.DisplayName
                Credential            = $Credential
                ApplicationId         = $ApplicationId
                TenantId              = $TenantId
                CertificateThumbprint = $CertificateThumbprint
                ManagedIdentity       = $ManagedIdentity.IsPresent
                AccessTokens          = $AccessTokens
            }

            $Results = Get-TargetResource @Params
            Write-Host "    |---[$i/$($Script:exportedInstances.Count)] $($Results.AdditionalProperties)" -NoNewline
            $Results = Update-M365DSCExportAuthenticationResults -ConnectionMode $ConnectionMode `
                -Results $Results

            $currentDSCBlock = Get-M365DSCExportContentForResource -ResourceName $ResourceName `
                -ConnectionMode $ConnectionMode `
                -ModulePath $PSScriptRoot `
                -Results $Results `
                -Credential $Credential
            $dscContent += $currentDSCBlock
            Save-M365DSCPartialExport -Content $currentDSCBlock `
                -FileName $Global:PartialExportFileName
            $i++
            Write-Host $Global:M365DSCEmojiGreenCheckMark
        }
        return $dscContent
    }
    catch {
        Write-Host $Global:M365DSCEmojiRedX

        New-M365DSCLogEntry -Message 'Error during Export:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        return ''
    }
}

Export-ModuleMember -Function *-TargetResource

