function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Identity,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Trustee,

        [Parameter(Mandatory = $true)]
        [ValidateSet('SendAs')]
        [System.String[]]
        $AccessRights,

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
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [Switch]
        $ManagedIdentity
    )

    Write-Verbose -Message "Getting configuration of Office 365 Recipient permission $Identity"
    if ($Script:ExportMode)
    {
        $ConnectionMode = New-M365DSCConnection -Workload 'ExchangeOnline' `
            -InboundParameters $PSBoundParameters `
            -SkipModuleReload $true
    }
    else
    {
        $ConnectionMode = New-M365DSCConnection -Workload 'ExchangeOnline' `
            -InboundParameters $PSBoundParameters
    }

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName -replace 'MSFT_', ''
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $nullReturn = $PSBoundParameters
    $nullReturn.Ensure = 'Absent'

    try
    {

        if ($null -ne $Script:recipientPermissions -and $Script:ExportMode)
        {
            $recipientPermission = $Script:recipientPermissions | Where-Object -FilterScript {
                $_.Identity -eq $Identity -and $_.Trustee -eq $Trustee -and $_.AccessRights -eq $AccessRights
            }
        }
        else
        {
            #Could include a switch for the different propertySets to retrieve https://learn.microsoft.com/en-us/powershell/exchange/cmdlet-property-sets?view=exchange-ps#get-exomailbox-property-sets
            #Could include a switch for the different recipientTypeDetails to retrieve
            $recipientPermission = Get-EXORecipientPermission -Identity $Identity -Trustee $Trustee -AccessRights $AccessRights -ErrorAction Stop
        }

        if ($null -eq $recipientPermission)
        {
            Write-Verbose -Message "The specified Recipient Permission doesn't already exist."
            return $nullReturn
        }

        #endregion

        $result = @{
            Identity              = $Identity
            Trustee               = $recipientPermission.Trustee
            AccessRights          = $recipientPermission.AccessRights

            Ensure                = 'Present'
            Credential            = $Credential
            ApplicationId         = $ApplicationId
            CertificateThumbprint = $CertificateThumbprint
            CertificatePath       = $CertificatePath
            CertificatePassword   = $CertificatePassword
            Managedidentity       = $ManagedIdentity.IsPresent
            TenantId              = $TenantId
        }

        Write-Verbose -Message "Found an existing instance of Recipient permissions '$($DisplayName)'"
        return $result
    }
    catch
    {
        New-M365DSCLogEntry -Message 'Error retrieving data:' `
            -Exception $_ `
            -Source $($MyInvocation.MyCommand.Source) `
            -TenantId $TenantId `
            -Credential $Credential

        return $nullReturn
    }
}

function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [System.String]
        $Identity,

        [Parameter()]
        [System.String]
        $Trustee,

        [Parameter()]
        [ValidateSet('SendAs')]
        [System.String]
        $AccessRights,

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
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [Switch]
        $ManagedIdentity
    )

    Write-Verbose -Message "Setting Mail Contact configuration for $Name"

    $currentState = Get-TargetResource @PSBoundParameters

    if ($Global:CurrentModeIsExport)
    {
        $ConnectionMode = New-M365DSCConnection -Workload 'ExchangeOnline' `
            -InboundParameters $PSBoundParameters `
            -SkipModuleReload $true
    }
    else
    {
        $ConnectionMode = New-M365DSCConnection -Workload 'ExchangeOnline' `
            -InboundParameters $PSBoundParameters
    }

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName -replace 'MSFT_', ''
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    $parameters = $PSBoundParameters
    $parameters.Remove('Credential') | Out-Null
    $parameters.Remove('ApplicationId') | Out-Null
    $parameters.Remove('TenantId') | Out-Null
    $parameters.Remove('CertificateThumbprint') | Out-Null
    $parameters.Remove('CertificatePath') | Out-Null
    $parameters.Remove('CertificatePassword') | Out-Null
    $parameters.Remove('ManagedIdentity') | Out-Null
    $parameters.Remove('Ensure') | Out-Null

    # Receipient Permission doesn't exist but it should
    if ($Ensure -eq 'Present' -and $currentState.Ensure -eq 'Absent')
    {
        Write-Verbose -Message "The Receipient Permission for '$Trustee' with Access Rights '$($AccessRights -join ', ')' on mailbox '$Identity' does not exist but it should. Adding it."
        Add-RecipientPermission @parameters -Confirm:$false
    }
    # Receipient Permission exists but shouldn't
    elseif ($Ensure -eq 'Absent' -and $currentState.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Receipient Permission for '$Trustee' with Access Rights '$($AccessRights -join ', ')' on mailbox '$Identity' exists but shouldn't. Removing it."
        Remove-RecipientPermission @parameters -Confirm:$false
    }
    elseif ($Ensure -eq 'Present' -and $currentState.Ensure -eq 'Present')
    {
        Write-Verbose -Message "Receipient Permission for '$Trustee' with Access Rights '$($AccessRights -join ', ')' on mailbox '$Identity' exists."
    }
}

function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [System.String]
        $Identity,

        [Parameter()]
        [System.String]
        $Trustee,

        [Parameter()]
        [ValidateSet('SendAs')]
        [System.String]
        $AccessRights,

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
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [Switch]
        $ManagedIdentity
    )

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName -replace 'MSFT_', ''
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    Write-Verbose -Message "Testing configuration of Office 365 Recipient permissions $DisplayName"

    $currentValues = Get-TargetResource @PSBoundParameters

    Write-Verbose -Message "Current Values: $(Convert-M365DscHashtableToString -Hashtable $currentValues)"
    Write-Verbose -Message "Target Values: $(Convert-M365DscHashtableToString -Hashtable $PSBoundParameters)"

    $testResult = Test-M365DSCParameterState -CurrentValues $currentValues `
        -Source $($MyInvocation.MyCommand.Source) `
        -DesiredValues $PSBoundParameters `
        -ValuesToCheck @('Ensure', 'Identity')

    Write-Verbose -Message "Test-TargetResource returned $testResult"

    return $testResult
}

function Export-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.String])]
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [System.String]
        $Identity,

        [Parameter()]
        [System.String]
        $Trustee,

        [Parameter()]
        [ValidateSet('SendAs')]
        [System.String]
        $AccessRights,

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
        [System.String]
        $CertificatePath,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $CertificatePassword,

        [Parameter()]
        [Switch]
        $ManagedIdentity
    )

    $ConnectionMode = New-M365DSCConnection -Workload 'ExchangeOnline' `
        -InboundParameters $PSBoundParameters `
        -SkipModuleReload $true

    #Ensure the proper dependencies are installed in the current environment.
    Confirm-M365DSCDependencies

    #region Telemetry
    $ResourceName = $MyInvocation.MyCommand.ModuleName -replace 'MSFT_', ''
    $CommandName = $MyInvocation.MyCommand
    $data = Format-M365DSCTelemetryParameters -ResourceName $ResourceName `
        -CommandName $CommandName `
        -Parameters $PSBoundParameters
    Add-M365DSCTelemetryEvent -Data $data
    #endregion

    try
    {
        $Script:ExportMode = $true

        [array]$Script:recipientPermissions = Get-EXORecipientPermission -ResultSize Unlimited

        $dscContent = ''
        $i = 1
        if ($recipientPermissions.Length -eq 0)
        {
            Write-Host $Global:M365DSCEmojiGreenCheckMark
        }
        else
        {
            Write-Host "`r`n" -NoNewline
        }
        foreach ($recipientPermission in $recipientPermissions)
        {
            Write-Host "    |---[$i/$($recipientPermissions.Length)] $($recipientPermission.Identity)" -NoNewline

            $params = @{
                Identity              = $recipientPermission.Identity
                Trustee               = $recipientPermission.Trustee
                AccessRights          = $recipientPermission.AccessRights

                Credential            = $Credential
                ApplicationId         = $ApplicationId
                TenantId              = $TenantId
                CertificateThumbprint = $CertificateThumbprint
                CertificatePassword   = $CertificatePassword
                Managedidentity       = $ManagedIdentity.IsPresent
                CertificatePath       = $CertificatePath
            }

            $Results = Get-TargetResource @Params

            if ($Results -is [System.Collections.Hashtable] -and $Results.Count -gt 1)
            {
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

                Write-Host $Global:M365DSCEmojiGreenCheckMark
            }
            else
            {
                Write-Host $Global:M365DSCEmojiRedX
            }

            $i++

        }
        return $dscContent
    }
    catch
    {
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
