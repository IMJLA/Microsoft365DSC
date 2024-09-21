<#
This example is used to test new resources and showcase the usage of new resources being worked on.
It is not meant to use as a production baseline.
#>

Configuration Example
{
    param(
        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint
    )
    Import-DscResource -ModuleName Microsoft365DSC
    node localhost
    {
        AADHomeRealmDiscoveryPolicy "AADHomeRealmDiscoveryPolicy-test2" {
            AdditionalProperties  = "@odata.context=https://graph.microsoft.com/beta/`$metadata#policies/homeRealmDiscoveryPolicies/`$entity";
            Definition            = "{`"HomeRealmDiscoveryPolicy`":{`"AccelerateToFederatedDomain`":true,`"PreferredDomain`":`"federated.example.edu`",`"AlternateIdLogin`":{`"Enabled`":true}}}";
            DisplayName           = "test2-renamed"; # DRIFT
            Ensure                = "Present";
            IsOrganizationDefault = $False;
            ApplicationId         = $ApplicationId;
            TenantId              = $TenantId;
            CertificateThumbprint = $CertificateThumbprint;
        }
    }
}
