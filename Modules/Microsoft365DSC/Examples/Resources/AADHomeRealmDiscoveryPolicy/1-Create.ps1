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
        AADHomeRealmDiscoveryPolicy "AADHomeRealmDiscoveryPolicy-test" {
            AdditionalProperties  = "@odata.context=https://graph.microsoft.com/beta/`$metadata#policies/homeRealmDiscoveryPolicies/`$entity";
            Definition            = "{`"HomeRealmDiscoveryPolicy`":{`"AccelerateToFederatedDomain`":true,`"PreferredDomain`":`"federated.example.edu`",`"AlternateIdLogin`":{`"Enabled`":true}}}";

            <# The Description parameter is not supported for Creation.  Resulting error is pasted below:

            New-MgPolicyHomeRealmDiscoveryPolicy_CreateExpanded: Resource '' does not exist or one of its queried reference-property objects are not present.

            Status: 404 (NotFound)
            ErrorCode: Request_ResourceNotFound
            Date: 2024-10-04T19:03:45

            Headers:
            Cache-Control                 : no-cache
            Vary                          : Accept-Encoding
            Strict-Transport-Security     : max-age=31536000
            request-id                    : db352a4e-e950-4661-901d-2dc6ef4d845c
            client-request-id             : 0e40782b-72f9-4495-abc7-95945e1f42a1
            x-ms-ags-diagnostic           : {"ServerInfo":{"DataCenter":"West US 2","Slice":"E","Ring":"4","ScaleUnit":"000","RoleInstance":"CO1PEPF00000D13"}}
            x-ms-resource-unit            : 1
            Date                          : Fri, 04 Oct 2024 19:03:45 GMT
            
            #>
            # Description           = "Example";

            DisplayName           = "test";
            Ensure                = "Present";
            IsOrganizationDefault = $False;
            ApplicationId         = $ApplicationId;
            TenantId              = $TenantId;
            CertificateThumbprint = $CertificateThumbprint;
        }        
    }
}

Example -ApplicationId $env:M365DSCApplicationId -CertificateThumbprint $env:M365DSCCertificateThumbprint -TenantId $env:M365DSCTenantId
