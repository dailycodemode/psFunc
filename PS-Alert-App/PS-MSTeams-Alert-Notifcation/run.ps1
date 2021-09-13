using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host "My app logs here."

$teamsWebhook = 'https://version1.webhook.office.com/webhookb2/16e57eab-3acf-46d6-ba76-ae0f05a110f0@3e0088dc-0629-4ae6-aa8c-813e7a296f50/IncomingWebhook/8cae23d604ca491b850a3037fb952d59/54683452-e2d6-4a3c-98b0-7b6ec58ddcbc'
$body = ConvertTo-Json -Depth 4 @{
    summary  = $Request.body.data.essentials.alertRule
    sections = @(
        @{
            activityTitle    = "Alert has fired: " + $Request.body.data.essentials.alertRule
            # activityText     = "Alert has fired: " + $Request.body.data.essentials.alertRule
            activitySubtitle = $Request.body.data.essentials.description
        },
        @{
            title = 'Details of Alert ' + $Request.body.data.essentials.alertRule
            facts = @(
                @{
                    name  = 'Resource Type'
                    value = $Request.body.data.alertContext.condition.allOf[0].metricNamespace
                },
                @{
                    name  = 'Name'
                    value = $Request.body.data.essentials.configurationItems[0]
                },
                @{
                    name  = 'Alert Type'
                    value = $Request.body.data.alertContext.condition.allOf[0].metricName
                },
                @{
                    name  = 'Operator'
                    value = $Request.body.data.alertContext.condition.allOf[0].operator
                },
                @{
                    name  = 'Threshold'
                    value = $Request.body.data.alertContext.condition.allOf[0].threshold
                },
                @{
                    name  = 'Severity'
                    value = $Request.body.data.essentials.severity
                },
                @{
                    name  = ' '
                    value = ''
                },
                @{
                    name  = 'Detected'
                    value = $Request.body.data.essentials.firedDateTime
                },                
                @{
                    name  = 'Link To Search Results'
                    value = "[Link]($($Request.body.data.essentials.alertTargetIDs[0]))"
                }
            )
        }
    )
}

Invoke-RestMethod -uri $teamsWebhook -Method Post -body $body -ContentType 'application/json'

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body       = $body
    })