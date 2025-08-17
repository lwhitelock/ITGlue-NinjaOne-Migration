function Get-NinjaOneToken {
    [CmdletBinding()]
    param()

    if ($Script:NinjaOneInstance -and $Script:NinjaOneClientID -and $Script:NinjaOneClientSecret ) {
        if ($Script:NinjaTokenExpiry -and (Get-Date) -lt $Script:NinjaTokenExpiry) {
            return $Script:NinjaToken
        } else {

            if ($Script:NinjaOneRefreshToken) {
                $Body = @{
                    'grant_type'    = 'refresh_token'
                    'client_id'     = $Script:NinjaOneClientID
                    'client_secret' = $Script:NinjaOneClientSecret
                    'refresh_token' = $Script:NinjaOneRefreshToken
                }
            } else {

                $body = @{
                    grant_type    = 'client_credentials'
                    client_id     = $Script:NinjaOneClientID
                    client_secret = $Script:NinjaOneClientSecret
                    scope         = 'monitoring management'
                }
            }

            $token = Invoke-RestMethod -Uri "https://$($Script:NinjaOneInstance -replace '/ws','')/ws/oauth/token" -Method Post -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
    
            $Script:NinjaTokenExpiry = (Get-Date).AddSeconds(3000)
            $Script:NinjaToken = $token
            

            Write-Host 'Fetched New Token'
            return $token
        } else {
            Throw 'Please run Connect-NinjaOne first'
        }
    }

}

function Connect-NinjaOne {
    [CmdletBinding()]
    param (
        [Parameter(mandatory = $true)]
        $NinjaOneInstance,
        [Parameter(mandatory = $true)]
        $NinjaOneClientID,
        [Parameter(mandatory = $true)]
        $NinjaOneClientSecret,
        $NinjaOneRefreshToken
    )

    $Script:NinjaOneInstance = $NinjaOneInstance
    $Script:NinjaOneClientID = $NinjaOneClientID
    $Script:NinjaOneClientSecret = $NinjaOneClientSecret
    $Script:NinjaOneRefreshToken = $NinjaOneRefreshToken
    

    try {
        $Null = Get-NinjaOneToken -ea Stop
    } catch {
        Throw "Failed to Connect to NinjaOne: $_"
    }

}

function Invoke-NinjaOneRequest {
    param(
        $Method,
        $Body,
        $InputObject,
        $Path,
        $QueryParams,
        [Switch]$Paginate,
        [Switch]$AsArray
    )

    $Token = Get-NinjaOneToken

    if ($InputObject) {
        if ($AsArray) {
            $Body = $InputObject | ConvertTo-Json -depth 100
            if (($InputObject | Measure-Object).count -eq 1 ) {
                $Body = '[' + $Body + ']'
            }
        } else {
            $Body = $InputObject | ConvertTo-Json -depth 100
        }
    }

    try {
        if ($Method -in @('GET', 'DELETE')) {
            if ($Paginate) {           
                $After = 0
                $PageSize = 1000
                $NinjaResult = do {
                    $Result = (Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)?pageSize=$PageSize&after=$After$(if ($QueryParams){"&$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json' -UseBasicParsing).content | ConvertFrom-Json
                    $Result
                    $ResultCount = ($Result.id | Measure-Object -Maximum)
                    $After = $ResultCount.maximum
                } while (($Result | Measure-Object).count -eq $PageSize)

                Return $NinjaResult

            } else {
                $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -ContentType 'application/json; charset=utf-8' -UseBasicParsing
            }       

        } elseif ($Method -in @('PATCH', 'PUT', 'POST')) {
            $NinjaResult = Invoke-WebRequest -uri "https://$($Script:NinjaOneInstance)/api/v2/$($Path)$(if ($QueryParams){"?$QueryParams"})" -Method $Method -Headers @{Authorization = "Bearer $($token.access_token)" } -Body $Body -ContentType 'application/json; charset=utf-8' -UseBasicParsing
        } else {
            Throw 'Unknown Method'
        }
    } catch {
        Throw "Error Occured: $_"
    }

    try {
        return $NinjaResult.content | ConvertFrom-Json -ea stop
    } catch {
        return $NinjaResult.content
    }

}

function Invoke-NinjaOneDocumentTemplate {
    [CmdletBinding()]
    param (
        $Template,
        $ID
    )

    if (!$ID) {
        $DocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates" -Method GET -QueryParams "templateName=$($Template.name)" | where-object { $_.name -eq $Template.name }
    } else {
        $DocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates/$($ID)" -Method GET
    }
    
    $PatchTemplate = $False

    $MatchedCount = ($DocumentTemplate | Measure-Object).count
    if ($MatchedCount -eq 1) {
        # Matched a single document template
        # Check fields are correct
        foreach ($Field in $Template.Fields) {
            if ($Field.fieldName) {
                $MatchedField = $DocumentTemplate.Fields | Where-Object { $_.fieldName -eq $Field.fieldName -and $_.fieldType -eq $Field.fieldType }
                if (($MatchedField | Measure-Object).count -ne 1) {
                    $MatchedField = $DocumentTemplate.Fields | Where-Object { $_.fieldName -eq $Field.fieldName }
                    $MatchCount = ($MatchedField | Measure-Object).count
                    if ($MatchCount -eq 1 ) {
                        Throw "$($Field.fieldName) exists with the wrong type. Please manually edit the template $($Template.name) to set it to a $($Field.fieldType) field."
                    } elseif ($MatchCount -eq 0) {
                        $PatchTemplate = $True
                    } else {
                        Throw "Mutliple Fields exists for $($Field.fieldName) in $($Template.name)"
                    }
                }
            } elseif ($field.uiElementType) {
                continue
            } else {
                Throw "$Field had no fieldName or uiElementName"
            }
        }

        if ($PatchTemplate -eq $True) {
            Write-Host "Updating Template"
            $Template.Fields = $Template.Fields + ($DocumentTemplate.Fields | Where-Object { $_.fieldName -notin $Template.Fields.fieldName })
            $NinjaDocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates/$($DocumentTemplate.id)" -Method PUT -InputObject ($Template | Select-Object * -ExcludeProperty allowMultiple)
        }

        $NinjaDocumentTemplate = $DocumentTemplate


    } elseif ($MatchedCount -eq 0) {
        # Create a new Document Template
        Write-Host "Creating Template"
        $NinjaDocumentTemplate = Invoke-NinjaOneRequest -Path "document-templates" -Method POST -InputObject $Template
    } else {
        # Matched multiple templates. Should be impossible but lets check anyway :D
        Throw "Multiiple Documents Matched the Provided Criteria"
    }

    return $NinjaDocumentTemplate

}

function Get-NinjaOneTime {
    [CmdletBinding()]
    param (
        [Parameter(Position = 0, Mandatory = $false)]
        [DateTime]$Date,
        [Switch]$Seconds
    )
    try {
        $unixEpoch = Get-Date -Year 1970 -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
        $timeSpan = $Date.ToUniversalTime() - $unixEpoch

        if ($Seconds) {
            return [int64]([math]::Round($timeSpan.TotalSeconds))
        } else {
            return [int64]([math]::Round($timeSpan.TotalMilliSeconds))
        }
    } catch {
        return ""
    }
}

