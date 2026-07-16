[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Assert-DiscoveryPath {
    param(
        [Parameter(Mandatory)] [object] $Document,
        [Parameter(Mandatory)] [string] $Path,
        [object] $ExpectedValue
    )

    $current = $Document
    foreach ($segment in $Path.Split(".")) {
        $property = $current.PSObject.Properties[$segment]
        if ($null -eq $property) {
            throw "Required discovery path '$Path' is missing at '$segment'."
        }
        $current = $property.Value
    }

    if ($PSBoundParameters.ContainsKey("ExpectedValue") -and $current -ne $ExpectedValue) {
        throw "Discovery path '$Path' expected '$ExpectedValue' but returned '$current'."
    }
}

$contracts = @(
    @{
        Name = "Google Docs v1"
        Uri = 'https://docs.googleapis.com/$discovery/rest?version=v1'
        Checks = @(
            @{ Path = "resources.documents.methods.create.httpMethod"; Value = "POST" },
            @{ Path = "resources.documents.methods.get.httpMethod"; Value = "GET" },
            @{ Path = "resources.documents.methods.batchUpdate.httpMethod"; Value = "POST" },
            @{ Path = "schemas.Document.properties.revisionId" },
            @{ Path = "schemas.Document.properties.tabs" },
            @{ Path = "schemas.WriteControl.properties.requiredRevisionId" },
            @{ Path = "schemas.WriteControl.properties.targetRevisionId" }
        )
    },
    @{
        Name = "Google Sheets v4"
        Uri = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
        Checks = @(
            @{ Path = "resources.spreadsheets.methods.create.httpMethod"; Value = "POST" },
            @{ Path = "resources.spreadsheets.methods.get.httpMethod"; Value = "GET" },
            @{ Path = "resources.spreadsheets.methods.batchUpdate.httpMethod"; Value = "POST" },
            @{ Path = "resources.spreadsheets.resources.values.methods.batchUpdate.httpMethod"; Value = "POST" },
            @{ Path = "schemas.Spreadsheet.properties.sheets" },
            @{ Path = "schemas.Spreadsheet.properties.developerMetadata" },
            @{ Path = "schemas.BatchUpdateSpreadsheetRequest.properties.requests" }
        )
    },
    @{
        Name = "Google Slides v1"
        Uri = 'https://slides.googleapis.com/$discovery/rest?version=v1'
        Checks = @(
            @{ Path = "resources.presentations.methods.create.httpMethod"; Value = "POST" },
            @{ Path = "resources.presentations.methods.get.httpMethod"; Value = "GET" },
            @{ Path = "resources.presentations.methods.batchUpdate.httpMethod"; Value = "POST" },
            @{ Path = "schemas.Presentation.properties.revisionId" },
            @{ Path = "schemas.Presentation.properties.slides" },
            @{ Path = "schemas.NotesProperties.properties.speakerNotesObjectId" },
            @{ Path = "schemas.WriteControl.properties.requiredRevisionId" }
        )
    },
    @{
        Name = "Google Drive v3"
        Uri = "https://www.googleapis.com/discovery/v1/apis/drive/v3/rest"
        Checks = @(
            @{ Path = "resources.files.methods.get.httpMethod"; Value = "GET" },
            @{ Path = "resources.files.methods.create.httpMethod"; Value = "POST" },
            @{ Path = "resources.files.methods.export.httpMethod"; Value = "GET" },
            @{ Path = "resources.changes.methods.getStartPageToken.httpMethod"; Value = "GET" },
            @{ Path = "resources.changes.methods.list.httpMethod"; Value = "GET" },
            @{ Path = "schemas.ChangeList.properties.nextPageToken" },
            @{ Path = "schemas.ChangeList.properties.newStartPageToken" },
            @{ Path = "schemas.ChangeList.properties.changes" }
        )
    }
)

foreach ($contract in $contracts) {
    Write-Host "Checking $($contract.Name) discovery contract..."
    $document = Invoke-RestMethod -Uri $contract.Uri -Method Get
    foreach ($check in $contract.Checks) {
        if ($check.ContainsKey("Value")) {
            Assert-DiscoveryPath -Document $document -Path $check.Path -ExpectedValue $check.Value
        } else {
            Assert-DiscoveryPath -Document $document -Path $check.Path
        }
    }
}

Write-Host "Google Workspace discovery contracts are present."
