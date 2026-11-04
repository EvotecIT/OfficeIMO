<#
.SYNOPSIS
Creates a project, a native Access database, and an EPUB handbook.
.DESCRIPTION
Install PSWriteOffice from PowerShell Gallery once, then run this example with a
destination directory. Project XML preserves the calculated scheduling fields.
#>
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory
)

Import-Module PSWriteOffice
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null

$project = New-OfficeProject -Name 'Handbook delivery' -StartDate '2026-10-20 08:00'
try {
    $draft = $project | Add-OfficeProjectTask -Name Draft -DurationDays 3 -PassThru
    $review = $project | Add-OfficeProjectTask -Name Review -DurationDays 2 -PassThru
    $project | Add-OfficeProjectDependency -PredecessorUid $draft.Uid -SuccessorUid $review.Uid
    $project | Invoke-OfficeProjectSchedule -Apply | Out-Null
    $project | Save-OfficeProject -Path (Join-Path $OutputDirectory 'delivery.xml')
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $project | Get-OfficeProjectView -Kind Gantt |
            Export-OfficeProjectReport -Path (Join-Path $OutputDirectory 'delivery.pdf')
    }
    $project | Get-OfficeProjectTask | Select-Object Uid, Name, Start, Finish
} finally {
    $project | Close-OfficeProject
}

$database = New-OfficeAccess -Title Contacts
try {
    $contacts = $database | Add-OfficeAccessTable -Name Contacts -PassThru
    $contacts | Add-OfficeAccessColumn -Name Id -Type AutoNumber
    $contacts | Add-OfficeAccessColumn -Name Name -MaxLength 100
    $contacts | Add-OfficeAccessColumn -Name Active -Type Boolean
    @(
        [pscustomobject]@{ Name = 'Operations'; Active = $true }
        [pscustomobject]@{ Name = 'Archive'; Active = $false }
    ) | Add-OfficeAccessRow -Table $contacts
    $database | Save-OfficeAccess -Path (Join-Path $OutputDirectory 'contacts.accdb')
} finally {
    $database | Close-OfficeAccess
}

$handbook = New-OfficeEpub -Title 'Operations handbook' -Language en -Author 'Example team'
$handbook | Add-OfficeEpubChapter -Title Welcome -Text 'Safety & care come first.'
$handbook | Add-OfficeEpubChapter -Title Procedures -Markdown '# Procedure A'
$handbook | Save-OfficeEpub -Path (Join-Path $OutputDirectory 'handbook.epub')
