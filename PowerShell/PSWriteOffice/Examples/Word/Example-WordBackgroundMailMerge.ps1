Import-Module PSWriteOffice -ErrorAction Stop

$Path = Join-Path $PSScriptRoot 'Example-WordBackgroundMailMerge.docx'

New-OfficeWord -Path $Path {
    Set-OfficeWordBackground -Color '#eef4ff'

    WordParagraph {
        WordText 'Dear '
        WordField -Type MergeField -Parameters '"FirstName"'
        WordText ','
    }

    WordParagraph {
        WordText 'Your order '
        WordField -Type MergeField -Parameters '"OrderId"'
        WordText ' is ready.'
    }

    Invoke-OfficeWordMailMerge -Data @{
        FirstName = 'Ada'
        OrderId   = 4242
    }
}

Get-OfficeWord -Path $Path -ReadOnly | ForEach-Object {
    try {
        $_.Background.Color
        (Get-OfficeWordField -Document $_ -FieldType MergeField).Count
        (Find-OfficeWord -Path $Path -Text 'Ada').Count
        (Find-OfficeWord -Path $Path -Text '4242').Count
    } finally {
        $_ | Close-OfficeWord
    }
}
