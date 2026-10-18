Search-OfficeDocument `
    -Path '.\Documents' `
    -Query 'Retention' `
    -Recurse `
    -AllResults |
    Select-Object DocumentType, Path, Match
