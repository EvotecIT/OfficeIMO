Search-OfficeDocument `
    -Path '.\Documents' `
    -Pattern 'Retention' `
    -Recurse `
    -MaximumResults 0 |
    Select-Object DocumentType, Path, Match
