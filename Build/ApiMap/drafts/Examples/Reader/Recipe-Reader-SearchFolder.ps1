# Word, Excel, PowerPoint, PDF, Markdown and other files in one search; -Pattern is a regular expression.
Search-OfficeDocument -Path .\Documents -Pattern 'retention (period|policy)' -Recurse

# Which documents mention it most.
Search-OfficeDocument -Path .\Documents -Pattern 'retention' -Recurse | Group-Object Path -NoElement | Sort-Object Count -Descending
