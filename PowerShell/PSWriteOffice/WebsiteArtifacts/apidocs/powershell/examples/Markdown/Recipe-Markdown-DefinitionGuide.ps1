$path = '.\Operations-Glossary.md'

MarkdownNew -Path $path {
    MarkdownHeading -Level 1 -Text 'Operations glossary'
    MarkdownDefinitionList -Definition @{
        SLO = 'The service level objective agreed with the owner.'
        RTO = 'The target time to restore the service after an incident.'
        RPO = 'The acceptable amount of data loss measured in time.'
    }
    MarkdownDetails -Summary 'How to use these terms' {
        MarkdownParagraph -Text 'Record SLO, RTO, and RPO in the service review before approval.'
    }
}
