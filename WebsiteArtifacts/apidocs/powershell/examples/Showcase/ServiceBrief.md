# Review workplace support demand

[Showcase-PowerPoint-ServiceBrief.ps1](Showcase-PowerPoint-ServiceBrief.ps1) creates an editable eight-slide brief for a service owner reviewing quarterly demand and a proposed support pilot. It combines a semantic deck plan with a directly composed chart and action table, then reopens the presentation to inspect its slides.

The script uses PSWriteOffice 3.0.8 or newer and runs on PowerShell 7 and Windows PowerShell 5.1. Run it from a checkout containing both the script and its [sample CSV](../Samples/ServiceBrief/Quarterly-Demand.csv):

```powershell
./Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1 `
    -OutputDirectory ./ServiceBrief -ReviewDate '2026-09-30'

# Use a reviewed quarterly export and choose the date for the proposed actions.
./Examples/Showcase/Showcase-PowerPoint-ServiceBrief.ps1 `
    -InputCsv ./Quarterly-Demand.csv -ReviewDate '2026-12-31' `
    -OutputDirectory ./DecemberReview
```

The CSV needs these columns:

| Column | Meaning |
| --- | --- |
| `Site` | Unique location name. |
| `Owner` | Team accountable for reviewing that location's demand. |
| `SupportedPeople` | Positive whole number of people supported at the location. |
| `PreviousQuarterTickets` | Nonnegative whole number of tickets in the comparison quarter. |
| `CurrentQuarterTickets` | Nonnegative whole number of tickets in the review quarter. |

This fixed eight-slide layout accepts one to six locations. Use a short, unique site label of at most 20 characters for the cards and chart; owner labels may contain up to 40 characters. Both must be single lines without a pipe character. Numeric fields use plain whole numbers without grouping separators. Invalid input stops before the output folder is created.

The synthetic input covers 600 people and four locations. Tickets fall from 1,200 to 1,080, a 10% decrease overall, while remote staff demand increases by 16 tickets. The decision slide therefore calls for a review of remote staff intake. Changing the CSV recalculates the totals, comparison, location cards, chart and priority review; a zero comparison total is reported as having no prior-quarter baseline.

The deck covers the transition stages, service scope, supported locations, acceptance criteria, decision, ticket comparison and three dated actions. The chart compares ticket counts in both series. Demand alone does not establish service quality or explain why it changed. Owners and actions are proposals to review before distributing the brief.

Open `Showcase-PowerPoint-ServiceBrief.pptx` to edit the proposal. To create a review copy:

```powershell
Export-OfficeDocumentPdf -InputPath ./ServiceBrief/Showcase-PowerPoint-ServiceBrief.pptx `
    -Path ./ServiceBrief/ServiceBrief.pdf
```

Inspect the saved slides and PDF with the fonts and converter used for delivery. Keep the CSV, editable presentation and approved review copy together; rerun from the CSV when the figures change. `-Open` opens the saved presentation, and `-ModulePath` selects an explicit module manifest for local qualification.
