#region QuickStart
using OfficeIMO;
using OfficeIMO.Word;

Milestone[] milestones = [new("Discovery", "Product", "Done"), new("Pilot", "Operations", "At risk"), new("Rollout", "IT", "Planned")];

using var document = WordDocument.Create("Status.docx");
document.AddHeading("Northwind migration", level: 1);

WordParagraph status = document.AddParagraph("Overall status: ");
status.AddText("On track").Bold = true;

document.AddHeading("Milestones", level: 2);
document.AddTable(milestones);

document.Save();

record Milestone(string Name, string Owner, string Status);
#endregion
