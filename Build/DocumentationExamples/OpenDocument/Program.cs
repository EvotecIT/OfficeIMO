#region QuickStart
using OfficeIMO.OpenDocument;

OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created with OfficeIMO.OpenDocument.");

OdtTable table = document.AddTable(2, 2, "Results");
table.Cell(0, 0).Text = "Metric";
table.Cell(0, 1).Text = "Value";
table.Cell(1, 0).Text = "Revenue";
table.Cell(1, 1).Text = "42";

document.Save("Summary.odt");
#endregion
