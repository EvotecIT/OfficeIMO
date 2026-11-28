#region QuickStart
using OfficeIMO;
using OfficeIMO.Visio;

using var document = VisioDocument.Create("Flow.vsdx");
document.Properties.Title = "Approval flow";
var page = document.AddPage("Page-1");
page.AddTitle("Approval flow");
page.AddRectangle("start", OfficeLength.Inches(1), OfficeLength.Inches(1),
    OfficeLength.Inches(2), OfficeLength.Inches(1), "Start").SetBold();
page.AddDiamond("decision", OfficeLength.Inches(4), OfficeLength.Inches(1.5),
    OfficeLength.Inches(2), OfficeLength.Inches(2), "Approved?");
page.AddEllipse("end", OfficeLength.Inches(7), OfficeLength.Inches(1.5),
    OfficeLength.Inches(2), OfficeLength.Inches(1), "Done");
page.AddConnector("start", "decision", fromSide: VisioSide.Right, toSide: VisioSide.Left)
    .SetKind(ConnectorKind.RightAngle).SetArrowEnd(EndArrow.Triangle);
page.AddConnector("decision", "end", fromSide: VisioSide.Right, toSide: VisioSide.Left)
    .SetKind(ConnectorKind.RightAngle).SetArrowEnd(EndArrow.Triangle).SetLabel("Yes");
document.Save();
#endregion
