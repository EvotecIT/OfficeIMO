using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Maps a purchase approval across requester, budget owner and operations with an explicit revision path.</summary>
internal static class ApprovalSwimlane {
    internal static void Create(string folder) {
        var document = VisioDocument.Create(Path.Combine(folder, "example.vsdx"))
            .SwimlaneDiagram("Purchase approval", swim => swim
                .Theme(VisioStyleTheme.Modern())
                .Lane("requester", "Requester").Lane("owner", "Budget owner").Lane("operations", "Operations")
                .Phase("intake", "Intake").Phase("review", "Review").Phase("decision", "Decision").Phase("fulfil", "Fulfil")
                .Start("submit", "Submit request", "requester", "intake")
                .Step("check", "Check need", "owner", "review")
                .Decision("approve", "Approved?", "owner", "decision")
                .Step("revise", "Clarify request", "requester", "decision")
                .Step("order", "Place order", "operations", "decision")
                .End("receive", "Confirm receipt", "operations", "fulfil")
                .Flow("submit", "check", "request").Flow("check", "approve")
                .Exception("approve", "revise", "revise").Handoff("approve", "order", "yes")
                .Flow("revise", "check", "resubmit").Flow("order", "receive"));
        document.Save();
    }
}
