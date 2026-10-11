using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Maps a new-starter handover across the manager, people team and IT without automating account changes.</summary>
internal static class OnboardingJourney {
    internal static void Create(string folder) {
        var document = VisioDocument.Create(Path.Combine(folder, "example.vsdx"))
            .SwimlaneDiagram("New starter journey", swim => swim
                .Theme(VisioStyleTheme.Modern())
                .Lane("manager", "Manager").Lane("people", "People team").Lane("it", "IT")
                .Phase("prepare", "Before arrival").Phase("arrival", "Day one").Phase("review", "Week one")
                .Start("plan", "Confirm role and start", "manager", "prepare")
                .Step("details", "Confirm starter details", "people", "prepare")
                .Step("equipment", "Prepare agreed equipment", "it", "prepare")
                .Step("welcome", "Welcome and introduce", "manager", "arrival")
                .Step("access", "Verify required access", "it", "arrival")
                .End("review", "Review first-week needs", "manager", "review")
                .Handoff("plan", "details").Handoff("details", "equipment")
                .Handoff("equipment", "welcome").Handoff("welcome", "access")
                .Handoff("access", "review"));
        document.Save();
    }
}
