using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Builds an organisation chart with team bands, an assistant, a vacancy and an external adviser.</summary>
internal static class TeamOrganisation {
    internal static void Create(string folder) {
        var theme = VisioStyleTheme.Modern();
        theme.Container.TextStyle!.VerticalAlignment = VisioTextVerticalAlignment.Top;
        theme.Container.TextStyle.TopMargin = 0.08;
        var document = VisioDocument.Create(Path.Combine(folder, "example.vsdx"))
            .OrgChartDiagram("Service team", org => org
                .Theme(theme).Title("Service team").TeamBandPadding(0.45).Spacing(0.7, 1.1)
                .Root("director", "Maya Ellis", "Service Director")
                .Assistant("coordinator", "Leon Brooks", "Team Coordinator", "director")
                .Manager("delivery", "Sofia Chen", "Delivery Lead", "director")
                .Manager("operations", "Amir Patel", "Operations Lead", "director")
                .TeamBand("delivery-team", "Delivery", "delivery")
                .TeamBand("operations-team", "Operations", "operations")
                .Position("analyst", "Eva Morgan", "Service Analyst", "delivery", "delivery-team")
                .Position("engineer", "Noah Reed", "Platform Engineer", "operations", "operations-team")
                .Vacancy("support", "Support Specialist", "operations", "operations-team")
                .External("adviser", "Jordan Lee", "Service Adviser", "delivery"));
        document.Save();
    }
}
