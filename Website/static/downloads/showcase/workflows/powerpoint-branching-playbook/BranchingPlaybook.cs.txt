using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a slide-based support playbook with internal links and return navigation.</summary>
internal static class BranchingPlaybook {
    internal static void Create(string folder) {
        using var deck = PowerPointPresentation.Create(Path.Combine(folder, "example.pptx"));
        var menu = deck.AddSlide(); var access = deck.AddSlide(); var service = deck.AddSlide();
        Title(menu, "Choose the support path");
        Text(menu, "Select a card in presentation mode to open the relevant checklist.", 1.5, 3.5, 30, 2, 22);
        var accessLink = menu.AddTextBoxCm("One person cannot sign in\nOpen the access checklist", 2, 7, 14, 5);
        accessLink.FontSize = 25; accessLink.Color = "17365D"; accessLink.FillColor = "EAF1FB";
        accessLink.SetHyperlink(access, "Open the access checklist");
        var serviceLink = menu.AddTextBoxCm("Several people report failure\nOpen the service checklist", 18, 7, 14, 5);
        serviceLink.FontSize = 25; serviceLink.Color = "17365D"; serviceLink.FillColor = "E7F6ED";
        serviceLink.SetHyperlink(service, "Open the service checklist");
        Title(access, "Access checklist");
        Text(access, "1. Confirm the affected account and exact error.\n2. Check the user's access and recent changes.\n3. Record the result and route to the access owner.", 2, 5, 29, 8, 26);
        Title(service, "Service checklist");
        Text(service, "1. Establish scope and check service health.\n2. Name an incident owner and capture evidence.\n3. Use the approved recovery procedure if needed.", 2, 5, 29, 8, 26);
        foreach (var slide in new[] { access, service }) {
            var back = slide.AddTextBoxCm("Return to support paths", 2, 15, 16, 2);
            back.FontSize = 20; back.Color = "2563EB"; back.SetHyperlink(menu, "Return to support paths");
        }
        menu.Notes.Text = "The cards are internal slide links. PDF and PNG review copies do not demonstrate presentation-mode behavior.";
        deck.Save();

        void Title(PowerPointSlide slide, string value) => Text(slide, value, 1.5, 1.3, 30, 2, 32);
        void Text(PowerPointSlide slide, string value, double x, double y, double width, double height, int size) {
            var box = slide.AddTextBoxCm(value, x, y, width, height); box.FontSize = size; box.Color = "17365D";
        }
    }
}
