using System.IO;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a service briefing deck with a cover, responsibilities, and a handover checklist.</summary>
internal static class ServiceOverview {
    internal static void Create(string folder) {
        using PowerPointPresentation deck = PowerPointPresentation.Create(Path.Combine(folder, "example.pptx"));
        PowerPointSlide cover = deck.AddSlide();
        cover.AddRectangleCm(0, 0, 33.866, 19.05, "Cover background").Fill("17365D").Stroke("17365D", 0);
        Text(cover, "NORTHWIND / OPERATIONS", 1.7, 2, 29, 1, 17, "B8D4FF");
        Text(cover, "A service the next team\ncan own", 1.7, 5, 29, 5.5, 40, "FFFFFF");
        Text(cover, "Ownership, support, and the first-week handover.", 1.7, 12.5, 29, 2, 23, "D8E5F7");

        PowerPointSlide owners = deck.AddSlide();
        Text(owners, "Make responsibility visible", 1.5, 1.5, 30, 2, 32, "17365D");
        string[] titles = { "Service owner", "Operations", "Engineering" };
        string[] bodies = { "Sets priorities\nAccepts the service\nReviews outcomes", "Runs the service\nCoordinates incidents\nMaintains the runbook", "Resolves defects\nShips improvements\nSupports recovery" };
        for (int index = 0; index < titles.Length; index++) {
            double x = 1.5 + index * 10.3;
            owners.AddRectangleCm(x, 5, 9.3, 10, titles[index]).Fill("EAF1FB").Stroke("CBD5E1", 1);
            Text(owners, titles[index], x + 0.5, 5.8, 8.3, 1.5, 25, "17365D");
            Text(owners, bodies[index], x + 0.5, 8, 8.3, 5.5, 21, "526179");
        }

        PowerPointSlide handover = deck.AddSlide();
        Text(handover, "The first-week handover", 1.5, 1.5, 30, 2, 32, "17365D");
        string[] checks = { "Walk through the support queue together.", "Exercise one recovery procedure.", "Confirm alert ownership and escalation.", "Agree the first service review date." };
        for (int index = 0; index < checks.Length; index++) {
            Text(handover, $"0{index + 1}", 1.7, 5 + index * 2.5, 2, 1.7, 25, "2563EB");
            Text(handover, checks[index], 4, 5 + index * 2.5, 27, 1.7, 23, "17365D");
        }
        deck.Save();

        void Text(PowerPointSlide slide, string value, double x, double y, double width, double height, int size, string color) {
            var box = slide.AddTextBoxCm(value, x, y, width, height);
            box.FontSize = size;
            box.Color = color;
        }
    }
}
