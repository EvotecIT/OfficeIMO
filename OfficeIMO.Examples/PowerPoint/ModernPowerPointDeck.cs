using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Examples.PowerPoint {
    /// <summary>
    /// Demonstrates a modern deck composed from content: the built-in Modern theme, the slide layouts that pick
    /// geometry from the content, native charts and tables, a full-bleed picture, speaker notes, transitions and
    /// validation. No coordinates are written by hand.
    /// </summary>
    public static class ModernPowerPointDeck {
        public static void Example_ModernPowerPointDeck(string folderPath, bool openPowerPoint) {
            Console.WriteLine("[*] PowerPoint - Modern themed deck");
            string filePath = Path.Combine(Path.GetFullPath(folderPath), "Modern PowerPoint Deck.pptx");
            string imagePath = Path.Combine(AppContext.BaseDirectory, "Images", "Kulek.jpg");

            OfficeTheme theme = OfficeTheme.Modern;
            using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath, new PowerPointCreateOptions { Theme = theme });

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            var plan = new PowerPointDeckPlan();
            plan.Add(new PowerPointContentPlanSlide("Building decks that explain themselves", Array.Empty<PowerPointLayoutBlock>(),
                PowerPointLayout.Title, "Composition, native charts and tables from one content plan") {
                Eyebrow = "OfficeIMO.PowerPoint",
                Metadata = { "Product team", "Half-year review" },
                Notes = "Cover: the deck is generated from content; the layouts choose the geometry."
            });

            plan.Add(new PowerPointContentPlanSlide("Agenda", new PowerPointLayoutBlock[] {
                new PowerPointLayoutStep("Composition", "Shared margins, one grid and a type scale on every slide."),
                new PowerPointLayoutStep("Content blocks", "Lists, cards, tables and notes that fit their frames."),
                new PowerPointLayoutStep("Native charts", "Readable visuals that stay editable in PowerPoint.")
            }, PowerPointLayout.Process, "What this deck covers") {
                Takeaway = "Charts and tables only help when the layout around them makes the story obvious."
            });

            plan.Add(new PowerPointContentPlanSlide("Service areas", new PowerPointLayoutBlock[] {
                new PowerPointLayoutCard("Device management", "Enrolment and policy from one console.", new[] { "Intune", "Autopilot", "Configuration Manager" }),
                new PowerPointLayoutCard("Identity", "Sign-in and access with explicit guardrails.", new[] { "Entra ID", "Conditional Access", "Reviews" }),
                new PowerPointLayoutCard("Security", "Baselines that are measured, not assumed.", new[] { "Defender", "Baselines", "Compliance" }),
                new PowerPointLayoutCard("Apple", "Mac and iPhone under the same ownership.", new[] { "macOS", "Jamf Pro", "Business Manager" })
            }, PowerPointLayout.Auto, "Four areas with one owner each") {
                Eyebrow = "Capabilities",
                Notes = "Cards keep their headings on one line and their bodies on another, whatever the text length."
            });

            plan.Add(new PowerPointContentPlanSlide("Sales grew every month", new PowerPointLayoutBlock[] {
                new PowerPointLayoutText("Sales nearly tripled from January to June while profit grew at a steadier pace."),
                new PowerPointLayoutChart(OfficeChartKind.ColumnClustered, new OfficeChartData(months, new[] {
                    new OfficeChartSeries("Sales", new[] { 10D, 12D, 16D, 19D, 24D, 27D }),
                    new OfficeChartSeries("Profit", new[] { 4D, 5D, 7D, 9D, 11D, 13D })
                })) { Title = "Sales and profit (USD m)" }
            }) {
                Eyebrow = "Performance",
                Notes = "Keep the conversation on the widening gap between revenue growth and profit."
            });

            plan.Add(new PowerPointContentPlanSlide("Half-year at a glance", new PowerPointLayoutBlock[] {
                new PowerPointLayoutKpi("27", "Sales in June (USD m)", "+170% vs Jan"),
                new PowerPointLayoutKpi("13", "Profit in June (USD m)", "+225% vs Jan"),
                new PowerPointLayoutKpi("Q3", "Runway", "Healthy")
            }) {
                Eyebrow = "Performance",
                Takeaway = "Profit grows faster than sales: the new pricing pays for itself."
            });

            plan.Add(new PowerPointContentPlanSlide("Where new customers came from", new PowerPointLayoutBlock[] {
                new PowerPointLayoutList(new[] { "Partners bring the largest share", "Self-service grew fastest", "Events stay a small channel" }),
                new PowerPointLayoutChart(OfficeChartKind.Doughnut, new OfficeChartData(new[] { "Partners", "Self-service", "Direct", "Events" },
                    new[] { new OfficeChartSeries("Share", new[] { 41D, 29D, 21D, 9D }) })) { Title = "Channel mix (%)" }
            }) { Eyebrow = "Growth" });

            plan.Add(new PowerPointContentPlanSlide("Next-quarter priorities", new PowerPointLayoutBlock[] {
                new PowerPointLayoutTable(new[] { "Priority", "Owner", "Due", "Status" }, new List<IReadOnlyList<string>> {
                    new[] { "Launch the partner portal", "Sales operations", "15 Sep", "On track" },
                    new[] { "Consolidate billing", "Finance", "30 Sep", "At risk" },
                    new[] { "Hire two solution architects", "Engineering", "15 Oct", "On track" },
                    new[] { "Extend support hours", "Customer success", "31 Oct", "Planned" }
                })
            }, PowerPointLayout.Auto, "Owners confirmed at the leadership meeting") { Eyebrow = "Plan" });

            if (File.Exists(imagePath)) {
                plan.Add(new PowerPointContentPlanSlide("Built for real content", new PowerPointLayoutBlock[] {
                    new PowerPointLayoutImage(imagePath, "Photograph used as a full-bleed background")
                }, PowerPointLayout.ImageFull, "Pictures keep their proportions and focal point") { Eyebrow = "Pictures" });
            }

            plan.Add(new PowerPointContentPlanSlide("Decisions we need", new PowerPointLayoutBlock[] {
                new PowerPointLayoutList(new[] { "Approve the partner portal budget", "Confirm the billing owner", "Agree the support hours" })
            }, PowerPointLayout.Closing, "Owners and dates go into the plan this week.") { Eyebrow = "Next steps" });

            // Brand values come from the theme itself, so composition keeps the Modern palette and fonts.
            PowerPointDeckDesign design = PowerPointDeckDesign.FromBrand("#" + theme.Palette.Accent1.ToRgbHex(), "modern-example",
                headingFontName: theme.Fonts.Heading.Family, bodyFontName: theme.Fonts.Body.Family, footerLeft: "OfficeIMO · Modern deck");
            PowerPointCompositionOptions options = PowerPointCompositionOptions.FromDesign(design);
            options.ApplyTheme = false;
            presentation.Compose(plan, options);

            foreach (PowerPointSlide slide in presentation.Slides) slide.Transition = PowerPointSlideTransition.Fade;
            presentation.Save();

            List<OfficeOpenXmlValidationError> errors = presentation.ValidateDocument();
            if (errors.Count > 0) {
                string details = string.Join(Environment.NewLine, errors.Take(5).Select(error => error.Description));
                throw new InvalidOperationException($"PowerPoint validation failed with {errors.Count} error(s).{Environment.NewLine}{details}");
            }

            Console.WriteLine($"    Saved: {filePath}");
            Console.WriteLine("    Validation: no Open XML errors found.");
            if (openPowerPoint) ExampleFileLauncher.Open(filePath);
        }
    }
}
