using System.Text;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class MarketReadinessProofGallery {
        private static void CreateModelTemplateBindingProof(string scenarioPath) {
            string templatePath = Path.Combine(scenarioPath, "service-summary-template.docx");
            using (WordDocument template = WordDocument.Create(templatePath)) {
                template.AddParagraph("Service summary for {{Client.Name}}").Style = WordParagraphStyles.Heading1;
                template.AddParagraph("{{#each Services}}");
                WordTable service = template.AddTable(1, 2, WordTableStyle.TableGrid);
                service.Rows[0].Cells[0].Paragraphs[0].Text = "{{Name}}";
                service.Rows[0].Cells[1].Paragraphs[0].Text = "{{Hours}} hours for {{Client.Name}}";
                template.AddParagraph("{{#Priority}}");
                template.AddParagraph("Priority delivery");
                template.AddParagraph("{{/Priority}}");
                template.AddParagraph("{{/each Services}}");
                template.AddParagraph("Portal: {{Portal}}");
                template.AddParagraph("Logo: {{Logo}}");
                template.Save();
            }

            string outputPath = Path.Combine(scenarioPath, "service-summary-generated.docx");
            File.Copy(templatePath, outputPath, overwrite: true);
            using WordDocument generated = WordDocument.Load(outputPath);
            var values = new Dictionary<string, object?> {
                ["Client"] = new Dictionary<string, object?> { ["Name"] = "Northwind Traders" },
                ["Services"] = new object[] {
                    new Dictionary<string, object?> { ["Name"] = "Assessment", ["Hours"] = 8, ["Priority"] = true },
                    new Dictionary<string, object?> { ["Name"] = "Implementation", ["Hours"] = 24, ["Priority"] = false }
                },
                ["Portal"] = new WordTemplateHyperlink("Open customer portal", new Uri("https://example.com/customer")),
                ["Logo"] = new WordTemplateImage(
                    File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png")),
                    "EvotecLogo.png",
                    width: 96,
                    height: 32,
                    description: "Evotec logo")
            };

            WordTemplateResult result = WordTemplate.Apply(generated, values).EnsureComplete();
            generated.Save();

            var evidence = new StringBuilder();
            evidence.AppendLine("# Plain template binding evidence");
            evidence.AppendLine();
            evidence.AppendLine($"- Placeholders discovered: {result.PlaceholderCount}");
            evidence.AppendLine($"- Placeholders replaced: {result.ReplacedPlaceholderCount}");
            evidence.AppendLine($"- Repeated block instances: {result.RepeatedBlockCount}");
            evidence.AppendLine($"- Conditional blocks evaluated: {result.ConditionalBlockCount}");
            evidence.AppendLine($"- Missing values: {(result.MissingValueNames.Count == 0 ? "none" : string.Join(", ", result.MissingValueNames))}");
            evidence.AppendLine("- Artifact validation: recorded in the gallery proof manifest after generation");
            File.WriteAllText(Path.Combine(scenarioPath, "template-binding-result.md"), evidence.ToString(), Encoding.UTF8);
        }
    }
}
