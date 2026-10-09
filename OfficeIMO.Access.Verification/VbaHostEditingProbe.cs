using System.Text;

namespace OfficeIMO.Access.Verification {
    internal static class VbaHostEditingProbe {
        internal static void Generate(string fixtures, string outputRoot) {
            Directory.CreateDirectory(outputRoot);
            foreach (string name in new[] { "designer-jet4.mdb", "designer-ace12.accdb" }) {
                string input = Path.Combine(fixtures, "Designer", name); byte[] original = File.ReadAllBytes(input);
                using AccessDocument document = AccessDocument.Load(input);
                var unchanged = document.ApplicationStreams.Where(x => !x.Path.StartsWith("VBA/") && x.Path != "Forms/0/PropData" && x.Path != "Reports/0/PropData")
                    .ToDictionary(x => x.Path, x => x.Payload.GetBytes());
                document.Forms["BoundForm1"].SetCodeBehind("Option Explicit\r\nPrivate Sub Form_Open(Cancel As Integer)\r\n 'Synthetic marker 61\r\nEnd Sub\r\n");
                document.Reports["BoundReport"].SetCodeBehind("Option Explicit\r\nPrivate Sub Report_Open(Cancel As Integer)\r\n 'Synthetic marker 62\r\nEnd Sub\r\n");
                if (name.EndsWith(".accdb", StringComparison.OrdinalIgnoreCase)) {
                    document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Open, "[Event Procedure]");
                    document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.AfterUpdate, "=Len(\"proof\")", "GroupChoice1");
                    document.Forms["BoundForm1"].SetEventBinding(AccessEventKind.Click, "=Len(\"click\")", "Title1");
                    document.Forms["BoundForm2"].SetEventBinding(AccessEventKind.Click, "=Len(\"new\")", "Title2");
                    document.Reports["BoundReport"].SetEventBinding(AccessEventKind.Open, "[Event Procedure]");
                    unchanged.Remove("Forms/0/Blob"); unchanged.Remove("Forms/1/Blob"); unchanged.Remove("Reports/0/Blob");
                }
                string output = Path.Combine(outputRoot, name); document.Save(output);
                using AccessDocument saved = AccessDocument.Load(output);
                foreach (var stream in unchanged) if (!stream.Value.SequenceEqual(saved.ApplicationStreams.Single(x => x.Path == stream.Key).Payload.GetBytes()))
                    throw new InvalidDataException("An unrelated application stream changed: " + stream.Key);
                if (!original.SequenceEqual(File.ReadAllBytes(input))) throw new InvalidDataException("The input fixture changed.");
                Expected(output + ".form.expected.txt", saved.GetVbaProject().GetModule("Form_BoundForm1").Source);
                Expected(output + ".report.expected.txt", saved.GetVbaProject().GetModule("Report_BoundReport").Source);
            }
            using AccessDocument empty = AccessDocument.Load(Path.Combine(fixtures, "CodeBehind", "empty-designer-ace12.accdb"));
            AccessApplicationObject host = empty.Forms["EventForm"];
            host.SetEventBinding(AccessEventKind.Click, "=Len(\"no-vba\")", "EventLabel");
            empty.Save(Path.Combine(outputRoot, "event-only.accdb"));
            host.SetEventBinding(AccessEventKind.Click, null, "EventLabel");
            empty.Save(Path.Combine(outputRoot, "cleared.accdb"));
            host.SetCodeBehind("Option Explicit\r\nPrivate Sub EventLabel_Click()\r\n 'Synthetic marker 71\r\nEnd Sub\r\n");
            host.SetEventBinding(AccessEventKind.Click, "[Event Procedure]", "EventLabel");
            string first = Path.Combine(outputRoot, "first-code-behind.accdb"); empty.Save(first);
            Expected(first + ".form.expected.txt", empty.GetVbaProject().GetModule("Form_EventForm").Source);
        }
        private static void Expected(string path, string source) => File.WriteAllText(path,
            string.Join("\n", source.Replace("\r\n", "\n").Split('\n').Where(x => !x.StartsWith("Attribute ", StringComparison.OrdinalIgnoreCase))).TrimEnd('\n'), new UTF8Encoding(false));
    }
}
