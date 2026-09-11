using System.Text.Json;
using OfficeIMO;
using OfficeIMO.Project;

internal static class CustomFieldProof {
    internal static int Create(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory."); Directory.CreateDirectory(output);
        using var document = ProjectDocument.Load(input);
        if (document.CustomFields.Any(d => d.FieldId == "188743767" && d.Formula != null))
            document.Tasks.GetByUid(1).CustomFields.Single(f => f.FieldId == "188743768").Value = "4";
        else document.AllTasks.First(t => !t.IsSummary && t.Uid != 0).CustomFields.Single(f => f.FieldId == "188743984").Value = "3";
        var result = document.CalculateCustomFields(new ProjectCustomFieldCalculationOptions { CultureName = "pl-PL" });
        document.ApplyCustomFields(result);
        document.Save(Path.Combine(output, Path.GetFileName(input)), new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Console.WriteLine(JsonSerializer.Serialize(result.Values, new JsonSerializerOptions { WriteIndented = true }));
        return 0;
    }
    internal static int Verify(string expectedFile, string actualFile) {
        using var expected = ProjectDocument.Load(expectedFile); using var actual = ProjectDocument.Load(actualFile);
        var result = expected.CalculateCustomFields(new ProjectCustomFieldCalculationOptions { CultureName = "pl-PL" }); result.Report.ThrowIfErrors();
        var differences = new List<object>();
        foreach (var value in result.Values) {
            // Project changes its project-summary label during Save As; ordinary tasks, nested summaries and resources retain identity.
            if (value.EntityKind == ProjectCustomFieldEntityKind.Task && value.EntityUid == 0) continue;
            var fields = value.EntityKind == ProjectCustomFieldEntityKind.Task ? actual.Tasks.GetByUid(value.EntityUid).CustomFields : actual.Resources.GetByUid(value.EntityUid).CustomFields;
            string? observed = fields.SingleOrDefault(f => f.FieldId == value.FieldId)?.Value;
            if (observed == null && (value.Value == "0" || value.Value == "")) continue;
            if (observed != value.Value) differences.Add(new { value.EntityKind, value.EntityUid, value.FieldId, expected = value.Value, actual = observed });
        }
        Console.WriteLine(JsonSerializer.Serialize(new { name = Path.GetFileName(expectedFile), differences }, new JsonSerializerOptions { WriteIndented = true }));
        return differences.Count == 0 ? 0 : 1;
    }
}
