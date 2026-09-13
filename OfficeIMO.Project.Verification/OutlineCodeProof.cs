using OfficeIMO.Project;

internal static class OutlineCodeProof {
    internal static int Create(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        Directory.CreateDirectory(output);
        using var edited = ProjectDocument.Load(input);
        edited.OutlineCodes.Single().Values[1].Value = "02";
        edited.Save(Path.Combine(output, "edited.xml"));
        using var authored = ProjectDocument.Create();
        authored.Calendar = authored.Calendars.AddStandardWorkingWeek();
        authored.Settings.StartDate = new DateTime(2026, 10, 5, 8, 0, 0);
        var task = authored.Tasks.Add("Concept"); task.Duration = ProjectDuration.WorkingDays(1);
        var table = authored.OutlineCodes.Add(); table.Guid = Guid.NewGuid().ToString();
        table.AllLevelsRequired = true; table.OnlyTableValuesAllowed = true;
        var mask = table.Masks.Add(); mask.Level = 1; mask.Type = 3; mask.Length = 0; mask.Separator = ".";
        mask = table.Masks.Add(); mask.Level = 2; mask.Type = 0; mask.Length = 2; mask.Separator = "-";
        var root = table.Values.Add(); root.ValueId = 1; root.ParentValueId = 0; root.Type = 21; root.Value = "Design"; root.Guid = Guid.NewGuid().ToString();
        var leaf = table.Values.Add(); leaf.ValueId = 2; leaf.ParentValueId = 1; leaf.Type = 21; leaf.Value = "03"; leaf.Guid = Guid.NewGuid().ToString();
        var field = authored.CustomFields.Add(); field.FieldId = "188744096"; field.FieldName = "Outline Code1"; field.Alias = "Discipline"; field.LookupTableGuid = table.Guid;
        authored.SetOutlineCodeValue(task, field.FieldId, leaf);
        authored.Save(Path.Combine(output, "authored.xml"));
        Console.WriteLine("Edited Design.02; authored Design.03.");
        return 0;
    }
}
