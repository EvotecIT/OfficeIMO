namespace OfficeIMO.Project.Tests;

public sealed class ProjectInlineLookupValidationTests {
    [Theory]
    [InlineData("task", false)]
    [InlineData("task", true)]
    [InlineData("resource", false)]
    [InlineData("resource", true)]
    [InlineData("assignment", false)]
    [InlineData("assignment", true)]
    public void MissingLookupListsRejectReferencesButAllowUnrestrictedText(string owner, bool defined) {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); var resource = document.Resources.AddWork("Resource");
        string fieldId = owner == "resource" ? "205521008" : "188743767";
        if (defined) document.CustomFields.Add().FieldId = fieldId;
        var selection = owner == "task" ? task.CustomFields.Add() : owner == "resource" ? resource.CustomFields.Add() : document.Assignments.Add(task, resource).CustomFields.Add();
        selection.FieldId = fieldId; selection.ValueId = "1";
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_LOOKUP_VALUE_REFERENCE");
        selection.ValueId = null; selection.Value = "10";
        using var copy = document.Clone(); copy.Validate().ThrowIfErrors();
        var reopened = owner == "task" ? copy.Tasks[0].CustomFields.Single() : owner == "resource" ? copy.Resources[0].CustomFields.Single() : copy.Assignments[0].CustomFields.Single();
        Assert.Equal("10", reopened.Value); Assert.Null(reopened.ValueId);
    }

    [Theory]
    [InlineData("task", "text")]
    [InlineData("resource", "text")]
    [InlineData("assignment", "text")]
    [InlineData("task", "id")]
    [InlineData("resource", "id")]
    [InlineData("assignment", "id")]
    [InlineData("task", "guid")]
    [InlineData("resource", "guid")]
    [InlineData("assignment", "guid")]
    [InlineData("task", "disagreement")]
    [InlineData("resource", "disagreement")]
    [InlineData("assignment", "disagreement")]
    public void InvalidInlineSelectionsAreRejectedBeforeXmlOutput(string owner, string invalid) {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Task"); var resource = document.Resources.AddWork("Resource");
        var definition = document.CustomFields.Add(); definition.FieldId = owner == "resource" ? "205521008" : "188743767";
        definition.RestrictValues = invalid == "text";
        var entry = definition.LookupValues.Add(); entry.Id = 1; entry.Value = "10"; entry.Guid = "FE8E76CE-E116-492D-B215-D94A1599B709";
        var selection = owner == "task" ? task.CustomFields.Add() : owner == "resource" ? resource.CustomFields.Add() : document.Assignments.Add(task, resource).CustomFields.Add();
        selection.FieldId = definition.FieldId;
        if (invalid == "text") selection.Value = "20";
        if (invalid == "id") selection.ValueId = "2";
        if (invalid == "guid") selection.ValueGuid = "FE8E76CE-E116-492D-B215-D94A1599B710";
        if (invalid == "disagreement") { selection.ValueId = "1"; selection.Value = "20"; }
        long revision = document.Revision;
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_LOOKUP_VALUE_REFERENCE");
        Assert.Contains(document.AssessSave(new ProjectSaveOptions()).Diagnostics, d => d.Code == "PROJECT_LOOKUP_VALUE_REFERENCE");
        using var bytes = new MemoryStream(); Assert.Throws<InvalidDataException>(() => document.Save(bytes)); Assert.Equal(0, bytes.Length);
        Assert.Equal(revision, document.Revision);
        selection.Value = "10"; selection.ValueId = "1"; selection.ValueGuid = entry.Guid;
        document.Validate().ThrowIfErrors();
        using var copy = document.Clone(); copy.Validate().ThrowIfErrors();
        var reopened = owner == "task" ? copy.Tasks[0].CustomFields.Single() : owner == "resource" ? copy.Resources[0].CustomFields.Single() : copy.Assignments[0].CustomFields.Single();
        Assert.Equal("10", reopened.Value); Assert.Equal("1", reopened.ValueId); Assert.Equal(entry.Guid, reopened.ValueGuid);
    }
}
