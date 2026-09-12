namespace OfficeIMO.Project.Tests;

public sealed class ProjectInlineLookupValidationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void MissingLookupListsRejectReferencesButAllowUnrestrictedText(bool resource, bool defined) {
        using var document = ProjectDocument.Create();
        string fieldId = resource ? "205521008" : "188743767";
        if (defined) document.CustomFields.Add().FieldId = fieldId;
        var selection = resource ? document.Resources.AddWork("Resource").CustomFields.Add() : document.Tasks.Add("Task").CustomFields.Add();
        selection.FieldId = fieldId; selection.ValueId = "1";
        Assert.Contains(document.Validate().Diagnostics, d => d.Code == "PROJECT_LOOKUP_VALUE_REFERENCE");
        selection.ValueId = null; selection.Value = "10";
        using var copy = document.Clone(); copy.Validate().ThrowIfErrors();
        var reopened = resource ? copy.Resources[0].CustomFields.Single() : copy.Tasks[0].CustomFields.Single();
        Assert.Equal("10", reopened.Value); Assert.Null(reopened.ValueId);
    }

    [Theory]
    [InlineData(false, "text")]
    [InlineData(true, "text")]
    [InlineData(false, "id")]
    [InlineData(true, "id")]
    [InlineData(false, "guid")]
    [InlineData(true, "guid")]
    [InlineData(false, "disagreement")]
    [InlineData(true, "disagreement")]
    public void InvalidInlineSelectionsAreRejectedBeforeXmlOutput(bool resource, string invalid) {
        using var document = ProjectDocument.Create();
        var definition = document.CustomFields.Add(); definition.FieldId = resource ? "205521008" : "188743767";
        definition.RestrictValues = invalid == "text";
        var entry = definition.LookupValues.Add(); entry.Id = 1; entry.Value = "10"; entry.Guid = "FE8E76CE-E116-492D-B215-D94A1599B709";
        var selection = resource ? document.Resources.AddWork("Resource").CustomFields.Add() : document.Tasks.Add("Task").CustomFields.Add();
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
        var reopened = resource ? copy.Resources[0].CustomFields.Single() : copy.Tasks[0].CustomFields.Single();
        Assert.Equal("10", reopened.Value); Assert.Equal("1", reopened.ValueId); Assert.Equal(entry.Guid, reopened.ValueGuid);
    }
}
