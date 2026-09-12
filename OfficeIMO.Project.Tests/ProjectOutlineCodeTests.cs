namespace OfficeIMO.Project.Tests;

public sealed class ProjectOutlineCodeTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void BrokenSharedScalarLookupIdentitiesBlockSave(bool changeTable) {
        using var project = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("fields"));
        var table = project.OutlineCodes[0];
        if (changeTable) table.Guid = System.Guid.NewGuid().ToString();
        else table.Values.Single(v => v.Value == "Ready").Guid = System.Guid.NewGuid().ToString();
        Assert.True(project.Validate().HasErrors);
        using var output = new MemoryStream();
        Assert.Throws<InvalidDataException>(() => project.Save(output));
        Assert.Empty(output.ToArray());
    }
    [Fact]
    public void ProducerSharedScalarLookupFeedsFormulasAndExplicitSelection() {
        using var project = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("fields"));
        var definition = project.CustomFields.Single(f => f.FieldId == "188743734");
        var table = project.GetCustomFieldLookupTable(definition)!;
        Assert.NotNull(table); Assert.Empty(definition.LookupValues);
        var task = project.Tasks.GetByUid(1);
        project.SetCustomFieldLookupValue(task, definition, table.Values.Single(v => v.Value == "Waiting"));
        var formula = project.CustomFields.Add(); formula.FieldId = "188743737"; formula.Formula = "[Text2]";
        var result = project.CalculateCustomFields(new ProjectCustomFieldCalculationOptions { CultureName = "pl-PL" });
        result.Report.ThrowIfErrors();
        Assert.Equal("Waiting", result.Values.Single(v => v.EntityUid == task.Uid && v.FieldId == formula.FieldId).Value);
        using var copy = project.Clone(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        var selected = copy.Tasks.GetByUid(1).CustomFields.Single(v => v.FieldId == definition.FieldId);
        Assert.Equal("Waiting", selected.Value); Assert.NotNull(selected.ValueGuid);
    }
    [Fact]
    public void ProducerHierarchyUsesLookupGuidAndKeepsTypedEditsAcrossXmlReopen() {
        using var project = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("outline"));
        var table = Assert.Single(project.OutlineCodes);
        var task = project.Tasks.GetByUid(1);
        Assert.Equal("Design.01", project.GetOutlineCodeText(task, "188744096"));
        Assert.Equal(2, table.Masks.Count); Assert.Equal(2, table.Values.Count);
        using var unchanged = new MemoryStream(); project.Save(unchanged);
        Assert.Equal(File.ReadAllBytes(ProjectResourceCapacityTests.Fixture("outline")), unchanged.ToArray());
        table.Values[1].Value = "02";
        using var copy = project.Clone();
        Assert.Equal("Design.02", copy.GetOutlineCodeText(copy.Tasks.GetByUid(1), "188744096"));
        long revision = project.Revision;
        Assert.Throws<InvalidDataException>(() => project.SetOutlineCodeValue(task, "188744096", table.Values[0]));
        Assert.Equal(revision, project.Revision);
        project.SetOutlineCodeValue(task, "188744096", table.Values[1]);
        Assert.Equal("Design.02", project.GetOutlineCodeText(task, "188744096"));
        table.Values[1].Value = "wrong";
        Assert.Contains(project.Validate().Diagnostics, d => d.Code == "PROJECT_OUTLINE_SELECTION");
        using var rejected = new MemoryStream(); Assert.Throws<InvalidDataException>(() => project.Save(rejected)); Assert.Empty(rejected.ToArray());
    }
    [Fact]
    public void SetOutlineCodeValueRepairsLegacyScalarSelectionPayload() {
        using var project = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("outline"));
        var table = Assert.Single(project.OutlineCodes);
        var task = project.Tasks.GetByUid(1);
        var selected = Assert.Single(task.OutlineCodes);
        selected.Value = "legacy";
        selected.DurationFormat = 7;
        Assert.Contains(project.Validate().Diagnostics, d => d.Code == "PROJECT_OUTLINE_SELECTION");

        project.SetOutlineCodeValue(task, "188744096", table.Values[1]);

        Assert.Null(selected.Value);
        Assert.Null(selected.DurationFormat);
        Assert.DoesNotContain(project.Validate().Diagnostics, d => d.Code == "PROJECT_OUTLINE_SELECTION");
        Assert.Equal("Design.01", project.GetOutlineCodeText(task, "188744096"));
    }
    [Fact]
    public void InvalidOutlineReferencesAndCyclesAreRejectedBeforeSerialization() {
        using var project = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("outline"));
        var table = project.OutlineCodes[0]; var task = project.Tasks.GetByUid(1);
        task.OutlineCodes[0].ValueGuid = System.Guid.NewGuid().ToString();
        Assert.Contains(project.Validate().Diagnostics, d => d.Code == "PROJECT_OUTLINE_SELECTION");
        task.OutlineCodes[0].ValueGuid = table.Values[1].Guid;
        table.Values[0].ParentValueId = table.Values[1].ValueId;
        Assert.Contains(project.Validate().Diagnostics, d => d.Code == "PROJECT_OUTLINE_HIERARCHY");
        table.Values[0].ParentValueId = 0;
        Assert.Throws<ArgumentException>(() => project.SetOutlineCodeValue(task, "205521174", table.Values[1]));
        table.Values.Remove(table.Values[1]);
        Assert.Contains(project.Validate().Diagnostics, d => d.Code == "PROJECT_OUTLINE_SELECTION");
    }
}
