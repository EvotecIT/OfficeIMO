namespace OfficeIMO.Project.Tests;

public sealed class ProjectCustomFieldCalculationTests {
    [Fact]
    public void ProducerFormulasMatchApplicationValuesAndSurviveApplyAndXmlReopen() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("fields"));
        var task = document.Tasks.GetByUid(1);
        var expected = task.CustomFields.Where(f => document.CustomFields.Any(d => d.FieldId == f.FieldId && d.Formula != null))
            .ToDictionary(f => f.FieldId!, f => f.Value);
        var options = new ProjectCustomFieldCalculationOptions { CultureName = "pl-PL" };
        var result = document.CalculateCustomFields(options); result.Report.ThrowIfErrors();
        var calculated = result.Values.Where(v => v.EntityKind == ProjectCustomFieldEntityKind.Task && v.EntityUid == task.Uid).ToDictionary(v => v.FieldId, v => v.Value);
        Assert.Equal(7, calculated.Count);
        Assert.Equal("64", calculated["188743992"]);
        Assert.Equal("0", calculated["188743993"]);
        foreach (var entry in expected) Assert.Equal(entry.Value, calculated[entry.Key]);
        var number2 = task.CustomFields.Single(f => f.FieldId == "188743768"); number2.Value = "4";
        Assert.Throws<InvalidOperationException>(() => document.ApplyCustomFields(result));
        result = document.CalculateCustomFields(options); result.Report.ThrowIfErrors();
        string revisedNumber = (8m + task.Cost!.Value / 100m).ToString(System.Globalization.CultureInfo.InvariantCulture);
        Assert.Equal(revisedNumber, result.Values.Single(v => v.EntityKind == ProjectCustomFieldEntityKind.Task && v.EntityUid == 1 && v.FieldId == "188743767").Value);
        long revision = document.Revision;
        Assert.Throws<OperationCanceledException>(() => document.ApplyCustomFields(result, new CancellationToken(true)));
        Assert.Equal(revision, document.Revision);
        document.ApplyCustomFields(result);
        // The controlled producer fixture contains opaque lookup/UI metadata; adding absent zero values triggers the structural retention warning.
        using var copy = document.Clone(new ProjectSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.Equal("DESIGN:" + revisedNumber.Replace('.', ','), copy.Tasks.GetByUid(1).CustomFields.Single(f => f.FieldId == "188743731").Value);
        var reopened = copy.CalculateCustomFields(options); reopened.Report.ThrowIfErrors();
        Assert.Equal(result.Values.Select(v => (v.EntityKind, v.EntityUid, v.FieldId, v.Value)), reopened.Values.Select(v => (v.EntityKind, v.EntityUid, v.FieldId, v.Value)));
    }
    [Fact]
    public void ProducerSummaryRollupsDistinguishDescendantsFromImmediateChildren() {
        using var document = ProjectDocument.Load(ProjectResourceCapacityTests.Fixture("rollups"));
        var result = document.CalculateCustomFields(); result.Report.ThrowIfErrors();
        Assert.Equal(30, result.Values.Count);
        foreach (var value in result.Values) {
            var task = document.Tasks.GetByUid(value.EntityUid);
            string expected = task.CustomFields.SingleOrDefault(f => f.FieldId == value.FieldId)?.Value ?? "0";
            Assert.Equal(expected, value.Value);
        }
        Assert.Equal("5.5", result.Values.Single(v => v.EntityUid == 0 && v.FieldId == "188743985").Value);
        Assert.Equal("6.75", result.Values.Single(v => v.EntityUid == 0 && v.FieldId == "188743986").Value);
        Assert.Equal("6", result.Values.Single(v => v.EntityUid == 0 && v.FieldId == "188743987").Value);
        Assert.Equal("4", result.Values.Single(v => v.EntityUid == 0 && v.FieldId == "188743989").Value);
    }
    [Theory]
    [InlineData("-2 ^ 2 + 10 / 2", "1")]
    [InlineData("IIf(2 > 1 And Not 3 = 4, Round(2.5, 0), 7)", "2")]
    [InlineData("(2 + 3) * 4", "20")]
    [InlineData("2 ^ 3 ^ 2", "64")]
    [InlineData("IIf(True Xor True Or True, 1, 0)", "0")]
    [InlineData("Int(-2.5) + Fix(-2.5)", "-5")]
    public void NumericExpressionsRespectPrecedenceAndRounding(string formula, string expected) {
        using var document = Example(formula);
        var result = document.CalculateCustomFields(); result.Report.ThrowIfErrors();
        Assert.Equal(expected, Assert.Single(result.Values).Value);
    }
    [Theory]
    [InlineData("Shell(\"anything\")")]
    [InlineData("1 / 0")]
    [InlineData("[Number1] + 1")]
    [InlineData("[Missing] + 1")]
    [InlineData("1;2")]
    [InlineData("IIf(True, 1, 1 / 0)")]
    [InlineData("((((((((1))))))))")]
    public void UnsupportedCyclesErrorsAndBoundsRejectTheWholeProposal(string formula) {
        using var document = Example(formula); long revision = document.Revision;
        var result = document.CalculateCustomFields(new ProjectCustomFieldCalculationOptions { MaxDepth = 5 });
        Assert.True(result.Report.HasErrors);
        Assert.Throws<InvalidDataException>(() => document.ApplyCustomFields(result)); Assert.Equal(revision, document.Revision);
    }
    [Fact]
    public void LookupReferencesResolveOnlyAgainstTheModeledList() {
        using var document = Example("[Number2] * 2");
        var definition = document.CustomFields.Add(); definition.FieldId = "188743768"; definition.RestrictValues = true;
        var lookup = definition.LookupValues.Add(); lookup.Id = 1; lookup.Value = "3"; lookup.Guid = "FE8E76CE-E116-492D-B215-D94A1599B709";
        document.SetCustomFieldLookupValue(document.Tasks[0], definition, lookup);
        var value = document.Tasks[0].CustomFields.Single();
        var result = document.CalculateCustomFields(); result.Report.ThrowIfErrors(); Assert.Equal("6", Assert.Single(result.Values).Value);
        using var copy = document.Clone(); Assert.Equal("6", Assert.Single(copy.CalculateCustomFields().Values).Value);
        value.Value = "4"; Assert.True(document.CalculateCustomFields().Report.HasErrors);
        value.Value = null; value.ValueId = "2"; Assert.True(document.CalculateCustomFields().Report.HasErrors);
    }
    [Theory]
    [InlineData("188743752", 2)]
    [InlineData("188743752", 6)]
    [InlineData("188743752", 7)]
    [InlineData("188743731", 0)]
    [InlineData("188743945", 3)]
    public void UnsupportedRollupAndFieldCombinationsRejectBeforeAggregation(string fieldId, int rollup) {
        using var document = ProjectDocument.Create(); var summary = document.Tasks.AddSummary("Summary");
        var first = summary.Children.Add("First"); var second = summary.Children.Add("Second");
        var definition = document.CustomFields.Add(); definition.FieldId = fieldId; definition.SummaryCalculation = 1; definition.RollupType = rollup;
        if (fieldId == "188743731") {
            var a = first.CustomFields.Add(); a.FieldId = fieldId; a.Value = "ä";
            var b = second.CustomFields.Add(); b.FieldId = fieldId; b.Value = "z";
        }
        long revision = document.Revision;
        var result = document.CalculateCustomFields(new ProjectCustomFieldCalculationOptions { CultureName = "sv-SE" });
        Assert.True(result.Report.HasErrors); Assert.Empty(result.Values);
        Assert.Throws<InvalidDataException>(() => document.ApplyCustomFields(result)); Assert.Equal(revision, document.Revision);
    }
    private static ProjectDocument Example(string formula) {
        var document = ProjectDocument.Create(); document.Tasks.Add("Task");
        var definition = document.CustomFields.Add(); definition.FieldId = "188743767"; definition.Formula = formula;
        return document;
    }
}
