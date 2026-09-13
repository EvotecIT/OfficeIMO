namespace OfficeIMO.Project.Tests;

public sealed class ProjectIndicatorTests {
    [Fact]
    public void RulesUseComputedFieldsInOrderWithoutApplyingThem() {
        using var document = ProjectDocument.Create();
        var task = document.Tasks.Add("Delivery"); task.PercentComplete = 60;
        var number = document.CustomFields.Add(); number.FieldId = "188743767"; number.Formula = "[% Complete] + 10";
        long revision = document.Revision;
        var rules = new[] {
            new ProjectIndicatorRule("[Number1] >= 70", ProjectIndicatorIcon.GreenCircle, "On track"),
            new ProjectIndicatorRule("[% Complete] > 0", ProjectIndicatorIcon.YellowCircle, "In progress")
        };
        var result = document.EvaluateIndicators(ProjectCustomFieldEntityKind.Task, rules); result.Report.ThrowIfErrors();
        Assert.Equal(ProjectIndicatorIcon.GreenCircle, Assert.Single(result.Values).Icon);
        Assert.Equal(0, result.Values[0].RuleIndex); Assert.Equal("On track", result.Values[0].Label);
        Assert.Equal(revision, document.Revision); Assert.Empty(task.CustomFields);
        task.PercentComplete = 0;
        Assert.Null(document.EvaluateIndicators(ProjectCustomFieldEntityKind.Task, rules).Values[0].Icon);
    }
    [Fact]
    public void SummarySelectionErrorsAndBudgetsAreExplicit() {
        using var document = ProjectDocument.Create(); var summary = document.Tasks.AddSummary("Summary"); summary.Children.Add("Leaf");
        var rules = new[] { new ProjectIndicatorRule("1 = 1", ProjectIndicatorIcon.BlueCircle, "Included") };
        var result = document.EvaluateIndicators(ProjectCustomFieldEntityKind.Task, rules, includeSummaries: false); result.Report.ThrowIfErrors();
        Assert.Equal(summary.Children[0].Uid, Assert.Single(result.Values).EntityUid);
        Assert.True(document.EvaluateIndicators(ProjectCustomFieldEntityKind.Task, rules, options: new ProjectCustomFieldCalculationOptions { MaxValues = 1 }).Report.HasErrors);
        var invalid = new[] { new ProjectIndicatorRule("[Unknown]", ProjectIndicatorIcon.RedFlag, "Invalid") };
        Assert.True(document.EvaluateIndicators(ProjectCustomFieldEntityKind.Task, invalid).Report.HasErrors);
        Assert.Throws<OperationCanceledException>(() => document.EvaluateIndicators(ProjectCustomFieldEntityKind.Task, rules, cancellationToken: new CancellationToken(true)));
    }
}
