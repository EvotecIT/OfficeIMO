using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelConditionalFormattingSource.Standard)]
    [InlineData(ExcelConditionalFormattingSource.Office2010Extension)]
    public void ConditionalFormattingLifecycle_UsesNativeWorksheetIdentityAcrossCollectionWrappers(ExcelConditionalFormattingSource source) {
        using var document = ExcelDocument.Create();
        var original = document.AddSheet("Data");
        original.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
            Source = source, Range = "A1:A3", Type = "Expression", Formulas = new[] { "A1>0" }, Priority = 1
        });
        original.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
            Source = source, Range = "C1:C3", Type = "Expression", Formulas = new[] { "C1>0" }, Priority = 2
        });
        document.AddSheet("Other");
        var rules = document.Sheets[0].GetConditionalFormattingRules().OrderBy(rule => rule.Priority).ToArray();
        using var foreign = ExcelDocument.Create();
        var sameNameElsewhere = foreign.AddSheet("Data");
        Assert.Throws<InvalidOperationException>(() => document.Sheets[1].UpdateConditionalFormattingRule(rules[0]));
        Assert.Throws<InvalidOperationException>(() => sameNameElsewhere.UpdateConditionalFormattingRule(rules[0]));
        rules[0].StopIfTrue = true;
        original.UpdateConditionalFormattingRule(rules[0]);
        var clone = original.CloneConditionalFormattingRule(rules[0], "B1:B3");
        original.ReorderConditionalFormattingRules(new[] { clone, rules[0], rules[1] });
        original.RemoveConditionalFormattingRule(rules[1]);
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var stored = reopened.Sheets[0].GetConditionalFormattingRules().OrderBy(rule => rule.Priority).ToArray();
        Assert.Equal(new[] { "B1:B3", "A1:A3" }, stored.Select(rule => rule.Range));
        Assert.All(stored, rule => { Assert.Equal(source, rule.Source); Assert.True(rule.StopIfTrue); });
        Assert.Empty(reopened.Sheets[1].GetConditionalFormattingRules());
        Assert.Empty(sameNameElsewhere.GetConditionalFormattingRules());
    }
}
