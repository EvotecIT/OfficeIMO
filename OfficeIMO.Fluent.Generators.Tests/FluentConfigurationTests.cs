using OfficeIMO.Fluent.Generators.Tests.Sample;
using Xunit;

namespace OfficeIMO.Fluent.Generators.Tests;

public sealed class FluentConfigurationTests {
    [Fact]
    public void ChainReturnsTheSameReceiverAndKeepsTheDerivedType() {
        var heading = new HeadingRun();

        HeadingRun result = heading.Bold().FontSize(14).FontFamily("Aptos");
        result.Level = 2;

        Assert.Same(heading, result);
        Assert.True(heading.Bold);
        Assert.Equal(14, heading.FontSize);
        Assert.Equal("Aptos", heading.FontFamily);
    }

    [Fact]
    public void BooleanSwitchMeansTrueAndKeepsFalseAndInheritanceDistinct() {
        var run = new SampleRun();

        Assert.True(run.Bold().Bold);
        Assert.False(run.Bold(false).Bold);
        Assert.Null(run.Bold(null).Bold);
        Assert.True(run.Hidden().IsHidden);
    }

    [Fact]
    public void FluentCallGoesThroughTheCanonicalMutatorAndItsValidation() {
        var run = new SampleRun().FontSize(12);

        Assert.Throws<ArgumentOutOfRangeException>(() => run.FontSize(0));

        Assert.Equal(12, run.FontSize);
        Assert.Equal(1, run.FontSizeWrites);
    }

    [Fact]
    public void NestedOptionsAreConfiguredInPlace() {
        var run = new SampleRun();
        var setup = run.PageSetup;

        run.PageSetup(page => page.Landscape = true).Border(border => border.Width = 2);
        var border = run.Border;
        run.Border(border => border.Style = "double");

        Assert.Same(setup, run.PageSetup);
        Assert.True(setup.Landscape);
        Assert.NotNull(border);
        Assert.Same(border, run.Border);
        Assert.Equal(2, border.Width);
        Assert.Equal("double", border.Style);
    }

    [Fact]
    public void DelegatePropertyIsAssignedThroughConfigureRatherThanInvoked() {
        var run = new SampleRun().Configure(r => r.TextTransform = text => text.ToUpperInvariant());

        Assert.Equal("ABC", run.TextTransform!("abc"));
    }

    [Fact]
    public void ValueHandleWritesThroughToItsOwnerAndKeepsItsIdentity() {
        var sheet = new SampleSheet();

        SampleCell returned = sheet.Cell(1, 2).Value("Total").Bold();

        Assert.Equal(sheet.Cell(1, 2), returned);
        Assert.Equal("Total", sheet.Cell(1, 2).CurrentValue);
        Assert.True(sheet.Cell(1, 2).IsBold);
        Assert.False(sheet.Cell(1, 2).Bold(false).IsBold);
    }
}
