using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Drawing;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Tests.Foundation;

public sealed class FoundationBindingTests {
    [Theory]
    [InlineData("5cm", 1800000L)]
    [InlineData("12pt", 152400L)]
    [InlineData("1in", 914400L)]
    [InlineData("10mm", 360000L)]
    [InlineData("96px", 914400L)]
    [InlineData("914400emu", 914400L)]
    [InlineData(12, 152400L)]
    [InlineData(2.5, 31750L)]
    public void LengthsUseInvariantCoreParser(object input, long emus) {
        var saved = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            var length = (OfficeLength)new OfficeLengthTransformationAttribute().Transform(null!, PSObject.AsPSObject(input));
            Assert.Equal(emus, length.ToEmus());
        } finally { CultureInfo.CurrentCulture = saved; }
    }
    [Theory]
    [InlineData("2,5cm")][InlineData("NaN")][InlineData("Infinity")][InlineData("2yards")][InlineData(true)]
    public void InvalidLengthsFailDuringBinding(object input) => Assert.Throws<ArgumentTransformationMetadataException>(() => new OfficeLengthTransformationAttribute().Transform(null!, input));

    [Theory]
    [InlineData("#123456")][InlineData("#12345678")][InlineData("RebeccaPurple")][InlineData("Accent1")][InlineData("Success")]
    public void ColorsDelegateToCore(string input) {
        var color = (OfficeColor)new OfficeColorTransformationAttribute().Transform(null!, input);
        Assert.Equal(OfficeColor.Parse(input), color);
    }
    [Theory]
    [InlineData("NoColor")][InlineData("1")]
    public void InvalidColorsAreRejected(string input) => Assert.Throws<ArgumentTransformationMetadataException>(() => new OfficeColorTransformationAttribute().Transform(null!, input));

    [Theory]
    [InlineData("Modern")][InlineData("Classic")][InlineData("Minimal")][InlineData("Boardroom")][InlineData("Dark")]
    public void ThemesBindByBuiltInName(string input) {
        var source = OfficeThemeInput.From(input, _ => throw new InvalidOperationException("Names need no path lookup."));
        Assert.Equal(input, source.Theme!.Name);
        Assert.Same(source.Theme, source.Resolve(_ => throw new InvalidOperationException()));
    }
    [Fact]
    public void TemplatePathsStayDeferredToCoreReader() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".dotx");
        File.WriteAllText(path, "fixture");
        try {
            var source = OfficeThemeInput.From("template.dotx", _ => path);
            Assert.Equal(path, source.TemplatePath);
            Assert.Same(OfficeTheme.Dark, source.Resolve(p => { Assert.Equal(path, p); return OfficeTheme.Dark; }));
        } finally { File.Delete(path); }
    }
    [Fact]
    public void ThemeObjectKeepsItsIdentity() => Assert.Same(OfficeTheme.Modern, OfficeThemeInput.From(OfficeTheme.Modern, _ => "").Theme);
    [Fact]
    public void PasswordsWarnOnlyForPlainText() {
        var warnings = new List<string>();
        using SecureString password = OfficePasswordTransformationAttribute.ConvertPassword("sensitive", warnings.Add);
        Assert.Equal(9, password.Length);
        Assert.Single(warnings);
        Assert.DoesNotContain("sensitive", warnings[0]);
        warnings.Clear();
        Assert.Same(password, OfficePasswordTransformationAttribute.ConvertPassword(password, warnings.Add));
        Assert.Same(password, OfficePasswordTransformationAttribute.ConvertPassword(new PSCredential("user", password), warnings.Add));
        Assert.Empty(warnings);
    }
    [Fact]
    public void SheetSelectorDistinguishesNameFromPosition() {
        Assert.Equal("1", OfficeSheetSelector.From("1").Name);
        Assert.Equal(1, OfficeSheetSelector.From(1).Position);
        Assert.Throws<ArgumentTransformationMetadataException>(() => OfficeSheetSelector.From(0));
        Assert.Throws<ArgumentTransformationMetadataException>(() => OfficeSheetSelector.From(1.5));
        Assert.Throws<ArgumentTransformationMetadataException>(() => OfficeSheetSelector.From(new object()));
        var sheets = new[] { new object(), new object() };
        Assert.Same(sheets[0], OfficeSheetSelector.From(1).Resolve(sheets, _ => "1"));
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.From(3).Resolve(sheets, _ => "1"));
    }
    [Fact]
    public void OptionsAreCaseInsensitiveTypedAndRejectUnknownKeys() {
        var binder = new OfficeOptionsTransformationAttribute(typeof(FoundationOptions));
        var options = (FoundationOptions)binder.Transform(null!, new Hashtable { ["title"] = "test", ["Count"] = 4, ["Width"] = "5cm", ["Color"] = "Accent1" });
        Assert.Equal("test", options.Title);
        Assert.Equal(4, options.Count);
        Assert.Equal(OfficeLength.Centimeters(5), options.Width);
        Assert.Equal(OfficeColor.Accent1, options.Color);
        Assert.Same(options, binder.Transform(null!, options));
        Assert.Throws<ArgumentTransformationMetadataException>(() => binder.Transform(null!, new Hashtable { ["Missing"] = 3 }));
        Assert.Throws<ArgumentTransformationMetadataException>(() => binder.Transform(null!, new Hashtable { ["Count"] = "bad" }));
    }
    [Theory]
    [InlineData("Pdf", OfficeDocumentFormat.Pdf)][InlineData("word", OfficeDocumentFormat.Word)]
    public void FormatsBindBySharedEnumName(string input, OfficeDocumentFormat expected) => Assert.Equal(expected, new OfficeFormatTransformationAttribute().Transform(null!, input));
    [Theory]
    [InlineData("Unknown")][InlineData("99")][InlineData("Unsupported")]
    public void InvalidFormatsAreRejected(string input) => Assert.Throws<ArgumentTransformationMetadataException>(() => new OfficeFormatTransformationAttribute().Transform(null!, input));
    [Fact]
    public void SwitchesRetainOmissionAndExplicitFalse() {
        var parameters = new Dictionary<string, object>();
        Assert.Null(OfficeSwitch.Get(parameters, "Bold"));
        parameters["Bold"] = new SwitchParameter(false);
        Assert.False(OfficeSwitch.Get(parameters, "Bold"));
        Assert.True(OfficeSwitch.Enabled(parameters, "Freeze"));
        parameters["NoFreeze"] = new SwitchParameter(true);
        Assert.False(OfficeSwitch.Enabled(parameters, "Freeze"));
        parameters["NoFreeze"] = new SwitchParameter(false);
        Assert.True(OfficeSwitch.Enabled(parameters, "Freeze"));
    }
}
