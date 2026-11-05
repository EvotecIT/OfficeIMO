using System.Collections;
using System.Globalization;
using System.Management.Automation;
using System.Security;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
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
    [Fact]
    public void LengthCollectionsBindEachInvariantValueAndRejectInvalidElements() {
        var binder = new OfficeLengthTransformationAttribute();
        var lengths = Assert.IsType<OfficeLength[]>(binder.Transform(null!, new object[] { "1in", 12, PSObject.AsPSObject("5cm") }));
        Assert.Equal(new[] { 914400L, 152400L, 1800000L }, lengths.Select(value => value.ToEmus()));
        Assert.Throws<ArgumentTransformationMetadataException>(() => binder.Transform(null!, new object[] { "1in", "bad" }));
        Assert.Throws<ArgumentTransformationMetadataException>(() => binder.Transform(null!, new object?[] { "1in", null }!));
    }
    [Fact]
    public void NativeWorksheetOptionTablesUseTheCanonicalDeferredSelectors() {
        var binder = new OfficeOptionsTransformationAttribute(typeof(OfficeIMO.Excel.Pdf.ExcelPdfExportOptions));
        var options = Assert.IsType<OfficeIMO.Excel.Pdf.ExcelPdfExportOptions>(binder.Transform(null!, new Hashtable {
            ["Sheets"] = new object[] { "Annual*", 2 }
        }));
        Assert.Equal("Annual*", options.Sheets![0].Name);
        Assert.Equal(2, options.Sheets[1].Position);
        var single = Assert.IsType<OfficeIMO.Excel.Pdf.ExcelPdfExportOptions>(binder.Transform(null!, new Hashtable { ["Sheets"] = "Only" }));
        Assert.Equal("Only", Assert.Single(single.Sheets!).Name);
    }
    [Fact]
    public void PdfReadOptionsBindAsIndependentNativePolicySnapshots() {
        var source = new PSWriteOffice.Models.Pdf.PdfReadOptions {
            LoadOptions = new OfficeIMO.Pdf.PdfLoadOptions { Password = "open", PermissionPolicy = OfficeIMO.Pdf.PdfPermissionPolicy.IgnoreRestrictions }
        };
        var policy = Assert.IsType<OfficeIMO.Pdf.PdfLoadOptions>(new OfficeOptionsTransformationAttribute(typeof(OfficeIMO.Pdf.PdfLoadOptions)).Transform(null!, source));
        Assert.NotSame(source.LoadOptions, policy);
        Assert.Equal("open", policy.Password);
        Assert.Equal(OfficeIMO.Pdf.PdfPermissionPolicy.IgnoreRestrictions, policy.PermissionPolicy);
        Assert.NotSame(source.LoadOptions.Limits, policy.Limits);
        source.LoadOptions = new OfficeIMO.Pdf.PdfLoadOptions { Password = "different" };
        Assert.Equal("open", policy.Password);
    }
    [Fact]
    public void AccountSecretsKeepCredentialIdentityWithoutSharingPasswordMetadata() {
        using var secret = new SecureString();
        secret.AppendChar('x');
        var binder = new OfficePasswordTransformationAttribute { PreserveUserName = true };
        using var first = Assert.IsType<SecureString>(binder.Transform(null!, new PSCredential("first@example.com", secret)));
        using var second = Assert.IsType<SecureString>(binder.Transform(null!, new PSCredential("second@example.com", secret)));
        Assert.NotSame(secret, first);
        Assert.Equal("first@example.com", OfficePasswordTransformationAttribute.ResolveUserName(first, null));
        Assert.Equal("second@example.com", OfficePasswordTransformationAttribute.ResolveUserName(second, null));
        Assert.Equal("first@example.com", OfficePasswordTransformationAttribute.ResolveUserName(first, "FIRST@example.com"));
        Assert.Throws<PSArgumentException>(() => OfficePasswordTransformationAttribute.ResolveUserName(first, "second@example.com"));
        Assert.Equal(1, secret.Length);
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
    public void EveryCmdletPasswordParameterBindsThroughThePasswordTransformation() {
        // PasswordVariable, LegacyPasswordHash and similar references end differently and are not passwords.
        var inconsistent = typeof(OfficePasswordTransformationAttribute).Assembly.GetTypes()
            .Where(type => typeof(Cmdlet).IsAssignableFrom(type) && !type.IsAbstract)
            .SelectMany(type => type.GetProperties().Select(property => (type, property)))
            .Where(item => item.property.Name.EndsWith("Password", StringComparison.Ordinal) &&
                item.property.IsDefined(typeof(ParameterAttribute), true))
            .Where(item => !item.property.IsDefined(typeof(OfficePasswordTransformationAttribute), true) ||
                (item.property.PropertyType != typeof(SecureString) && item.property.PropertyType != typeof(SecureString[])))
            .Select(item => item.type.Name + "." + item.property.Name)
            .ToArray();
        Assert.Empty(inconsistent);
    }
    [Fact]
    public void SheetSelectorDistinguishesNameFromPosition() {
        var binder = new OfficeSheetTransformationAttribute();
        var byName = (OfficeSheetSelector)binder.Transform(null!, PSObject.AsPSObject("1"));
        var byPosition = (OfficeSheetSelector)binder.Transform(null!, PSObject.AsPSObject(1));
        Assert.Equal("1", byName.Name);
        Assert.Equal(1, byPosition.Position);
        Assert.Throws<ArgumentTransformationMetadataException>(() => binder.Transform(null!, 0));
        Assert.Throws<ArgumentTransformationMetadataException>(() => binder.Transform(null!, 1.5));
        Assert.Throws<ArgumentTransformationMetadataException>(() => binder.Transform(null!, new object()));
        using var book = ExcelDocument.Create();
        book.AddSheet("Report");
        book.AddSheet("1");
        book.AddSheet("Sales West");
        book.AddSheet("Sales East");
        Assert.Equal("1", byName.Resolve(book.Sheets, sheet => sheet.Name).Name);
        Assert.Equal("Report", byPosition.Resolve(book.Sheets, sheet => sheet.Name).Name);
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.From(5).Resolve(book.Sheets, sheet => sheet.Name));
        var mixed = (OfficeSheetSelector[])binder.Transform(null!, PSObject.AsPSObject(new object[] {
            PSObject.AsPSObject("sALes *"), PSObject.AsPSObject(4), PSObject.AsPSObject(1)
        }));
        Assert.Equal(new[] { "Report", "Sales West", "Sales East" },
            OfficeSheetSelector.ResolveMany(book, mixed).Select(sheet => sheet.Name));
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveMany(book,
            new[] { OfficeSheetSelector.From("Missing*") }));
    }

    [Fact]
    public void NativeSheetMembershipSurvivesWrapperRefreshButRejectsForeignAndReplacedSheets() {
        using var book = ExcelDocument.Create();
        var original = book.AddSheet("Data");
        var binder = new OfficeSheetTransformationAttribute();
        var selector = (OfficeSheetSelector)binder.Transform(null!, PSObject.AsPSObject(original));
        book.AddSheet("Other");
        Assert.Equal("Data", selector.Resolve(book.Sheets, sheet => sheet.Name).Name);
        book.SheetCachingEnabled = false;
        var selected = selector.Resolve(book.Sheets, sheet => sheet.Name);
        selected.Cell(1, 1).SetValue("Same native worksheet");
        Assert.Equal("Same native worksheet", original.Cell(1, 1).GetValue<string>());
        using var foreign = ExcelDocument.Create();
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.From(foreign.AddSheet("Data"))
            .Resolve(book.Sheets, sheet => sheet.Name));
        book.RemoveWorksheet("Data");
        book.AddSheet("Data");
        Assert.Throws<ArgumentException>(() => selector.Resolve(book.Sheets, sheet => sheet.Name));
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
