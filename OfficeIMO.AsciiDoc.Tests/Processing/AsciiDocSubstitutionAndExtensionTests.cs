namespace OfficeIMO.AsciiDoc.Tests;

public sealed class AsciiDocSubstitutionAndExtensionTests {
    [Fact]
    public void SubstitutionPlans_UseOrderedBlockDefaults() {
        const string source = "= Title\nParagraph\n----\ncode\n----\n++++\nraw\n++++\n";
        AsciiDocDocument document = AsciiDocDocument.Parse(source).Document;

        AsciiDocSubstitutionPlan heading = AsciiDocSubstitutionResolver.GetPlan(document.BlocksOfType<AsciiDocHeading>().Single());
        AsciiDocSubstitutionPlan paragraph = AsciiDocSubstitutionResolver.GetPlan(document.BlocksOfType<AsciiDocParagraph>().Single());
        AsciiDocDelimitedBlock[] delimited = document.BlocksOfType<AsciiDocDelimitedBlock>().ToArray();
        AsciiDocSubstitutionPlan listing = AsciiDocSubstitutionResolver.GetPlan(delimited[0]);
        AsciiDocSubstitutionPlan passthrough = AsciiDocSubstitutionResolver.GetPlan(delimited[1]);

        Assert.Equal("header", heading.Group);
        Assert.Equal("normal", paragraph.Group);
        Assert.Equal(new[] {
            AsciiDocSubstitutionType.SpecialCharacters,
            AsciiDocSubstitutionType.Quotes,
            AsciiDocSubstitutionType.Attributes,
            AsciiDocSubstitutionType.Replacements,
            AsciiDocSubstitutionType.Macros,
            AsciiDocSubstitutionType.PostReplacements
        }, paragraph.Substitutions);
        Assert.Equal(new[] { AsciiDocSubstitutionType.SpecialCharacters }, listing.Substitutions);
        Assert.Empty(passthrough.Substitutions);
    }

    [Fact]
    public void SubsOverride_IsParsedAndRetainsMandatedOrder() {
        const string source = "[subs=\"macros,quotes,attributes\"]\n----\ncode\n----\n";
        AsciiDocDelimitedBlock block = Assert.Single(AsciiDocDocument.Parse(source).Document.BlocksOfType<AsciiDocDelimitedBlock>());

        AsciiDocSubstitutionPlan plan = AsciiDocSubstitutionResolver.GetPlan(block);

        Assert.Equal("custom", plan.Group);
        Assert.Equal(new[] {
            AsciiDocSubstitutionType.Quotes,
            AsciiDocSubstitutionType.Attributes,
            AsciiDocSubstitutionType.Macros
        }, plan.Substitutions);
    }

    [Fact]
    public void ExplicitExtension_ReplacesOnlyRegisteredDirective() {
        const string source = ":name: OfficeIMO\nissue::42[label=Bug]\nunknown::value[]\n";
        var processor = new RecordingDirectiveProcessor();
        var registry = new AsciiDocExtensionRegistry().RegisterDirective("issue", processor);

        AsciiDocProcessingResult result = AsciiDocProcessor.Process(source, new AsciiDocProcessorOptions { Extensions = registry });

        Assert.Equal(":name: OfficeIMO\nIssue 42 for OfficeIMO\nunknown::value[]\n", result.ProcessedSource);
        Assert.NotNull(processor.Context);
        Assert.Equal("label=Bug", processor.Context!.AttributeList);
        Assert.Equal(2, processor.Context.Line);
    }

    [Fact]
    public void BuiltInDirectiveProcessors_CannotBeReplaced() {
        var registry = new AsciiDocExtensionRegistry();

        Assert.Throws<ArgumentException>(() => registry.RegisterDirective("include", new RecordingDirectiveProcessor()));
    }

    [Fact]
    public void ExtensionInvocationLimit_LeavesExcessDirectiveVisible() {
        var registry = new AsciiDocExtensionRegistry().RegisterDirective("issue", new RecordingDirectiveProcessor());

        AsciiDocProcessingResult result = AsciiDocProcessor.Process("issue::1[]\nissue::2[]\n", new AsciiDocProcessorOptions {
            Extensions = registry,
            MaximumExtensionInvocations = 1
        });

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "ADOCPROC008");
        Assert.EndsWith("issue::2[]\n", result.ProcessedSource, StringComparison.Ordinal);
    }

    private sealed class RecordingDirectiveProcessor : IAsciiDocDirectiveProcessor {
        internal AsciiDocDirectiveContext? Context { get; private set; }

        public AsciiDocDirectiveResult Process(AsciiDocDirectiveContext context) {
            Context = context;
            string owner = context.Attributes.GetValueOrDefault("name") ?? "unknown";
            return AsciiDocDirectiveResult.Replace("Issue " + context.Target + " for " + owner + "\n");
        }
    }
}
