using System;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Markdown;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a Word document from Markdown.</summary>
/// <para>Returns a <see cref="WordDocument"/> or saves it to disk when <c>-OutputPath</c> is provided.</para>
/// <example>
///   <summary>Create a .docx from Markdown text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertFrom-OfficeWordMarkdown -Markdown '# Hello' -OutputPath .\hello.docx</code>
///   <para>Writes a Word document containing the supplied Markdown.</para>
/// </example>
/// <example>
///   <summary>Pipe a Markdown document into Word conversion.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeMarkdown -Path .\README.md | ConvertFrom-OfficeWordMarkdown</code>
///   <para>Returns a Word document instance for further edits.</para>
/// </example>
/// <example>
///   <summary>Insert Markdown into a Word template bookmark.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertFrom-OfficeWordMarkdown -Path .\SOP.md -TemplatePath .\Template.docx -BookmarkName MainContent -OutputPath .\SOP.docx</code>
///   <para>Copies the template and replaces the bookmark paragraph with generated Markdown content.</para>
/// </example>
[Cmdlet(VerbsData.ConvertFrom, "OfficeWordMarkdown", DefaultParameterSetName = ParameterSetMarkdown, SupportsShouldProcess = true)]
[Alias("ConvertFrom-WordMarkdown")]
[OutputType(typeof(WordDocument), typeof(FileInfo))]
public sealed class ConvertFromOfficeWordMarkdownCommand : AsyncPSCmdlet {
    private const string ParameterSetMarkdown = "Markdown";
    private const string ParameterSetPath = "Path";
    private const string ParameterSetDocument = "Document";

    /// <summary>Markdown text to convert.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ParameterSetName = ParameterSetMarkdown)]
    public string Markdown { get; set; } = string.Empty;

    /// <summary>Path to a Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Markdown document instance to convert.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Optional output path for the .docx file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Optional Word template document to copy before inserting Markdown content.</summary>
    [Parameter]
    [Alias("Template")]
    public string? TemplatePath { get; set; }

    /// <summary>Bookmark name that marks where Markdown content should be inserted in the template.</summary>
    [Parameter]
    public string? BookmarkName { get; set; }

    /// <summary>Block content control tag that marks where Markdown content should be inserted in the template.</summary>
    [Parameter]
    public string? ContentControlTag { get; set; }

    /// <summary>Block content control alias that marks where Markdown content should be inserted in the template.</summary>
    [Parameter]
    public string? ContentControlAlias { get; set; }

    /// <summary>Keep the target bookmark or content-control placeholder after inserting Markdown content.</summary>
    [Parameter]
    public SwitchParameter KeepPlaceholder { get; set; }

    /// <summary>Render YAML front matter as visible Word content. Template conversions hide front matter by default.</summary>
    [Parameter]
    public SwitchParameter RenderFrontMatter { get; set; }

    /// <summary>Optional font family applied during conversion.</summary>
    [Parameter]
    public string? FontFamily { get; set; }

    /// <summary>Base URI used to resolve relative links and images.</summary>
    [Parameter]
    [Alias("BasePath")]
    public string? BaseUri { get; set; }

    /// <summary>Allow local Markdown images to be inserted into the document.</summary>
    [Parameter]
    public SwitchParameter AllowLocalImages { get; set; }

    /// <summary>Restrict local images to one or more directories.</summary>
    [Parameter]
    public string[]? AllowedImageDirectory { get; set; }

    /// <summary>Allow remote HTTP(S) images to be downloaded and inserted.</summary>
    [Parameter]
    public SwitchParameter AllowRemoteImages { get; set; }

    /// <summary>Optional Markdown reader options used before Word conversion.</summary>
    [Parameter]
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>Named Markdown reader profile used when <see cref="ReaderOptions"/> is not supplied.</summary>
    [Parameter]
    public MarkdownReaderOptions.MarkdownDialectProfile? Profile { get; set; }

    /// <summary>Applies a built-in Markdown input normalization preset before parsing.</summary>
    [Parameter]
    public MarkdownInputNormalizationPreset? NormalizeInput { get; set; }

    /// <summary>Shared Markdown visual theme for generated Word output.</summary>
    [Parameter]
    public OfficeVisualThemeKind? Theme { get; set; }

    /// <summary>Allow data URI Markdown images to be embedded in Word output.</summary>
    [Parameter]
    public bool? AllowDataUriImages { get; set; }

    /// <summary>Maximum decoded size for one data URI image.</summary>
    [Parameter]
    public long? MaxDataUriImageBytes { get; set; }

    /// <summary>Prefer narrative paragraphs for isolated single-line definition-list patterns.</summary>
    [Parameter]
    public SwitchParameter PreferNarrativeSingleLineDefinitions { get; set; }

    /// <summary>Fit Markdown images to the page content width.</summary>
    [Parameter]
    public SwitchParameter FitImagesToPageContentWidth { get; set; }

    /// <summary>Fit Markdown images to the current content context width.</summary>
    [Parameter]
    public SwitchParameter FitImagesToContextWidth { get; set; }

    /// <summary>Optional hard cap for Markdown image width in pixels.</summary>
    [Parameter]
    public double? MaxImageWidthPixels { get; set; }

    /// <summary>Optional hard cap for Markdown image height in pixels.</summary>
    [Parameter]
    public double? MaxImageHeightPixels { get; set; }

    /// <summary>Optional hard cap for Markdown image width as a percentage of available content width.</summary>
    [Parameter]
    public double? MaxImageWidthPercentOfContent { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync() {
        WordDocument? document = null;

        try {
            if (FitImagesToPageContentWidth.IsPresent && FitImagesToContextWidth.IsPresent) {
                throw new ArgumentException("Use either -FitImagesToPageContentWidth or -FitImagesToContextWidth, not both.");
            }

            ValidateTemplateParameters();
            var options = BuildOptions();

            switch (ParameterSetName) {
                case ParameterSetPath: {
                        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
                        if (!File.Exists(resolvedPath)) {
                            throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
                        }

                        if (string.IsNullOrWhiteSpace(options.BaseUri)) {
                            options.BaseUri = BuildDirectoryUri(System.IO.Path.GetDirectoryName(resolvedPath) ?? Directory.GetCurrentDirectory());
                        }

                        document = await ConvertMarkdownTextAsync(File.ReadAllText(resolvedPath), options).ConfigureAwait(false);
                        break;
                    }
                case ParameterSetDocument:
                    document = await ConvertMarkdownDocumentAsync(Document, options).ConfigureAwait(false);
                    break;
                default:
                    if (string.IsNullOrWhiteSpace(Markdown)) {
                        throw new ArgumentException("Markdown content cannot be empty.", nameof(Markdown));
                    }

                    document = await ConvertMarkdownTextAsync(Markdown, options).ConfigureAwait(false);
                    break;
            }

            if (document == null) {
                throw new InvalidOperationException("Word document could not be created from Markdown.");
            }

            if (!string.IsNullOrWhiteSpace(OutputPath)) {
                var resolvedOutput = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
                if (!ShouldProcess(resolvedOutput, "Write Word document converted from Markdown")) {
                    document.Dispose();
                    return;
                }

                var directory = System.IO.Path.GetDirectoryName(resolvedOutput);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory)) {
                    Directory.CreateDirectory(directory);
                }

                try {
                    document.Save(resolvedOutput);
                } finally {
                    document.Dispose();
                }

                if (Open.IsPresent) {
                    FileOpenService.Open(resolvedOutput);
                }

                if (PassThru.IsPresent) {
                    WriteObject(new FileInfo(resolvedOutput));
                }
            } else {
                WriteObject(document);
            }
        } catch (Exception ex) {
            object? target = ParameterSetName switch {
                ParameterSetPath => Path,
                ParameterSetDocument => Document,
                _ => Markdown
            };
            WriteError(new ErrorRecord(ex, "MarkdownToWordFailed", ErrorCategory.InvalidOperation, target));
        }
    }

    private MarkdownToWordOptions BuildOptions() {
        if (ReaderOptions != null && Profile.HasValue) {
            throw new PSArgumentException("Specify either -ReaderOptions or -Profile, not both.");
        }

        var options = CreateOptions();
        options.AllowLocalImages = AllowLocalImages.IsPresent;
        options.OnWarning = WriteWarning;

        if (RenderFrontMatter.IsPresent) {
            options.RenderFrontMatter = true;
        }

        if (options is MarkdownToWordTemplateOptions templateOptions) {
            templateOptions.BookmarkName = NormalizeOptionalText(BookmarkName);
            templateOptions.ContentControlTag = NormalizeOptionalText(ContentControlTag);
            templateOptions.ContentControlAlias = NormalizeOptionalText(ContentControlAlias);
            templateOptions.ReplacePlaceholder = !KeepPlaceholder.IsPresent;
        }

        if (!string.IsNullOrWhiteSpace(FontFamily)) {
            options.FontFamily = FontFamily;
        }

        if (!string.IsNullOrWhiteSpace(BaseUri)) {
            options.BaseUri = ResolveBaseUri(BaseUri!);
        }

        if (Theme.HasValue) {
            options.Theme = MarkdownVisualTheme.Create(Theme.Value);
        }

        if (AllowDataUriImages.HasValue) {
            options.AllowDataUriImages = AllowDataUriImages.Value;
        }

        if (MaxDataUriImageBytes.HasValue) {
            options.MaxDataUriImageBytes = MaxDataUriImageBytes.Value;
        }

        if (PreferNarrativeSingleLineDefinitions.IsPresent) {
            options.PreferNarrativeSingleLineDefinitions = true;
        }

        var readerOptions = ReaderOptions ?? (Profile.HasValue
            ? MarkdownReaderOptions.CreateProfile(Profile.Value)
            : null);
        if (readerOptions != null && NormalizeInput.HasValue) {
            readerOptions.InputNormalization.ApplyPreset(NormalizeInput.Value);
        } else if (readerOptions == null && NormalizeInput.HasValue) {
            readerOptions = MarkdownReaderOptions.CreateOfficeIMOProfile();
            readerOptions.InputNormalization.ApplyPreset(NormalizeInput.Value);
        }

        options.ReaderOptions = readerOptions;
        TrySetOptionProperty(options, "FitImagesToPageContentWidth", FitImagesToPageContentWidth.IsPresent ? true : null);
        TrySetOptionProperty(options, "FitImagesToContextWidth", FitImagesToContextWidth.IsPresent ? true : null);
        TrySetOptionProperty(options, "MaxImageWidthPixels", MaxImageWidthPixels);
        TrySetOptionProperty(options, "MaxImageHeightPixels", MaxImageHeightPixels);
        TrySetOptionProperty(options, "MaxImageWidthPercentOfContent", MaxImageWidthPercentOfContent);

        if (AllowedImageDirectory != null) {
            foreach (var entry in AllowedImageDirectory) {
                if (string.IsNullOrWhiteSpace(entry)) {
                    continue;
                }

                options.AllowedImageDirectories.Add(SessionState.Path.GetUnresolvedProviderPathFromPSPath(entry));
            }
        }

        return options;
    }

    private MarkdownToWordOptions CreateOptions() {
        return string.IsNullOrWhiteSpace(TemplatePath)
            ? new MarkdownToWordOptions()
            : new MarkdownToWordTemplateOptions();
    }

    private Task<WordDocument> ConvertMarkdownTextAsync(string markdown, MarkdownToWordOptions options) {
        var markdownDocument = MarkdownReader.Parse(markdown, options.CreateReaderOptions());
        return ConvertMarkdownDocumentAsync(markdownDocument, options);
    }

    private async Task<WordDocument> ConvertMarkdownDocumentAsync(MarkdownDoc markdownDocument, MarkdownToWordOptions options) {
        if (AllowRemoteImages.IsPresent) {
            await MarkdownRemoteImageService.ConfigureResolverAsync(markdownDocument, options, CancelToken).ConfigureAwait(false);
        }

        if (options is not MarkdownToWordTemplateOptions templateOptions) {
            return markdownDocument.ToWordDocument(options);
        }

        var templateDocument = WordDocument.Load(ResolveTemplatePath(), new WordLoadOptions {
            AccessMode = DocumentAccessMode.ReadWrite,
            PersistenceMode = DocumentPersistenceMode.Explicit
        });
        return markdownDocument.ToWordDocument(templateDocument, templateOptions);
    }

    private void ValidateTemplateParameters() {
        var hasTemplateTarget = !string.IsNullOrWhiteSpace(BookmarkName)
            || !string.IsNullOrWhiteSpace(ContentControlTag)
            || !string.IsNullOrWhiteSpace(ContentControlAlias)
            || KeepPlaceholder.IsPresent;

        if (hasTemplateTarget && string.IsNullOrWhiteSpace(TemplatePath)) {
            throw new ArgumentException("Template insertion parameters require -TemplatePath.", nameof(TemplatePath));
        }
    }

    private string ResolveTemplatePath() {
        if (string.IsNullOrWhiteSpace(TemplatePath)) {
            throw new ArgumentException("Template path cannot be empty.", nameof(TemplatePath));
        }

        var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(TemplatePath);
        if (!File.Exists(resolvedPath)) {
            throw new FileNotFoundException($"Template file '{resolvedPath}' was not found.", resolvedPath);
        }

        return resolvedPath;
    }

    private static string? NormalizeOptionalText(string? value) {
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static void TrySetOptionProperty(object target, string propertyName, object? value) {
        if (value == null) {
            return;
        }

        var property = target.GetType().GetProperty(propertyName, BindingFlags.Instance | BindingFlags.Public);
        if (property == null || !property.CanWrite) {
            return;
        }

        var convertedValue = ConvertOptionValue(value, property.PropertyType);
        property.SetValue(target, convertedValue);
    }

    private static object? ConvertOptionValue(object value, Type propertyType) {
        var targetType = Nullable.GetUnderlyingType(propertyType) ?? propertyType;
        if (targetType.IsInstanceOfType(value)) {
            return value;
        }

        if (targetType.IsEnum && value is string text) {
            return Enum.Parse(targetType, text, ignoreCase: true);
        }

        return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
    }

    private string ResolveBaseUri(string value) {
        if (Uri.TryCreate(value, UriKind.Absolute, out var uri)) {
            return uri.ToString();
        }

        var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(value);
        if (Directory.Exists(resolved)) {
            return BuildDirectoryUri(resolved);
        }

        if (File.Exists(resolved)) {
            return new Uri(resolved).AbsoluteUri;
        }

        if (System.IO.Path.HasExtension(resolved)) {
            return new Uri(System.IO.Path.GetFullPath(resolved)).AbsoluteUri;
        }

        return BuildDirectoryUri(resolved);
    }

    private static string BuildDirectoryUri(string path) {
        var fullPath = System.IO.Path.GetFullPath(path);
        if (!fullPath.EndsWith(System.IO.Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
            && !fullPath.EndsWith(System.IO.Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)) {
            fullPath += System.IO.Path.DirectorySeparatorChar;
        }

        return new Uri(fullPath).AbsoluteUri;
    }
}