namespace OfficeIMO.Markdown;

/// <summary>
/// Helpers to scaffold common Markdown files such as README.
/// </summary>
public static class Scaffold {
    /// <summary>
    /// Builds a README document with common sections using provided options.
    /// </summary>
    public static MarkdownDoc Readme(string projectName, Action<ReadmeOptions> configure) {
        var opts = new ReadmeOptions(projectName);
        configure(opts);

        var md = MarkdownDoc.Create()
            .H1(projectName)
            .P(opts.DescriptionText ?? "");

        if (opts.BadgeImageList.Count > 0) {
            md.P(p => {
                for (int i = 0; i < opts.BadgeImageList.Count; i++) {
                    var b = opts.BadgeImageList[i];
                    p.ImageLink(b.Alt, b.ImageUrl, b.LinkUrl);
                }
            });
        } else if (opts.BadgeList.Count > 0) {
            md.P(p => {
                for (int i = 0; i < opts.BadgeList.Count; i++) {
                    var b = opts.BadgeList[i];
                    p.Link(b.Text, b.Url, b.Title);
                }
            });
        }

        if (!string.IsNullOrWhiteSpace(opts.NuGetPackageId)) {
            md.H2("Install")
              .Code("bash", string.IsNullOrWhiteSpace(opts.InstallCommand)
                    ? $"dotnet add package {opts.NuGetPackageId}"
                    : opts.InstallCommand!);
        }

        if (!string.IsNullOrWhiteSpace(opts.GettingStartedCode)) {
            md.H2("Getting started")
              .Code("csharp", opts.GettingStartedCode!);
        }

        if (opts.LinkList.Count > 0) {
            md.H2("Links").Ul(ul => {
                foreach (var l in opts.LinkList) ul.ItemLink(l.Text, l.Url);
            });
        }

        if (opts.LicenseMITRequested) {
            md.H2("License").P("MIT");
        }

        return md;
    }
}

// ReadmeOptions and BadgeBuilder moved to separate files for one-class-per-file.
