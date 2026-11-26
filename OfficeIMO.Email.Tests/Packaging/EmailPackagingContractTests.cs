using OfficeIMO.Email;
using System.Reflection;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class EmailPackagingContractTests {
    [Fact]
    public void ProductProjectAndAssembly_DeclareOnlyTheRequiredDependencies() {
        string projectPath = Path.Combine(GetRepositoryRoot(), "OfficeIMO.Email", "OfficeIMO.Email.csproj");
        XDocument project = XDocument.Load(projectPath);
        XNamespace ns = project.Root?.Name.Namespace ?? XNamespace.None;

        var packageReferences = project.Descendants(ns + "PackageReference")
            .Select(element => ((string?)element.Attribute("Include"), (string?)element.Attribute("Version"), (string?)element.Attribute("Condition")))
            .ToArray();
        Assert.Equal(
            new[] {
                ("System.Text.Encoding.CodePages", "8.0.0", "'$(TargetFramework)' == 'netstandard2.0'"),
                ("System.Memory", "4.6.3", "'$(TargetFramework)' == 'net472'")
            },
            packageReferences);
        string[] projectReferences = project.Descendants(ns + "ProjectReference")
            .Select(element => ((string?)element.Attribute("Include") ?? string.Empty).Replace('\\', '/'))
            .ToArray();
        Assert.Equal(
            new[] {
                "../OfficeIMO.Core/OfficeIMO.Core.csproj",
                "../OfficeIMO.Rtf/OfficeIMO.Rtf.csproj"
            },
            projectReferences);
        string[] linkedSources = project.Descendants(ns + "Compile")
            .Select(element => (string?)element.Attribute("Include"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!.Replace('\\', '/'))
            .ToArray();
        Assert.Empty(linkedSources);

        string[] references = typeof(EmailDocumentReader).Assembly.GetReferencedAssemblies()
            .Select(reference => reference.Name ?? string.Empty)
            .ToArray();
        Assert.DoesNotContain(references, name => string.Equals(name, "MsgKit", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, name => string.Equals(name, "MsgReader", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, name => string.Equals(name, "OpenMcdf", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, name => string.Equals(name, "RtfPipe", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, name => string.Equals(name, "MimeKit", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, name => string.Equals(name, "MailKit", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, name => string.Equals(name, "Microsoft.Maui.Graphics", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(references, name => string.Equals(name, "OfficeIMO.Rtf", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(references, name => string.Equals(name, "OfficeIMO.Core", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(references, name => string.Equals(name, "OfficeIMO.Security", StringComparison.OrdinalIgnoreCase));
#if NETFRAMEWORK
        Assert.DoesNotContain(references, name => string.Equals(name, "System.Text.Encoding.CodePages", StringComparison.OrdinalIgnoreCase));
#else
        Assert.Contains(references, name => string.Equals(name, "System.Text.Encoding.CodePages", StringComparison.OrdinalIgnoreCase));
#endif
    }

    [Fact]
    public void ExistingOptionsConstructorSignaturesRemainAvailable() {
        Assert.NotNull(typeof(EmailWriterOptions).GetConstructor(new[] {
            typeof(bool), typeof(bool), typeof(int), typeof(int), typeof(long)
        }));
        Assert.NotNull(typeof(EmailMailboxReaderOptions).GetConstructor(new[] {
            typeof(EmailReaderOptions), typeof(MboxVariant), typeof(int)
        }));
    }

    private static string GetRepositoryRoot() => EmailTestRepository.FindRoot();
}
