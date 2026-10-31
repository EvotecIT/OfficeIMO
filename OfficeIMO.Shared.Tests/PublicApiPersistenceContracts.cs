using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class PublicApiPersistenceContracts {
    [Fact]
    public void VisioHasOneAuthoringThemeAndSeparatePackageMetadata() {
        Assembly assembly = typeof(VisioDocument).Assembly;

        Assert.Null(assembly.GetType("OfficeIMO.Visio.VisioTheme"));
        Assert.Null(assembly.GetType("OfficeIMO.Visio.Diagrams.VisioFlowchartTheme"));
        Assert.Null(assembly.GetType("OfficeIMO.Visio.Diagrams.VisioBlockDiagramTheme"));
        Assert.Equal(typeof(VisioPackageTheme), typeof(VisioDocument).GetProperty("PackageTheme")?.PropertyType);
        Assert.Null(typeof(VisioDocument).GetProperty("Theme"));
        AssertLayoutOnly(typeof(VisioFlowchartLayoutOptions));
        AssertLayoutOnly(typeof(VisioBlockDiagramLayoutOptions));

        Type[] builderTypes = {
            typeof(VisioFlowchartBuilder),
            typeof(VisioBlockDiagramBuilder)
        };
        foreach (Type builderType in builderTypes) {
            MethodInfo theme = Assert.Single(GetDeclaredMethods(builderType, "Theme"));
            Assert.Equal(typeof(VisioStyleTheme), Assert.Single(theme.GetParameters()).ParameterType);
        }
    }

    [Theory]
    [InlineData(typeof(WordEmbeddedDocument))]
    [InlineData(typeof(WordImage))]
    [InlineData(typeof(VisioStencilCatalog))]
    [InlineData(typeof(VisioStencilPreviewImageData))]
    public void BinaryArtifactsExposeCanonicalPersistenceVocabulary(Type artifactType) {
        string[] names = artifactType
            .GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Select(method => method.Name)
            .ToArray();

        Assert.Contains("Save", names);
        Assert.Contains("SaveAsync", names);
        Assert.Contains("ToBytes", names);
        Assert.Contains("ToStream", names);
    }

    private static void AssertImageSaveContract(Type type, params string[] methodNames) {
        foreach (string methodName in methodNames) {
            MethodInfo[] synchronous = GetDeclaredMethods(type, methodName);
            MethodInfo[] asynchronous = GetDeclaredMethods(type, methodName + "Async");

            Assert.NotEmpty(synchronous);
            Assert.NotEmpty(asynchronous);
            Assert.All(synchronous, method => Assert.Equal(typeof(OfficeImageExportResult), method.ReturnType));
            Assert.All(asynchronous, method => Assert.Equal(typeof(Task<OfficeImageExportResult>), method.ReturnType));
            Assert.Contains(synchronous, method => method.GetParameters().Any(parameter => parameter.ParameterType == typeof(string)));
            Assert.Contains(synchronous, method => method.GetParameters().Any(parameter => parameter.ParameterType == typeof(Stream)));
            Assert.Contains(asynchronous, method => method.GetParameters().Any(parameter => parameter.ParameterType == typeof(string)));
            Assert.Contains(asynchronous, method => method.GetParameters().Any(parameter => parameter.ParameterType == typeof(Stream)));
        }
    }

    private static void AssertLayoutOnly(Type type) {
        PropertyInfo[] properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly);

        Assert.NotEmpty(properties);
        Assert.All(properties, property => Assert.Equal(typeof(double), property.PropertyType));
    }

    private static MethodInfo[] GetDeclaredMethods(Type type, string name) => type
        .GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)
        .Where(method => method.Name == name)
        .ToArray();
}
