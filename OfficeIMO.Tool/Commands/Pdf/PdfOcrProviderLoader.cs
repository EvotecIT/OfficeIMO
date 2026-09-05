using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Loader;
using OfficeIMO.Ocr;

namespace OfficeIMO.Tool.Commands.Pdf;

internal static class PdfOcrProviderLoader {
    private const int MaximumProviderAssemblies = 32;

    [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The separately deployed provider assembly is not part of the trimmed application graph; NativeAOT is rejected before loading.")]
    [UnconditionalSuppressMessage("Trimming", "IL2070", Justification = "Provider types come from an untrimmed external assembly and are required to expose a public parameterless constructor.")]
    [UnconditionalSuppressMessage("Trimming", "IL2072", Justification = "Provider types come from an untrimmed external assembly and are required to expose a public parameterless constructor.")]
    internal static void LoadExplicitAssemblies(OcrEngineCatalog catalog, IEnumerable<string> assemblyPaths) {
        ArgumentNullException.ThrowIfNull(catalog);
        ArgumentNullException.ThrowIfNull(assemblyPaths);
        string[] paths = assemblyPaths.Take(MaximumProviderAssemblies + 1).ToArray();
        if (paths.Length > MaximumProviderAssemblies) {
            throw new ArgumentException("OCR provider assembly paths cannot exceed " + MaximumProviderAssemblies + " entries.", nameof(assemblyPaths));
        }
        if (paths.Length > 0 && !RuntimeFeature.IsDynamicCodeSupported) {
            throw new PlatformNotSupportedException("Loading OCR provider assemblies is unavailable in NativeAOT. Register a statically linked provider in OcrEngineCatalog through the host API.");
        }
        foreach (string suppliedPath in paths) {
            string path = Path.GetFullPath(suppliedPath);
            if (!File.Exists(path)) throw new FileNotFoundException("OCR provider assembly was not found.", path);
            Assembly assembly = AssemblyLoadContext.Default.LoadFromAssemblyPath(path);
            Type[] providerTypes;
            try {
                providerTypes = assembly.GetTypes();
            } catch (ReflectionTypeLoadException exception) {
                providerTypes = exception.Types.Where(static type => type is not null).Cast<Type>().ToArray();
            }
            foreach (Type type in providerTypes.Where(static type =>
                         type.IsClass && !type.IsAbstract && type.IsPublic &&
                         typeof(IOcrEngineProvider).IsAssignableFrom(type) &&
                         type.GetConstructor(Type.EmptyTypes) is not null)) {
                var provider = (IOcrEngineProvider?)Activator.CreateInstance(type)
                    ?? throw new InvalidOperationException("Could not create OCR provider type '" + type.FullName + "'.");
                catalog.Register(provider);
            }
        }
    }
}
