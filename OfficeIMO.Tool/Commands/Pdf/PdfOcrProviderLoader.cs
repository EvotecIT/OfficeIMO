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
            Assembly assembly = new ProviderAssemblyLoadContext(path).LoadFromAssemblyPath(path);
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

    /// <summary>Keeps optional provider dependencies beside their explicit assembly while sharing the host OCR contract.</summary>
    private sealed class ProviderAssemblyLoadContext : AssemblyLoadContext {
        private readonly AssemblyDependencyResolver _resolver;
        private readonly string _directory;
        internal ProviderAssemblyLoadContext(string path) : base(isCollectible: false) {
            _resolver = new AssemblyDependencyResolver(path);
            _directory = Path.GetDirectoryName(path)!;
        }
        [UnconditionalSuppressMessage("Trimming", "IL2026", Justification = "The separately deployed provider dependencies are not part of the trimmed application graph; NativeAOT is rejected before this context is created.")]
        protected override Assembly? Load(AssemblyName assemblyName) {
            if (assemblyName.Name == typeof(IOcrEngine).Assembly.GetName().Name) return typeof(IOcrEngine).Assembly;
            Assembly? shared = Default.Assemblies.FirstOrDefault(assembly => assembly.FullName == assemblyName.FullName);
            if (shared is not null) return shared;
            string? dependency = _resolver.ResolveAssemblyToPath(assemblyName);
            if (dependency is null && assemblyName.Name is { } name && name.IndexOfAny([Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar]) < 0) {
                string adjacent = Path.Combine(_directory, name + ".dll");
                if (File.Exists(adjacent)) dependency = adjacent;
            }
            return dependency is null ? null : LoadFromAssemblyPath(dependency);
        }
        protected override IntPtr LoadUnmanagedDll(string unmanagedDllName) {
            string? path = _resolver.ResolveUnmanagedDllToPath(unmanagedDllName);
            return path is null ? IntPtr.Zero : LoadUnmanagedDllFromPath(path);
        }
    }
}
