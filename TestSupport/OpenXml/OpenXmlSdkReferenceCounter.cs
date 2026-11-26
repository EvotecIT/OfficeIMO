#if NET
using System.Collections.Immutable;
using System.Reflection;
using System.Reflection.Emit;
using System.Reflection.Metadata;
using System.Reflection.Metadata.Ecma335;
using System.Reflection.PortableExecutable;
using System.Text.RegularExpressions;

namespace OfficeIMO.TestSupport.OpenXml;

/// <summary>
/// Counts the source files of an engine assembly whose code uses Open XML SDK types, or <c>System.IO.Packaging</c> types
/// (the package API the SDK depends on, which isn't part of .NET), outside the folder that holds the engine's SDK bridge. It reads the compiled assembly and its portable PDB: a file counts when a method compiled from it
/// names an SDK type or member in its signature, its locals or its IL, or uses an engine field or method whose signature
/// does. That catches SDK types reached through <c>var</c> and member chains, which a search for <c>using</c> lines
/// misses. A file whose code (not its comments) names the namespace counts as well, for declarations that compile to no method.
/// </summary>
internal static class OpenXmlSdkReferenceCounter {
    /// <summary>An assembly family to count: its assemblies' name and the namespace its types live in.</summary>
    internal sealed class Dependency {
        private Dependency(string assemblyName, string namespaceName) {
            AssemblyName = assemblyName;
            NamespaceText = new Regex(@"\b" + Regex.Escape(namespaceName) + @"\b", RegexOptions.Compiled);
        }

        /// <summary>The Open XML SDK (<c>DocumentFormat.OpenXml</c> and its framework assemblies).</summary>
        internal static Dependency OpenXmlSdk { get; } = new("DocumentFormat.OpenXml", "DocumentFormat.OpenXml");

        /// <summary><c>System.IO.Packaging</c>, the NuGet package the SDK builds on.</summary>
        internal static Dependency SystemIOPackaging { get; } = new("System.IO.Packaging", "System.IO.Packaging");

        internal Regex NamespaceText { get; }

        internal bool Matches(string referencedAssembly) =>
            referencedAssembly.Equals(AssemblyName, StringComparison.Ordinal)
            || referencedAssembly.StartsWith(AssemblyName + ".", StringComparison.Ordinal);

        private string AssemblyName { get; }
    }

    /// <summary>The files outside <paramref name="bridgeFolder"/> that use the Open XML SDK.</summary>
    internal static IReadOnlyList<string> FindFiles(string assemblyPath, string projectDirectory, string bridgeFolder) =>
        FindFiles(assemblyPath, projectDirectory, bridgeFolder, Dependency.OpenXmlSdk);

    /// <summary>The files outside <paramref name="bridgeFolder"/> that use <paramref name="dependency"/>.</summary>
    internal static IReadOnlyList<string> FindFiles(string assemblyPath, string projectDirectory, string bridgeFolder, Dependency dependency) {
        string projectRoot = Path.GetFullPath(projectDirectory).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        string bridgeRoot = Path.Combine(projectRoot, bridgeFolder).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        var files = new SortedSet<string>(StringComparer.Ordinal);

        string? mappedRoot = FindMappedSourceRoot(projectRoot);
        foreach (string document in FindDocumentsUsing(assemblyPath, dependency)) {
            // Deterministic builds record source paths under /_/, the repository root's PathMap.
            string path = mappedRoot != null && document.StartsWith(MappedSourcePrefix, StringComparison.Ordinal)
                ? Path.Combine(mappedRoot, document.Substring(MappedSourcePrefix.Length))
                : document;
            string full = NormalizePath(path);
            if (!full.StartsWith(projectRoot, StringComparison.Ordinal)) continue;
            files.Add(full);
        }

        foreach (string source in Directory.EnumerateFiles(projectRoot, "*.cs", SearchOption.AllDirectories)) {
            string full = NormalizePath(source);
            string relative = full.Substring(projectRoot.Length);
            if (relative.StartsWith("obj" + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                || relative.StartsWith("bin" + Path.DirectorySeparatorChar, StringComparison.Ordinal)) continue;
            if (NamesNamespaceInCode(full, dependency)) files.Add(full);
        }

        return files
            .Where(file => !file.StartsWith(bridgeRoot, StringComparison.Ordinal) && File.Exists(file))
            .Select(file => file.Substring(projectRoot.Length).Replace(Path.DirectorySeparatorChar, '/'))
            .ToList();
    }

    // Comments that mention the namespace don't use it.
    private static bool NamesNamespaceInCode(string path, Dependency dependency) {
        foreach (string line in File.ReadLines(path)) {
            string code = line.TrimStart();
            if (code.StartsWith("//", StringComparison.Ordinal) || code.StartsWith("*", StringComparison.Ordinal)
                || code.StartsWith("/*", StringComparison.Ordinal)) continue;
            int comment = code.IndexOf("//", StringComparison.Ordinal);
            if (comment >= 0) code = code.Substring(0, comment);
            if (dependency.NamespaceText.IsMatch(code)) return true;
        }
        return false;
    }

    private const string MappedSourcePrefix = "/_/";

    // The folder Directory.Build.targets maps to /_/, the nearest one above the project that holds it.
    private static string? FindMappedSourceRoot(string projectRoot) {
        for (DirectoryInfo? directory = new DirectoryInfo(projectRoot); directory != null; directory = directory.Parent) {
            if (File.Exists(Path.Combine(directory.FullName, "Directory.Build.targets"))) return directory.FullName;
        }
        return null;
    }

    private static string NormalizePath(string path) =>
        Path.GetFullPath(path.Replace('\\', Path.DirectorySeparatorChar).Replace('/', Path.DirectorySeparatorChar));

    private static IEnumerable<string> FindDocumentsUsing(string assemblyPath, Dependency dependency) {
        using var stream = File.OpenRead(assemblyPath);
        using var peReader = new PEReader(stream);
        MetadataReader metadata = peReader.GetMetadataReader();
        MetadataReaderProvider pdbProvider = OpenPdb(peReader, assemblyPath);
        using (pdbProvider) {
            MetadataReader pdb = pdbProvider.GetMetadataReader();
            var analysis = new Analysis(metadata, dependency);
            var documents = new HashSet<string>(StringComparer.Ordinal);
            foreach (MethodDefinitionHandle handle in metadata.MethodDefinitions) {
                MethodDefinition method = metadata.GetMethodDefinition(handle);
                if (!analysis.MethodUsesSdk(method, peReader)) continue;
                MethodDebugInformation debug = pdb.GetMethodDebugInformation(handle);
                if (!debug.Document.IsNil) {
                    documents.Add(pdb.GetString(pdb.GetDocument(debug.Document).Name));
                    continue;
                }
                foreach (SequencePoint point in debug.GetSequencePoints()) {
                    documents.Add(pdb.GetString(pdb.GetDocument(point.Document).Name));
                }
            }
            return documents;
        }
    }

    private static MetadataReaderProvider OpenPdb(PEReader peReader, string assemblyPath) {
        foreach (DebugDirectoryEntry entry in peReader.ReadDebugDirectory()) {
            if (entry.Type == DebugDirectoryEntryType.EmbeddedPortablePdb) {
                return peReader.ReadEmbeddedPortablePdbDebugDirectoryData(entry);
            }
        }
        if (peReader.TryOpenAssociatedPortablePdb(assemblyPath, path => File.Exists(path) ? File.OpenRead(path) : null,
                out MetadataReaderProvider? provider, out _) && provider != null) {
            return provider;
        }
        throw new InvalidOperationException($"No portable PDB found for '{assemblyPath}'.");
    }

    private sealed class Analysis : ISignatureTypeProvider<bool, object?> {
        private static readonly Dictionary<short, OperandType> OperandTypes = typeof(OpCodes)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Select(field => (OpCode)field.GetValue(null)!)
            .ToDictionary(code => code.Value, code => code.OperandType);

        private readonly MetadataReader _metadata;
        private readonly HashSet<AssemblyReferenceHandle> _sdkAssemblies = new();
        private readonly Dictionary<EntityHandle, bool> _cache = new();

        internal Analysis(MetadataReader metadata, Dependency dependency) {
            _metadata = metadata;
            foreach (AssemblyReferenceHandle handle in metadata.AssemblyReferences) {
                string name = metadata.GetString(metadata.GetAssemblyReference(handle).Name);
                if (dependency.Matches(name)) _sdkAssemblies.Add(handle);
            }
        }

        internal bool MethodUsesSdk(MethodDefinition method, PEReader peReader) {
            if (_sdkAssemblies.Count == 0) return false;
            if (SignatureUsesSdk(method.Signature, isMethod: true)) return true;
            if (method.RelativeVirtualAddress == 0) return false;
            MethodBodyBlock body = peReader.GetMethodBody(method.RelativeVirtualAddress);
            if (!body.LocalSignature.IsNil) {
                StandaloneSignature locals = _metadata.GetStandaloneSignature(body.LocalSignature);
                if (locals.DecodeLocalSignature(this, null).Any(local => local)) return true;
            }
            // A typed catch names its exception type in the exception table, not in the IL.
            foreach (ExceptionRegion region in body.ExceptionRegions) {
                if (region.Kind == ExceptionRegionKind.Catch && TokenUsesSdk(region.CatchType)) return true;
            }
            return IlUsesSdk(body.GetILReader());
        }

        private bool IlUsesSdk(BlobReader il) {
            while (il.RemainingBytes > 0) {
                short code = il.ReadByte();
                if (code == 0xFE) code = unchecked((short)(0xFE00 | il.ReadByte()));
                if (!OperandTypes.TryGetValue(code, out OperandType operand)) return false;
                switch (operand) {
                    case OperandType.InlineNone: break;
                    case OperandType.ShortInlineBrTarget:
                    case OperandType.ShortInlineI:
                    case OperandType.ShortInlineVar: il.Offset += 1; break;
                    case OperandType.InlineVar: il.Offset += 2; break;
                    case OperandType.InlineI8:
                    case OperandType.InlineR: il.Offset += 8; break;
                    case OperandType.InlineSwitch:
                        int targets = il.ReadInt32();
                        il.Offset += 4 * targets;
                        break;
                    case OperandType.InlineField:
                    case OperandType.InlineMethod:
                    case OperandType.InlineTok:
                    case OperandType.InlineType:
                    case OperandType.InlineSig:
                        if (TokenUsesSdk(MetadataTokens.EntityHandle(il.ReadInt32()))) return true;
                        break;
                    default: il.Offset += 4; break;
                }
            }
            return false;
        }

        private bool TokenUsesSdk(EntityHandle handle) {
            if (_cache.TryGetValue(handle, out bool known)) return known;
            _cache[handle] = false;
            bool result = handle.Kind switch {
                HandleKind.TypeReference => TypeReferenceIsSdk((TypeReferenceHandle)handle),
                HandleKind.TypeDefinition => false,
                HandleKind.TypeSpecification => _metadata.GetTypeSpecification((TypeSpecificationHandle)handle).DecodeSignature(this, null),
                HandleKind.MemberReference => MemberReferenceUsesSdk((MemberReferenceHandle)handle),
                HandleKind.MethodDefinition => SignatureUsesSdk(_metadata.GetMethodDefinition((MethodDefinitionHandle)handle).Signature, isMethod: true),
                HandleKind.FieldDefinition => SignatureUsesSdk(_metadata.GetFieldDefinition((FieldDefinitionHandle)handle).Signature, isMethod: false),
                HandleKind.MethodSpecification => MethodSpecificationUsesSdk((MethodSpecificationHandle)handle),
                HandleKind.StandaloneSignature => _metadata.GetStandaloneSignature((StandaloneSignatureHandle)handle).GetKind() == StandaloneSignatureKind.Method
                    && SignatureUsesSdk(_metadata.GetStandaloneSignature((StandaloneSignatureHandle)handle).Signature, isMethod: true),
                _ => false,
            };
            _cache[handle] = result;
            return result;
        }

        private bool MemberReferenceUsesSdk(MemberReferenceHandle handle) {
            MemberReference member = _metadata.GetMemberReference(handle);
            if (!member.Parent.IsNil && TokenUsesSdk(member.Parent)) return true;
            return SignatureUsesSdk(member.Signature, member.GetKind() == MemberReferenceKind.Method);
        }

        private bool MethodSpecificationUsesSdk(MethodSpecificationHandle handle) {
            MethodSpecification specification = _metadata.GetMethodSpecification(handle);
            return TokenUsesSdk(specification.Method) || specification.DecodeSignature(this, null).Any(argument => argument);
        }

        private bool SignatureUsesSdk(BlobHandle signature, bool isMethod) {
            var decoder = new SignatureDecoder<bool, object?>(this, _metadata, null);
            BlobReader reader = _metadata.GetBlobReader(signature);
            if (isMethod) {
                MethodSignature<bool> method = decoder.DecodeMethodSignature(ref reader);
                return method.ReturnType || method.ParameterTypes.Any(parameter => parameter);
            }
            return decoder.DecodeFieldSignature(ref reader);
        }

        private bool TypeReferenceIsSdk(TypeReferenceHandle handle) {
            for (int depth = 0; depth < 32; depth++) {
                TypeReference reference = _metadata.GetTypeReference(handle);
                EntityHandle scope = reference.ResolutionScope;
                if (scope.Kind == HandleKind.AssemblyReference) return _sdkAssemblies.Contains((AssemblyReferenceHandle)scope);
                if (scope.Kind != HandleKind.TypeReference) return false;
                handle = (TypeReferenceHandle)scope;
            }
            return false;
        }

        public bool GetArrayType(bool elementType, ArrayShape shape) => elementType;
        public bool GetByReferenceType(bool elementType) => elementType;
        public bool GetFunctionPointerType(MethodSignature<bool> signature) => signature.ReturnType || signature.ParameterTypes.Any(type => type);
        public bool GetGenericInstantiation(bool genericType, ImmutableArray<bool> typeArguments) => genericType || typeArguments.Any(type => type);
        public bool GetGenericMethodParameter(object? genericContext, int index) => false;
        public bool GetGenericTypeParameter(object? genericContext, int index) => false;
        public bool GetModifiedType(bool modifier, bool unmodifiedType, bool isRequired) => unmodifiedType;
        public bool GetPinnedType(bool elementType) => elementType;
        public bool GetPointerType(bool elementType) => elementType;
        public bool GetPrimitiveType(PrimitiveTypeCode typeCode) => false;
        public bool GetSZArrayType(bool elementType) => elementType;
        public bool GetTypeFromDefinition(MetadataReader reader, TypeDefinitionHandle handle, byte rawTypeKind) => false;
        public bool GetTypeFromReference(MetadataReader reader, TypeReferenceHandle handle, byte rawTypeKind) => TokenUsesSdk(handle);
        public bool GetTypeFromSpecification(MetadataReader reader, object? genericContext, TypeSpecificationHandle handle, byte rawTypeKind) => TokenUsesSdk(handle);
    }
}
#endif
