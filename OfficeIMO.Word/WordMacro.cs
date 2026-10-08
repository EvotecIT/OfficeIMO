using OfficeIMO.Core.Internal;
using DocumentFormat.OpenXml.Packaging;
using System.Diagnostics.CodeAnalysis;
using System.IO.Packaging;
using System.Reflection;

namespace OfficeIMO.Word {
    /// <summary>
    /// Represents a single macro module within a document.
    /// </summary>
    /// <remarks>
    /// Instances are returned by <see cref="WordDocument.Macros"/> and can be
    /// removed individually using <see cref="Remove"/>.
    /// </remarks>
    public class WordMacro {
        private readonly WordDocument _document;
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.NonPublicProperties)]
        private static readonly Type OpenXmlPackageType = typeof(OpenXmlPackage);

        /// <summary>
        /// Gets the macro module name.
        /// </summary>
        public string Name { get; }

        internal WordMacro(WordDocument document, string name) {
            _document = document ?? throw new ArgumentNullException(nameof(document));
            Name = name ?? throw new ArgumentNullException(nameof(name));
        }

        /// <summary>
        /// Removes this macro module from the document.
        /// </summary>
        public void Remove() {
            WordMacro.RemoveMacro(_document, Name);
        }

        /// <summary>
        /// Enumerates all macro modules in the specified document.
        /// </summary>
        /// <param name="document">Document to inspect.</param>
        /// <returns>List of <see cref="WordMacro"/> instances.</returns>
        internal static IReadOnlyList<WordMacro> GetMacros(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (!document.HasMacros) return new List<WordMacro>();

            var mainPart = document._wordprocessingDocument.MainDocumentPart;
            var vbaPart = mainPart?.VbaProjectPart;
            if (vbaPart == null) {
                return new List<WordMacro>();
            }

            byte[] bytes = OfficeIMO.OpenXml.Internal.OfficeVbaProjectPartEditor.Read(vbaPart, OfficeVbaProjectInfo.DefaultMaximumProjectBytes);
            IReadOnlyList<string> names = OfficeVbaProjectInspector.GetModuleNames(bytes);
            var modules = new List<WordMacro>(names.Count);
            foreach (var name in names) {
                modules.Add(new WordMacro(document, name));
            }
            return modules;
        }

        /// <summary>
        /// Adds a VBA project from a file to the specified document.
        /// </summary>
        /// <param name="document">Target document.</param>
        /// <param name="filePath">Path to a <c>vbaProject.bin</c> file.</param>
        internal static void AddMacro(WordDocument document, string filePath) {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException(nameof(filePath));
            if (!File.Exists(filePath)) throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);

            AddMacro(document, File.ReadAllBytes(filePath));
        }

        /// <summary>
        /// Adds a VBA project from a byte array to the specified document.
        /// </summary>
        /// <param name="document">Target document.</param>
        /// <param name="data">VBA project data.</param>
        internal static void AddMacro(WordDocument document, byte[] data) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (data == null || data.Length == 0) throw new ArgumentNullException(nameof(data));

            var main = document._wordprocessingDocument.MainDocumentPart;
            if (main == null) {
                throw new InvalidOperationException("The document does not contain a main document part.");
            }

            if (main.VbaProjectPart != null) {
                main.DeletePart(main.VbaProjectPart);
            }
            var vbaPart = main.AddNewPart<VbaProjectPart>();
            using var stream = new MemoryStream(data);
            vbaPart.FeedData(stream);
        }

        /// <summary>
        /// Returns the raw VBA project from the given document.
        /// </summary>
        /// <param name="document">Document containing macros.</param>
        /// <returns>Byte array with the macros or <c>null</c> when absent.</returns>
        internal static byte[]? ExtractMacros(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var mainPart = document._wordprocessingDocument.MainDocumentPart;
            var vbaPart = mainPart?.VbaProjectPart;
            if (vbaPart == null) return null;
            using var ms = new MemoryStream();
            using var partStream = vbaPart.GetStream();
            partStream.CopyTo(ms);
            return ms.ToArray();
        }

        /// <summary>
        /// Saves the VBA project from a document to the specified file.
        /// </summary>
        /// <param name="document">Source document.</param>
        /// <param name="filePath">Destination path.</param>
        internal static void SaveMacros(WordDocument document, string filePath) {
            if (string.IsNullOrEmpty(filePath)) throw new ArgumentNullException(nameof(filePath));
            var data = ExtractMacros(document);
            if (data == null) return;
            OfficeFileCommit.WriteAllBytes(filePath, data);
        }

        /// <summary>
        /// Removes a single macro module from a document.
        /// </summary>
        /// <param name="document">Document to modify.</param>
        /// <param name="name">Module name to remove.</param>
        internal static void RemoveMacro(WordDocument document, string name) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (string.IsNullOrEmpty(name)) throw new ArgumentNullException(nameof(name));
            if (!document.HasMacros) return;

            var mainPart = document._wordprocessingDocument.MainDocumentPart;
            var vbaPart = mainPart?.VbaProjectPart;
            if (vbaPart == null) {
                return;
            }

            byte[] data = OfficeIMO.OpenXml.Internal.OfficeVbaProjectPartEditor.Read(vbaPart, OfficeVbaProjectInfo.DefaultMaximumProjectBytes);
            if (!OfficeCompoundFileReader.TryRead(data, out OfficeCompoundFile? compound, out string? error) || compound == null) {
                throw new InvalidDataException(error ?? "The VBA project is not a valid compound file.");
            }
            byte[] changed;
            int remaining;
            if (compound.Streams.TryGetValue("VBA/dir", out byte[]? directory) && directory.Length > 0) {
                OfficeVbaProject project = OfficeVbaProject.Load(data);
                if (!project.Modules.Any(module => module.Name.Equals(name, StringComparison.OrdinalIgnoreCase))) return;
                project.DeleteModule(name);
                changed = project.Write().GetBytes();
                remaining = project.Modules.Count;
            } else {
                // Preserve the established opaque-stream removal contract for incomplete imported blobs.
                string? target = compound.Entries.FirstOrDefault(entry => entry.IsStream
                    && entry.Path.Equals("VBA/" + name, StringComparison.OrdinalIgnoreCase))?.Path;
                if (target == null || !OfficeVbaProjectInspector.GetModuleNames(data).Contains(name, StringComparer.OrdinalIgnoreCase)) return;
                changed = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]>(), new[] { target }, OfficeVbaProjectInfo.DefaultMaximumProjectBytes);
                remaining = OfficeVbaProjectInspector.GetModuleNames(changed).Count;
            }
            // Check signature policy before either replacing or deleting the project part.
            OfficeIMO.OpenXml.Internal.OfficeVbaProjectPartEditor.Apply(vbaPart, changed, new OfficeVbaWriteOptions());
            if (remaining == 0) RemoveMacros(document);
        }

        /// <summary>
        /// Removes the entire VBA project from the document.
        /// </summary>
        /// <param name="document">Document to modify.</param>
        internal static void RemoveMacros(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var main = document._wordprocessingDocument.MainDocumentPart;
            if (main?.VbaProjectPart == null) {
                return;
            }

            var vbaUri = main.VbaProjectPart.Uri;
            var packageProp = OpenXmlPackageType.GetProperty("Package", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (packageProp?.GetValue(document._wordprocessingDocument) is Package package) {
                foreach (var rel in package.GetRelationshipsByType("http://schemas.microsoft.com/office/2006/relationships/vbaProject").ToList()) {
                    if (rel.TargetUri == vbaUri) {
                        package.DeleteRelationship(rel.Id);
                    }
                }
            }
            main.DeletePart(main.VbaProjectPart);
        }

    }
}
