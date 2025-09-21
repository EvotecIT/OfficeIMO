using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Validates and optionally fixes Visio VSDX packages for common structural issues.
    /// </summary>
    public partial class VsdxPackageValidator {
        private static readonly XNamespace nsCore = "http://schemas.microsoft.com/office/visio/2011/1/core";
        private static readonly XNamespace nsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace nsDocRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace nsCT = "http://schemas.openxmlformats.org/package/2006/content-types";

        private const string RT_Document = "http://schemas.microsoft.com/visio/2010/relationships/document";
        private const string RT_Pages = "http://schemas.microsoft.com/visio/2010/relationships/pages";
        private const string RT_Page = "http://schemas.microsoft.com/visio/2010/relationships/page";
        private const string RT_Masters = "http://schemas.microsoft.com/visio/2010/relationships/masters";

        private const string CT_Document = "application/vnd.ms-visio.drawing.main+xml";
        private const string CT_Pages = "application/vnd.ms-visio.pages+xml";
        private const string CT_Page = "application/vnd.ms-visio.page+xml";

        private readonly List<string> _errors = new();
        private readonly List<string> _warnings = new();
        private readonly List<string> _fixes = new();

        /// <summary>
        /// Gets the list of validation errors.
        /// </summary>
        public IReadOnlyList<string> Errors => _errors.AsReadOnly();

        /// <summary>
        /// Gets the list of warnings encountered during validation.
        /// </summary>
        public IReadOnlyList<string> Warnings => _warnings.AsReadOnly();

        /// <summary>
        /// Gets the list of fixes applied when running in fix mode.
        /// </summary>
        public IReadOnlyList<string> Fixes => _fixes.AsReadOnly();

        /// <summary>
        /// Validates the specified VSDX file.
        /// </summary>
        /// <param name="inputPath">Path to the input VSDX file.</param>
        /// <returns><c>true</c> if no errors were found; otherwise, <c>false</c>.</returns>
        public bool ValidateFile(string inputPath) {
            _errors.Clear();
            _warnings.Clear();
            _fixes.Clear();

            if (!File.Exists(inputPath)) {
                _errors.Add($"File not found: {inputPath}");
                return false;
            }

            var tempPath = ExtractToTemp(inputPath);
            try {
                ValidatePackageStructure(tempPath);
                return _errors.Count == 0;
            } finally {
                Directory.Delete(tempPath, recursive: true);
            }
        }

        /// <summary>
        /// Validates and fixes the specified VSDX file.
        /// </summary>
        /// <param name="inputPath">Path to the input VSDX file.</param>
        /// <param name="outputPath">Path where the fixed file will be saved.</param>
        /// <returns><c>true</c> if the file was fixed successfully; otherwise, <c>false</c>.</returns>
        public bool FixFile(string inputPath, string outputPath) {
            _errors.Clear();
            _warnings.Clear();
            _fixes.Clear();

            if (!File.Exists(inputPath)) {
                _errors.Add($"File not found: {inputPath}");
                return false;
            }

            var tempPath = ExtractToTemp(inputPath);
            try {
                ValidateAndFix(tempPath);

                if (File.Exists(outputPath)) {
                    File.Delete(outputPath);
                }

                ZipFile.CreateFromDirectory(tempPath, outputPath, CompressionLevel.Optimal, includeBaseDirectory: false);
                return true;
            } finally {
                Directory.Delete(tempPath, recursive: true);
            }
        }

        private string ExtractToTemp(string inputPath) {
            var tempPath = Path.Combine(Path.GetTempPath(), "VsdxValidator_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tempPath);
            ZipFile.ExtractToDirectory(inputPath, tempPath);
            return tempPath;
        }

        private void ValidatePackageStructure(string tempPath) {
            ValidateContentTypes(tempPath, fix: false);
            ValidatePackageRelationships(tempPath, fix: false);
            ValidateDocumentRelationships(tempPath, fix: false);
            ValidatePagesStructure(tempPath, fix: false);
            ValidateStyleReferences(tempPath, fix: false);
        }

        private void ValidateAndFix(string tempPath) {
            ValidateContentTypes(tempPath, fix: true);
            ValidatePackageRelationships(tempPath, fix: true);
            ValidateDocumentRelationships(tempPath, fix: true);
            ValidatePagesStructure(tempPath, fix: true);
            ValidateStyleReferences(tempPath, fix: true);
        }
    }
}
