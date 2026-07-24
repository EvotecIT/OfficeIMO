using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>Controls which source slides remain in a presentation created from a template.</summary>
    public enum PowerPointTemplateSlideRetention {
        /// <summary>Keep every source slide.</summary>
        All,
        /// <summary>Remove every source slide while preserving masters, layouts, theme, and assets.</summary>
        None,
        /// <summary>Keep only zero-based indexes listed in <see cref="PowerPointTemplateCreationOptions.SourceSlideIndexes"/>.</summary>
        Selected
    }

    /// <summary>Options for creating an editable presentation from a `.pptx` or `.potx` template.</summary>
    public sealed class PowerPointTemplateCreationOptions {
        /// <summary>Source-slide retention policy.</summary>
        public PowerPointTemplateSlideRetention SlideRetention { get; set; } =
            PowerPointTemplateSlideRetention.None;

        /// <summary>Zero-based source slide indexes used when <see cref="SlideRetention"/> is Selected.</summary>
        public ISet<int> SourceSlideIndexes { get; } = new HashSet<int>();

        /// <summary>Marks retained source slides hidden, useful for reference or appendix templates.</summary>
        public bool HideRetainedSourceSlides { get; set; }

        /// <summary>Allows replacing an existing output file.</summary>
        public bool Overwrite { get; set; }
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Copies a `.pptx` or `.potx` template into a new editable `.pptx`, preserving its masters,
        ///     layouts, theme, relationships, and assets while applying an explicit source-slide policy.
        /// </summary>
        internal static PowerPointPresentation CreateFromTemplate(string templatePath, string outputPath,
            PowerPointTemplateCreationOptions? options = null) {
            if (string.IsNullOrWhiteSpace(templatePath)) {
                throw new ArgumentException("Template path cannot be empty.", nameof(templatePath));
            }
            if (string.IsNullOrWhiteSpace(outputPath)) {
                throw new ArgumentException("Output path cannot be empty.", nameof(outputPath));
            }

            string source = Path.GetFullPath(templatePath);
            string destination = Path.GetFullPath(outputPath);
            if (!File.Exists(source)) {
                throw new FileNotFoundException("PowerPoint template was not found.", source);
            }
            string sourceExtension = Path.GetExtension(source);
            if (!string.Equals(sourceExtension, ".pptx", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(sourceExtension, ".potx", StringComparison.OrdinalIgnoreCase)) {
                throw new NotSupportedException("Template input must be a .pptx or .potx file.");
            }
            if (!string.Equals(Path.GetExtension(destination), ".pptx", StringComparison.OrdinalIgnoreCase)) {
                throw new ArgumentException("Template output must use the .pptx extension.", nameof(outputPath));
            }
            if (string.Equals(source, destination, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidOperationException("Template source and output paths must be different.");
            }

            PowerPointTemplateCreationOptions resolved = options ?? new PowerPointTemplateCreationOptions();
            OfficeFileCommit.EnsureTargetDirectory(destination);
            string stagingPath = OfficeFileCommit.CreateStagingPath(destination);
            try {
                File.Copy(source, stagingPath, overwrite: false);

                if (string.Equals(sourceExtension, ".potx", StringComparison.OrdinalIgnoreCase)) {
                    using PresentationDocument document = PresentationDocument.Open(stagingPath, true);
                    document.ChangeDocumentType(PresentationDocumentType.Presentation);
                    document.Save();
                }

                using (PowerPointPresentation presentation = Load(stagingPath)) {
                    presentation.ApplyTemplateSlideRetention(resolved);
                    presentation.Save();
                }

                OfficeFileCommit.CommitTemporaryFile(
                    stagingPath,
                    destination,
                    resolved.Overwrite
                        ? OfficeFileCommit.ConflictPolicy.Replace
                        : OfficeFileCommit.ConflictPolicy.FailIfExists);
                stagingPath = string.Empty;
                return Load(destination);
            } finally {
                OfficeFileCommit.DeleteIfExists(stagingPath);
            }
        }

        /// <summary>Removes all slides while preserving reusable template masters, layouts, themes, and assets.</summary>
        public void ClearSlides() {
            ThrowIfDisposed();
            SlideIdList? slideIdList = PresentationRoot.SlideIdList;
            if (slideIdList != null) {
                foreach (SlideId slideId in slideIdList.Elements<SlideId>().ToList()) {
                    string? relationshipId = PowerPointUtils.GetRelationshipIdValue(slideId);
                    slideId.Remove();
                    if (string.IsNullOrWhiteSpace(relationshipId)) continue;
                    try {
                        OpenXmlPart part = _presentationPart.GetPartById(relationshipId!);
                        _presentationPart.DeletePart(part);
                    } catch (ArgumentOutOfRangeException) {
                        // Damaged source templates can carry stale slide relationships; the ID is still removed.
                    }
                }
            }

            _slides.Clear();
            SyncSectionsWithSlides();
            PresentationRoot.Save();
        }

        /// <summary>Adds a slide using a layout selected from a template inventory.</summary>
        public PowerPointSlide AddSlide(PowerPointTemplateLayoutInfo layout) {
            if (layout == null) throw new ArgumentNullException(nameof(layout));
            return AddSlide(layout.MasterIndex, layout.LayoutIndex);
        }

        private void ApplyTemplateSlideRetention(PowerPointTemplateCreationOptions options) {
            if (options.SlideRetention == PowerPointTemplateSlideRetention.None) {
                ClearSlides();
                return;
            }

            if (options.SlideRetention == PowerPointTemplateSlideRetention.Selected) {
                int originalCount = Slides.Count;
                foreach (int index in options.SourceSlideIndexes) {
                    if (index < 0 || index >= originalCount) {
                        throw new ArgumentOutOfRangeException(nameof(options.SourceSlideIndexes),
                            "Selected source slide index " + index + " is outside the template slide range.");
                    }
                }

                if (options.SourceSlideIndexes.Count == 0) {
                    ClearSlides();
                    return;
                }

                for (int slideIndex = originalCount - 1; slideIndex >= 0; slideIndex--) {
                    if (!options.SourceSlideIndexes.Contains(slideIndex)) RemoveSlide(slideIndex);
                }
            }

            if (options.HideRetainedSourceSlides) {
                for (int slideIndex = 0; slideIndex < Slides.Count; slideIndex++) {
                    Slides[slideIndex].Hidden = true;
                }
            }
        }
    }

    public partial class PowerPointSlide {
        /// <summary>
        ///     Creates a slide textbox from an inventoried text placeholder while preserving semantic type,
        ///     placeholder index, authored name, and bounds.
        /// </summary>
        public PowerPointTextBox AddTextToPlaceholder(PowerPointTemplatePlaceholderInfo placeholder, string text) {
            if (placeholder == null) throw new ArgumentNullException(nameof(placeholder));
            if (!placeholder.Bounds.HasValue) {
                throw new InvalidOperationException("Template placeholder '" + placeholder.Name +
                    "' does not define usable bounds.");
            }
            if (placeholder.Role == PowerPointTemplatePlaceholderRole.Image ||
                placeholder.Role == PowerPointTemplatePlaceholderRole.Chart ||
                placeholder.Role == PowerPointTemplatePlaceholderRole.Table) {
                throw new InvalidOperationException("Template placeholder '" + placeholder.Name +
                    "' is not a text placeholder. Use its Bounds for the matching native visual.");
            }

            PowerPointTextBox textBox = AddTextBox(text ?? string.Empty, placeholder.Bounds.Value);
            textBox.Name = placeholder.Name;
            textBox.PlaceholderType = placeholder.PlaceholderType;
            textBox.PlaceholderIndex = placeholder.PlaceholderIndex;
            return textBox;
        }
    }
}
