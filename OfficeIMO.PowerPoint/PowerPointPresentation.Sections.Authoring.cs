using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>Creates a native section boundary that the next appended slide will occupy.</summary>
        internal PowerPointSectionInfo StartSectionAtNextSlide(string name) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Section name cannot be null or empty.", nameof(name));
            List<SlideId> slides = PresentationRoot?.SlideIdList?.Elements<SlideId>().ToList() ?? new List<SlideId>();
            P14.SectionList list = EnsureSectionList(slides);
            P14.Section section = CreateSection(name, Array.Empty<uint>());
            list.Append(section);
            return BuildSectionInfo(section, slides);
        }

        /// <summary>Restores only section metadata when a collected slide or column is rejected.</summary>
        internal Action CaptureSectionRollback() {
            var root = PresentationRoot ?? throw new InvalidOperationException("Presentation is missing.");
            PresentationExtensionList? originalList = root.GetFirstChild<PresentationExtensionList>();
            var original = (PresentationExtension?)originalList?.Elements<PresentationExtension>()
                .FirstOrDefault(extension => string.Equals(extension.Uri?.Value, SectionListUri, StringComparison.Ordinal))?.CloneNode(true);
            bool hadList = originalList != null;
            return () => {
                PresentationExtensionList? list = root.GetFirstChild<PresentationExtensionList>();
                PresentationExtension? current = list?.Elements<PresentationExtension>()
                    .FirstOrDefault(extension => string.Equals(extension.Uri?.Value, SectionListUri, StringComparison.Ordinal));
                if (original != null) {
                    if (list == null) {
                        list = new PresentationExtensionList();
                        root.Append(list);
                    }
                    var restored = (PresentationExtension)original.CloneNode(true);
                    if (current == null) list.Append(restored);
                    else list.ReplaceChild(restored, current);
                } else {
                    current?.Remove();
                    if (!hadList && list != null && !list.HasChildren) list.Remove();
                }
            };
        }
    }
}
