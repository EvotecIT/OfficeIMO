using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Editing helpers for maintaining Visio-native container membership and bounds.
    /// </summary>
    public static class VisioContainerEditingExtensions {
        /// <summary>
        /// Gets typed metadata and membership information for a Visio-native container.
        /// </summary>
        /// <param name="page">Page that owns the container.</param>
        /// <param name="container">Container shape.</param>
        public static VisioContainerInfo GetContainerInfo(this VisioPage page, VisioShape container) {
            ValidatePageAndContainer(page, container);
            VisioContainerOptions options = VisioContainerSemantics.CreateOptionsFrom(container, page.DefaultUnit);
            return new VisioContainerInfo(
                container.Id,
                container.Text,
                container.ContainerMemberIds.Where(id => !string.IsNullOrWhiteSpace(id)).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                options.Margin,
                options.HeadingHeight,
                options.AutoResize,
                options.Locked,
                options.NoHighlight,
                options.NoRibbon,
                options.ContainerStyle,
                options.HeadingStyle,
                container.FillColor,
                container.LineColor,
                container.LineWeight,
                container.FillPattern,
                container.LinePattern,
                container.TextStyle);
        }

        /// <summary>
        /// Gets editable options initialized from a Visio-native container's current metadata and visual style.
        /// </summary>
        /// <param name="page">Page that owns the container.</param>
        /// <param name="container">Container shape.</param>
        public static VisioContainerOptions GetContainerOptions(this VisioPage page, VisioShape container) {
            ValidatePageAndContainer(page, container);
            return VisioContainerSemantics.CreateOptionsFrom(container, page.DefaultUnit);
        }

        /// <summary>
        /// Applies native container metadata and visual style to an existing Visio-native container.
        /// </summary>
        /// <param name="page">Page that owns the container.</param>
        /// <param name="container">Container shape to update.</param>
        /// <param name="options">Container metadata and style options.</param>
        /// <param name="refit">Whether the container should be refit around current members after metadata is applied.</param>
        public static VisioShape ApplyContainerOptions(this VisioPage page, VisioShape container, VisioContainerOptions options, bool refit = false) {
            ValidatePageAndContainer(page, container);
            VisioContainerSemantics.Apply(container, options, page.DefaultUnit);
            ClearRelationshipFormula(container);

            if (refit) {
                page.RefitContainer(container, options);
            }

            return container;
        }

        /// <summary>
        /// Updates native container metadata and visual style using a callback initialized from the current container state.
        /// </summary>
        /// <param name="page">Page that owns the container.</param>
        /// <param name="container">Container shape to update.</param>
        /// <param name="configure">Option callback.</param>
        /// <param name="refit">Whether the container should be refit around current members after metadata is applied.</param>
        public static VisioShape ConfigureContainer(this VisioPage page, VisioShape container, Action<VisioContainerOptions> configure, bool refit = false) {
            if (configure == null) {
                throw new ArgumentNullException(nameof(configure));
            }

            VisioContainerOptions options = page.GetContainerOptions(container);
            configure(options);
            return page.ApplyContainerOptions(container, options, refit);
        }

        /// <summary>
        /// Gets the shapes currently referenced by a Visio-native container.
        /// </summary>
        /// <param name="page">Page that owns the container.</param>
        /// <param name="container">Container shape.</param>
        public static IReadOnlyList<VisioShape> GetContainerMembers(this VisioPage page, VisioShape container) {
            ValidatePageAndContainer(page, container);

            List<VisioShape> members = new();
            foreach (string memberId in container.ContainerMemberIds.Where(id => !string.IsNullOrWhiteSpace(id)).Distinct(StringComparer.OrdinalIgnoreCase)) {
                VisioShape? member = page.FindShapeById(memberId);
                if (member != null) {
                    members.Add(member);
                }
            }

            return members;
        }

        /// <summary>
        /// Adds one member shape to a Visio-native container and optionally resizes the container around all current members.
        /// </summary>
        public static VisioShape AddToContainer(this VisioPage page, VisioShape container, VisioShape member, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            if (member == null) {
                throw new ArgumentNullException(nameof(member));
            }

            return page.AddToContainer(container, new[] { member }, resizeToFit, resizeOptions);
        }

        /// <summary>
        /// Adds member shapes to a Visio-native container and optionally resizes the container around all current members.
        /// </summary>
        public static VisioShape AddToContainer(this VisioPage page, VisioShape container, IEnumerable<VisioShape> members, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            ValidatePageAndContainer(page, container);
            List<VisioShape> memberList = ValidateMembers(page, container, members);

            foreach (VisioShape member in memberList) {
                AddUnique(container.ContainerMemberIds, member.Id);
                AddUnique(member.ContainerOwnerIds, container.Id);
                ClearRelationshipFormula(member);
            }

            ClearRelationshipFormula(container);
            if (resizeToFit) {
                page.RefitContainer(container, resizeOptions);
            }

            return container;
        }

        /// <summary>
        /// Removes one member shape from a Visio-native container and optionally resizes the container around remaining members.
        /// </summary>
        public static VisioShape RemoveFromContainer(this VisioPage page, VisioShape container, VisioShape member, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            if (member == null) {
                throw new ArgumentNullException(nameof(member));
            }

            return page.RemoveFromContainer(container, new[] { member }, resizeToFit, resizeOptions);
        }

        /// <summary>
        /// Removes member shapes from a Visio-native container and optionally resizes the container around remaining members.
        /// </summary>
        public static VisioShape RemoveFromContainer(this VisioPage page, VisioShape container, IEnumerable<VisioShape> members, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            ValidatePageAndContainer(page, container);
            List<VisioShape> memberList = ValidateMembers(page, container, members);

            foreach (VisioShape member in memberList) {
                RemoveAll(container.ContainerMemberIds, member.Id);
                RemoveAll(member.ContainerOwnerIds, container.Id);
                ClearRelationshipFormula(member);
            }

            ClearRelationshipFormula(container);
            if (resizeToFit && container.ContainerMemberIds.Count > 0) {
                page.RefitContainer(container, resizeOptions);
            }

            return container;
        }

        /// <summary>
        /// Resizes a Visio-native container so it encloses its current typed members.
        /// </summary>
        /// <param name="page">Page that owns the container.</param>
        /// <param name="container">Container shape to resize.</param>
        /// <param name="options">Optional margin and heading settings used for the calculated bounds.</param>
        public static VisioShape RefitContainer(this VisioPage page, VisioShape container, VisioContainerOptions? options = null) {
            ValidatePageAndContainer(page, container);

            IReadOnlyList<VisioShape> members = page.GetContainerMembers(container);
            if (members.Count == 0) {
                return container;
            }

            VisioContainerOptions effectiveOptions = options ?? VisioContainerSemantics.CreateOptionsFrom(container, page.DefaultUnit);
            VisioContainerSemantics.Validate(effectiveOptions);
            GetContainerBounds(members, effectiveOptions, page.DefaultUnit, out double pinX, out double pinY, out double width, out double height);
            container.PinX = pinX;
            container.PinY = pinY;
            container.Width = width;
            container.Height = height;
            container.LocPinX = width / 2D;
            container.LocPinY = height / 2D;
            VisioContainerSemantics.Apply(container, effectiveOptions, page.DefaultUnit);
            return container;
        }

        /// <summary>
        /// Adds all selected shapes to an existing Visio-native container.
        /// </summary>
        public static VisioShapeSelection AddToContainer(this VisioShapeSelection selection, VisioShape container, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            VisioPage page = RequireOwnerPage(selection);
            page.AddToContainer(container, selection, resizeToFit, resizeOptions);
            return selection;
        }

        /// <summary>
        /// Removes all selected shapes from an existing Visio-native container.
        /// </summary>
        public static VisioShapeSelection RemoveFromContainer(this VisioShapeSelection selection, VisioShape container, bool resizeToFit = true, VisioContainerOptions? resizeOptions = null) {
            VisioPage page = RequireOwnerPage(selection);
            page.RemoveFromContainer(container, selection, resizeToFit, resizeOptions);
            return selection;
        }

        /// <summary>
        /// Creates a Visio-native container around the selected shapes.
        /// </summary>
        public static VisioShape WrapInContainer(this VisioShapeSelection selection, string id, string text, VisioContainerOptions? options = null) {
            VisioPage page = RequireOwnerPage(selection);
            return page.AddContainer(id, text, selection, options);
        }

        private static void ValidatePageAndContainer(VisioPage page, VisioShape container) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (container == null) {
                throw new ArgumentNullException(nameof(container));
            }

            if (!container.IsContainer) {
                throw new ArgumentException("The shape is not marked as a Visio-native container.", nameof(container));
            }

            if (!page.AllShapes().Contains(container)) {
                throw new InvalidOperationException("The container must belong to the page.");
            }
        }

        private static List<VisioShape> ValidateMembers(VisioPage page, VisioShape container, IEnumerable<VisioShape> members) {
            if (members == null) {
                throw new ArgumentNullException(nameof(members));
            }

            List<VisioShape> memberList = members.Where(member => member != null).Distinct().ToList();
            if (memberList.Count == 0) {
                throw new ArgumentException("At least one member shape is required.", nameof(members));
            }

            IReadOnlyList<VisioShape> pageShapes = page.AllShapes();
            foreach (VisioShape member in memberList) {
                if (ReferenceEquals(member, container)) {
                    throw new ArgumentException("A container cannot contain itself.", nameof(members));
                }

                if (!pageShapes.Contains(member)) {
                    throw new InvalidOperationException("All container members must belong to the page.");
                }
            }

            return memberList;
        }

        private static VisioPage RequireOwnerPage(VisioShapeSelection selection) {
            if (selection == null) {
                throw new ArgumentNullException(nameof(selection));
            }

            return selection.OwnerPage ?? throw new InvalidOperationException("The selection must be created from a VisioPage query before it can edit container membership.");
        }

        private static void GetContainerBounds(IReadOnlyList<VisioShape> members, VisioContainerOptions options, VisioMeasurementUnit unit, out double pinX, out double pinY, out double width, out double height) {
            VisioShapeBounds bounds = members.GetShapeBounds();
            double margin = options.Margin.ToInches(unit);
            double headingHeight = options.HeadingHeight.ToInches(unit);
            double left = bounds.Left - margin;
            double right = bounds.Right + margin;
            double bottom = bounds.Bottom - margin;
            double top = bounds.Top + margin + headingHeight;
            width = Math.Max(0.1D, right - left);
            height = Math.Max(0.1D, top - bottom);
            pinX = left + width / 2D;
            pinY = bottom + height / 2D;
        }

        private static void AddUnique(IList<string> values, string value) {
            if (!values.Contains(value, StringComparer.OrdinalIgnoreCase)) {
                values.Add(value);
            }
        }

        private static void RemoveAll(IList<string> values, string value) {
            for (int i = values.Count - 1; i >= 0; i--) {
                if (string.Equals(values[i], value, StringComparison.OrdinalIgnoreCase)) {
                    values.RemoveAt(i);
                }
            }
        }

        private static void ClearRelationshipFormula(VisioShape shape) {
            shape.RelationshipsFormula = null;
            shape.RelationshipsValue = null;
        }
    }
}
