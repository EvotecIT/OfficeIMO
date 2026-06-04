using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    public partial class VisioPage {

        /// <summary>
        /// Adds a page layer, or returns the existing layer with the same name or universal name.
        /// </summary>
        /// <param name="name">Layer display name.</param>
        /// <param name="nameU">Optional universal name.</param>
        public VisioLayer AddLayer(string name, string? nameU = null) {
            VisioLayer? existing = FindLayer(nameU ?? name);
            if (existing != null) {
                return existing;
            }

            VisioLayer layer = new(name, nameU);
            _layers.Add(layer);
            return layer;
        }

        /// <summary>
        /// Finds a layer by display name or universal name.
        /// </summary>
        /// <param name="nameOrNameU">Layer name to find.</param>
        public VisioLayer? FindLayer(string nameOrNameU) {
            if (string.IsNullOrWhiteSpace(nameOrNameU)) {
                throw new ArgumentException("Layer name cannot be null or whitespace.", nameof(nameOrNameU));
            }

            foreach (VisioLayer layer in _layers) {
                if (string.Equals(layer.Name, nameOrNameU, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(layer.NameU, nameOrNameU, StringComparison.OrdinalIgnoreCase)) {
                    return layer;
                }
            }

            return null;
        }

        /// <summary>
        /// Adds a shape to a page layer, creating the layer if needed.
        /// </summary>
        /// <param name="layerName">Layer name.</param>
        /// <param name="shape">Shape to assign.</param>
        public VisioPage AddToLayer(string layerName, VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            EnsureShapeBelongsToPage(shape, "The shape must belong to the page.");
            AddLayer(layerName);
            shape.LayerNames.Add(layerName);
            return this;
        }

        /// <summary>
        /// Adds a connector to a page layer, creating the layer if needed.
        /// </summary>
        /// <param name="layerName">Layer name.</param>
        /// <param name="connector">Connector to assign.</param>
        public VisioPage AddToLayer(string layerName, VisioConnector connector) {
            if (connector == null) {
                throw new ArgumentNullException(nameof(connector));
            }

            EnsureConnectorBelongsToPage(connector);
            AddLayer(layerName);
            connector.LayerNames.Add(layerName);
            return this;
        }

        /// <summary>
        /// Adds a Visio-native container shape around existing member shapes.
        /// </summary>
        /// <param name="id">Container shape identifier.</param>
        /// <param name="text">Container heading text.</param>
        /// <param name="members">Shapes that should belong to the container.</param>
        /// <param name="options">Optional container layout and style settings.</param>
        /// <returns>The created container shape.</returns>
        public VisioShape AddContainer(string id, string text, IEnumerable<VisioShape> members, VisioContainerOptions? options = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Container id cannot be empty.", nameof(id));
            }

            if (members == null) {
                throw new ArgumentNullException(nameof(members));
            }

            List<VisioShape> memberList = members.Distinct().ToList();
            if (memberList.Count == 0) {
                throw new ArgumentException("A container requires at least one member shape.", nameof(members));
            }

            foreach (VisioShape member in memberList) {
                if (!ContainsShape(member)) {
                    throw new InvalidOperationException("All container members must already belong to this page.");
                }
            }

            VisioContainerOptions effectiveOptions = options ?? new VisioContainerOptions();
            GetContainerBounds(memberList, effectiveOptions, DefaultUnit, out double pinX, out double pinY, out double width, out double height);
            VisioShape container = AddContainer(id, pinX, pinY, width, height, text, effectiveOptions);

            foreach (VisioShape member in memberList) {
                if (!container.ContainerMemberIds.Contains(member.Id, StringComparer.OrdinalIgnoreCase)) {
                    container.ContainerMemberIds.Add(member.Id);
                }

                if (!member.ContainerOwnerIds.Contains(container.Id, StringComparer.OrdinalIgnoreCase)) {
                    member.ContainerOwnerIds.Add(container.Id);
                }
            }

            return container;
        }

        /// <summary>
        /// Adds a Visio-native container shape at an explicit location and size.
        /// </summary>
        /// <param name="id">Container shape identifier.</param>
        /// <param name="pinX">Container pin X coordinate.</param>
        /// <param name="pinY">Container pin Y coordinate.</param>
        /// <param name="width">Container width.</param>
        /// <param name="height">Container height.</param>
        /// <param name="text">Container heading text.</param>
        /// <param name="options">Optional style and semantic settings.</param>
        /// <returns>The created container shape.</returns>
        public VisioShape AddContainer(string id, double pinX, double pinY, double width, double height, string? text = null, VisioContainerOptions? options = null) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException("Container id cannot be empty.", nameof(id));
            }

            VisioContainerOptions effectiveOptions = options ?? new VisioContainerOptions();
            VisioShape container = new(id, pinX, pinY, width, height, text ?? string.Empty) {
                Name = "Container",
                NameU = "Container"
            };
            ApplyContainerSemantics(container, effectiveOptions, DefaultUnit);
            Shapes.Insert(0, container);
            return container;
        }

        private static void ApplyContainerSemantics(VisioShape container, VisioContainerOptions options, VisioMeasurementUnit unit) {
            VisioContainerSemantics.Apply(container, options, unit);
        }

        private static void GetContainerBounds(IReadOnlyList<VisioShape> members, VisioContainerOptions options, VisioMeasurementUnit unit, out double pinX, out double pinY, out double width, out double height) {
            double left = double.PositiveInfinity;
            double bottom = double.PositiveInfinity;
            double right = double.NegativeInfinity;
            double top = double.NegativeInfinity;

            foreach (VisioShape member in members) {
                (double memberLeft, double memberBottom, double memberRight, double memberTop) = member.GetBounds();
                left = Math.Min(left, memberLeft);
                bottom = Math.Min(bottom, memberBottom);
                right = Math.Max(right, memberRight);
                top = Math.Max(top, memberTop);
            }

            double margin = options.Margin.ToInches(unit);
            double headingHeight = options.HeadingHeight.ToInches(unit);
            left -= margin;
            right += margin;
            bottom -= margin;
            top += margin + headingHeight;
            width = Math.Max(0.1D, right - left);
            height = Math.Max(0.1D, top - bottom);
            pinX = left + (width / 2D);
            pinY = bottom + (height / 2D);
        }

        private bool ContainsShape(VisioShape target) {
            foreach (VisioShape shape in _shapes) {
                if (ContainsShape(shape, target)) {
                    return true;
                }
            }

            return false;
        }

        private static bool ContainsShape(VisioShape current, VisioShape target) {
            if (ReferenceEquals(current, target)) {
                return true;
            }

            foreach (VisioShape child in current.Children) {
                if (ContainsShape(child, target)) {
                    return true;
                }
            }

            return false;
        }
    }
}
