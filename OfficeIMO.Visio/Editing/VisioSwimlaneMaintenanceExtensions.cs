using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Helpers for discovering and maintaining swimlane diagrams in new or loaded pages.
    /// </summary>
    public static class VisioSwimlaneMaintenanceExtensions {
        private const double BoundsTolerance = 0.000001D;

        private static readonly string[] ActivityStencilIds = {
            "swim.activity",
            "swim.decision",
            "swim.data",
            "swim.start-end"
        };

        /// <summary>
        /// Finds swimlane lanes on the page, using OfficeIMO semantic metadata first and generated IDs/stencil metadata as a fallback.
        /// </summary>
        public static IReadOnlyList<VisioSwimlaneLane> GetSwimlaneLanes(this VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            Dictionary<string, VisioShape> headers = new(StringComparer.Ordinal);
            foreach (VisioShape shape in page.Shapes) {
                if (IsLaneHeader(shape)) {
                    string? id = GetLaneId(shape, header: true);
                    if (!string.IsNullOrWhiteSpace(id) && !headers.ContainsKey(id!)) {
                        headers.Add(id!, shape);
                    }
                }
            }

            List<VisioSwimlaneLane> lanes = new();
            foreach (VisioShape body in page.Shapes.Where(IsLaneBody).OrderByDescending(shape => shape.PinY).ThenBy(shape => shape.PinX)) {
                string? id = GetLaneId(body, header: false);
                if (string.IsNullOrWhiteSpace(id)) {
                    continue;
                }

                headers.TryGetValue(id!, out VisioShape? header);
                lanes.Add(new VisioSwimlaneLane(id!, body, header));
            }

            return lanes;
        }

        /// <summary>
        /// Finds swimlane phase columns on the page, using OfficeIMO semantic metadata first and generated IDs/stencil metadata as a fallback.
        /// </summary>
        public static IReadOnlyList<VisioSwimlanePhase> GetSwimlanePhases(this VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            List<VisioSwimlanePhase> phases = new();
            foreach (VisioShape header in page.Shapes.Where(IsPhaseHeader).OrderBy(shape => shape.PinX)) {
                string? id = GetPhaseId(header);
                if (!string.IsNullOrWhiteSpace(id)) {
                    phases.Add(new VisioSwimlanePhase(id!, header));
                }
            }

            return phases;
        }

        /// <summary>
        /// Finds swimlane activities and their current lane/phase placement.
        /// </summary>
        public static IReadOnlyList<VisioSwimlaneActivityPlacement> GetSwimlaneActivities(this VisioPage page) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            IReadOnlyList<VisioSwimlaneLane> lanes = page.GetSwimlaneLanes();
            IReadOnlyList<VisioSwimlanePhase> phases = page.GetSwimlanePhases();
            List<VisioSwimlaneActivityPlacement> activities = new();
            foreach (VisioShape shape in page.Shapes.Where(IsActivityShape).OrderByDescending(shape => shape.PinY).ThenBy(shape => shape.PinX)) {
                string? laneId = GetPlacementLaneId(shape) ?? InferLaneId(shape, lanes);
                string? phaseId = GetPlacementPhaseId(shape) ?? InferPhaseId(shape, phases);
                activities.Add(new VisioSwimlaneActivityPlacement(shape, laneId, phaseId, GetActivityKind(shape)));
            }

            return activities;
        }

        /// <summary>
        /// Finds a swimlane lane by identifier.
        /// </summary>
        public static VisioSwimlaneLane? FindSwimlaneLane(this VisioPage page, string laneId) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(laneId)) {
                throw new ArgumentException("Lane id cannot be empty.", nameof(laneId));
            }

            return page.GetSwimlaneLanes().FirstOrDefault(lane => string.Equals(lane.Id, laneId, StringComparison.Ordinal));
        }

        /// <summary>
        /// Finds a swimlane phase by identifier.
        /// </summary>
        public static VisioSwimlanePhase? FindSwimlanePhase(this VisioPage page, string phaseId) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(phaseId)) {
                throw new ArgumentException("Phase id cannot be empty.", nameof(phaseId));
            }

            return page.GetSwimlanePhases().FirstOrDefault(phase => string.Equals(phase.Id, phaseId, StringComparison.Ordinal));
        }

        /// <summary>
        /// Moves a swimlane activity to a target lane/phase cell and relayouts the affected swimlane activities.
        /// </summary>
        public static VisioPage MoveSwimlaneActivity(this VisioPage page, string activityId, string laneId, string phaseId, VisioSwimlaneRelayoutOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            if (string.IsNullOrWhiteSpace(activityId)) {
                throw new ArgumentException("Activity id cannot be empty.", nameof(activityId));
            }

            if (string.IsNullOrWhiteSpace(laneId)) {
                throw new ArgumentException("Lane id cannot be empty.", nameof(laneId));
            }

            if (string.IsNullOrWhiteSpace(phaseId)) {
                throw new ArgumentException("Phase id cannot be empty.", nameof(phaseId));
            }

            VisioShape? activity = page.Shapes.FirstOrDefault(shape => string.Equals(shape.Id, activityId, StringComparison.Ordinal));
            if (activity == null || !IsActivityShape(activity)) {
                throw new ArgumentException("Swimlane activity '" + activityId + "' was not found.", nameof(activityId));
            }

            if (page.FindSwimlaneLane(laneId) == null) {
                throw new ArgumentException("Swimlane lane '" + laneId + "' was not found.", nameof(laneId));
            }

            if (page.FindSwimlanePhase(phaseId) == null) {
                throw new ArgumentException("Swimlane phase '" + phaseId + "' was not found.", nameof(phaseId));
            }

            MarkActivityPlacement(activity, laneId, phaseId, GetActivityKind(activity));
            return page.RelayoutSwimlaneActivities(options);
        }

        /// <summary>
        /// Re-centers swimlane activities inside their lane/phase cells, stacking multiple activities deterministically.
        /// </summary>
        public static VisioPage RelayoutSwimlaneActivities(this VisioPage page, VisioSwimlaneRelayoutOptions? options = null) {
            if (page == null) {
                throw new ArgumentNullException(nameof(page));
            }

            VisioSwimlaneRelayoutOptions resolvedOptions = options ?? new VisioSwimlaneRelayoutOptions();
            ValidateOptions(resolvedOptions);

            Dictionary<string, VisioSwimlaneLane> lanes = page.GetSwimlaneLanes().ToDictionary(lane => lane.Id, StringComparer.Ordinal);
            Dictionary<string, VisioSwimlanePhase> phases = page.GetSwimlanePhases().ToDictionary(phase => phase.Id, StringComparer.Ordinal);
            List<VisioSwimlaneActivityPlacement> activities = page.GetSwimlaneActivities()
                .Where(activity => activity.LaneId != null && activity.PhaseId != null && lanes.ContainsKey(activity.LaneId) && phases.ContainsKey(activity.PhaseId))
                .ToList();

            foreach (IGrouping<string, VisioSwimlaneActivityPlacement> group in activities.GroupBy(activity => activity.LaneId + "\u001f" + activity.PhaseId, StringComparer.Ordinal)) {
                List<VisioSwimlaneActivityPlacement> ordered = group
                    .OrderByDescending(activity => activity.Shape.PinY)
                    .ThenBy(activity => activity.Shape.PinX)
                    .ThenBy(activity => activity.Shape.Id, StringComparer.Ordinal)
                    .ToList();
                if (ordered.Count == 0) {
                    continue;
                }

                string laneId = ordered[0].LaneId!;
                string phaseId = ordered[0].PhaseId!;
                VisioShape laneBody = lanes[laneId].Body;
                VisioShape phaseHeader = phases[phaseId].Header;
                double totalHeight = ordered.Sum(activity => activity.Shape.Height) + Math.Max(0, ordered.Count - 1) * resolvedOptions.ActivityGap;
                double y = laneBody.PinY + (totalHeight / 2D);
                foreach (VisioSwimlaneActivityPlacement activity in ordered) {
                    VisioShape shape = activity.Shape;
                    y -= shape.Height / 2D;
                    shape.PinX = phaseHeader.PinX;
                    shape.PinY = y;
                    MarkActivityPlacement(shape, laneId, phaseId, activity.ActivityKind);
                    y -= shape.Height / 2D + resolvedOptions.ActivityGap;
                }
            }

            if (resolvedOptions.RerouteConnectors) {
                RerouteActivityConnectors(page, activities.Select(activity => activity.Shape), resolvedOptions);
            }

            return page;
        }

        internal static void MarkActivityPlacement(VisioShape shape, string laneId, string phaseId, VisioSwimlaneActivityKind? kind) {
            shape.SetUserCell(VisioSemanticUserCells.Kind, VisioSemanticUserCells.SwimlaneActivityKind, "STR", prompt: "OfficeIMO semantic kind");
            shape.SetUserCell(VisioSemanticUserCells.SwimlaneLaneId, laneId, "STR", prompt: "OfficeIMO swimlane lane id");
            shape.SetUserCell(VisioSemanticUserCells.SwimlanePhaseId, phaseId, "STR", prompt: "OfficeIMO swimlane phase id");
            if (kind.HasValue) {
                shape.SetUserCell(VisioSemanticUserCells.SwimlaneActivityType, kind.Value.ToString(), "STR", prompt: "OfficeIMO swimlane activity type");
            }
        }

        private static void RerouteActivityConnectors(VisioPage page, IEnumerable<VisioShape> activityShapes, VisioSwimlaneRelayoutOptions options) {
            HashSet<VisioShape> activities = new(activityShapes);
            int routeIndex = 0;
            foreach (VisioConnector connector in page.Connectors) {
                if (!activities.Contains(connector.From) && !activities.Contains(connector.To)) {
                    continue;
                }

                if (options.AvoidShapes) {
                    connector.RouteOrthogonalAroundShapes(page.Shapes, new VisioConnectorRoutingOptions {
                        Padding = options.RoutingPadding,
                        MaxLanes = options.MaxRoutingLanes
                    });
                } else {
                    connector.RouteOrthogonal(offset: (routeIndex % 4) * 0.04D);
                }

                routeIndex++;
            }
        }

        private static void ValidateOptions(VisioSwimlaneRelayoutOptions options) {
            if (double.IsNaN(options.ActivityGap) || double.IsInfinity(options.ActivityGap) || options.ActivityGap < 0D) {
                throw new ArgumentOutOfRangeException(nameof(options), "Activity gap must be a non-negative finite value.");
            }

            if (double.IsNaN(options.RoutingPadding) || double.IsInfinity(options.RoutingPadding) || options.RoutingPadding < 0D) {
                throw new ArgumentOutOfRangeException(nameof(options), "Routing padding must be a non-negative finite value.");
            }

            if (options.MaxRoutingLanes < 0) {
                throw new ArgumentOutOfRangeException(nameof(options), "Maximum routing lanes cannot be negative.");
            }
        }

        private static bool IsLaneHeader(VisioShape shape) {
            return string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.SwimlaneLaneHeaderKind, StringComparison.OrdinalIgnoreCase)
                || shape.Id.StartsWith("lane-header-", StringComparison.Ordinal);
        }

        private static bool IsLaneBody(VisioShape shape) {
            if (string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.SwimlaneLaneKind, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (shape.Id.StartsWith("lane-", StringComparison.Ordinal) && !shape.Id.StartsWith("lane-header-", StringComparison.Ordinal)) {
                return true;
            }

            return HasStencilId(shape, "swim.lane");
        }

        private static bool IsPhaseHeader(VisioShape shape) {
            return string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.SwimlanePhaseKind, StringComparison.OrdinalIgnoreCase)
                || shape.Id.StartsWith("phase-", StringComparison.Ordinal)
                || HasStencilId(shape, "swim.phase");
        }

        private static bool IsActivityShape(VisioShape shape) {
            if (string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.Kind), VisioSemanticUserCells.SwimlaneActivityKind, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            string? stencilId = shape.GetUserCellValue(VisioSemanticUserCells.StencilId);
            return stencilId != null && ActivityStencilIds.Any(id => string.Equals(id, stencilId, StringComparison.OrdinalIgnoreCase));
        }

        private static string? GetLaneId(VisioShape shape, bool header) {
            string? id = shape.GetUserCellValue(VisioSemanticUserCells.SwimlaneLaneId);
            if (!string.IsNullOrWhiteSpace(id)) {
                return id;
            }

            string prefix = header ? "lane-header-" : "lane-";
            if (shape.Id.StartsWith(prefix, StringComparison.Ordinal)) {
                return shape.Id.Substring(prefix.Length);
            }

            return shape.Id;
        }

        private static string? GetPhaseId(VisioShape shape) {
            string? id = shape.GetUserCellValue(VisioSemanticUserCells.SwimlanePhaseId);
            if (!string.IsNullOrWhiteSpace(id)) {
                return id;
            }

            const string prefix = "phase-";
            return shape.Id.StartsWith(prefix, StringComparison.Ordinal)
                ? shape.Id.Substring(prefix.Length)
                : shape.Id;
        }

        private static string? GetPlacementLaneId(VisioShape shape) {
            string? id = shape.GetUserCellValue(VisioSemanticUserCells.SwimlaneLaneId);
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        private static string? GetPlacementPhaseId(VisioShape shape) {
            string? id = shape.GetUserCellValue(VisioSemanticUserCells.SwimlanePhaseId);
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        private static string? InferLaneId(VisioShape activity, IReadOnlyList<VisioSwimlaneLane> lanes) {
            foreach (VisioSwimlaneLane lane in lanes) {
                VisioShapeBounds bounds = lane.Body.GetShapeBounds();
                if (activity.PinY >= bounds.Bottom - BoundsTolerance && activity.PinY <= bounds.Top + BoundsTolerance) {
                    return lane.Id;
                }
            }

            return null;
        }

        private static string? InferPhaseId(VisioShape activity, IReadOnlyList<VisioSwimlanePhase> phases) {
            foreach (VisioSwimlanePhase phase in phases) {
                VisioShapeBounds bounds = phase.Header.GetShapeBounds();
                if (activity.PinX >= bounds.Left - BoundsTolerance && activity.PinX <= bounds.Right + BoundsTolerance) {
                    return phase.Id;
                }
            }

            return null;
        }

        private static VisioSwimlaneActivityKind? GetActivityKind(VisioShape shape) {
            string? value = shape.GetUserCellValue(VisioSemanticUserCells.SwimlaneActivityType);
            if (!string.IsNullOrWhiteSpace(value) && Enum.TryParse(value, ignoreCase: true, out VisioSwimlaneActivityKind kind)) {
                return kind;
            }

            string? stencilId = shape.GetUserCellValue(VisioSemanticUserCells.StencilId);
            if (string.Equals(stencilId, "swim.activity", StringComparison.OrdinalIgnoreCase)) {
                return VisioSwimlaneActivityKind.Step;
            }

            if (string.Equals(stencilId, "swim.decision", StringComparison.OrdinalIgnoreCase)) {
                return VisioSwimlaneActivityKind.Decision;
            }

            if (string.Equals(stencilId, "swim.data", StringComparison.OrdinalIgnoreCase)) {
                return VisioSwimlaneActivityKind.Data;
            }

            return null;
        }

        private static bool HasStencilId(VisioShape shape, string stencilId) {
            return string.Equals(shape.GetUserCellValue(VisioSemanticUserCells.StencilId), stencilId, StringComparison.OrdinalIgnoreCase);
        }
    }
}
