using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Adds a basic clustered column chart with default data.
        /// </summary>
        public PowerPointChart AddChart() {
            return AddChartInternal(PowerPointChartKind.ClusteredColumn, null, 0L, 0L, 5486400L, 3200400L);
        }

        /// <summary>
        ///     Adds a basic clustered column chart with default data at a specific position.
        /// </summary>
        public PowerPointChart AddChart(long left, long top, long width, long height) {
            return AddChartInternal(PowerPointChartKind.ClusteredColumn, null, left, top, width, height);
        }

        /// <summary>
        ///     Adds a basic clustered column chart with default data using centimeter measurements.
        /// </summary>
        public PowerPointChart AddChartCm(double leftCm, double topCm, double widthCm, double heightCm) {
            return AddChart(
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a basic clustered column chart with default data using inch measurements.
        /// </summary>
        public PowerPointChart AddChartInches(double leftInches, double topInches, double widthInches,
            double heightInches) {
            return AddChart(
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a basic clustered column chart with default data using point measurements.
        /// </summary>
        public PowerPointChart AddChartPoints(double leftPoints, double topPoints, double widthPoints,
            double heightPoints) {
            return AddChart(
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a clustered column chart using the supplied data.
        /// </summary>
        public PowerPointChart AddChart(PowerPointChartData data, long left = 0L, long top = 0L, long width = 5486400L,
            long height = 3200400L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            return AddChartInternal(PowerPointChartKind.ClusteredColumn, data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a clustered column chart using the supplied data with centimeter measurements.
        /// </summary>
        public PowerPointChart AddChartCm(PowerPointChartData data, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddChart(data,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a clustered column chart using the supplied data with inch measurements.
        /// </summary>
        public PowerPointChart AddChartInches(PowerPointChartData data, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddChart(data,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a clustered column chart using the supplied data with point measurements.
        /// </summary>
        public PowerPointChart AddChartPoints(PowerPointChartData data, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddChart(data,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a clustered column chart using object data selectors.      
        /// </summary>
        public PowerPointChart AddChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddChart(items, categorySelector, 0L, 0L, 5486400L, 3200400L, seriesDefinitions);
        }

        /// <summary>
        ///     Adds a clustered column chart using object data selectors (centimeters).
        /// </summary>
        public PowerPointChart AddChartCm<T>(IEnumerable<T> items, Func<T, string> categorySelector, double leftCm,
            double topCm, double widthCm, double heightCm, params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddChart(items, categorySelector,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a clustered column chart using object data selectors (inches).
        /// </summary>
        public PowerPointChart AddChartInches<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftInches, double topInches, double widthInches, double heightInches,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddChart(items, categorySelector,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a clustered column chart using object data selectors (points).
        /// </summary>
        public PowerPointChart AddChartPoints<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddChart(items, categorySelector,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a clustered column chart using object data selectors at a specific position.
        /// </summary>
        public PowerPointChart AddChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            long left, long top, long width, long height,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            var data = PowerPointChartData.From(items, categorySelector, seriesDefinitions);
            return AddChartInternal(PowerPointChartKind.ClusteredColumn, data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a line chart with default data.
        /// </summary>
        public PowerPointChart AddLineChart() {
            return AddChartInternal(PowerPointChartKind.Line, null, 0L, 0L, 5486400L, 3200400L);
        }

        /// <summary>
        ///     Adds a line chart with default data at a specific position.
        /// </summary>
        public PowerPointChart AddLineChart(long left, long top, long width, long height) {
            return AddChartInternal(PowerPointChartKind.Line, null, left, top, width, height);
        }

        /// <summary>
        ///     Adds a line chart with default data using centimeter measurements.
        /// </summary>
        public PowerPointChart AddLineChartCm(double leftCm, double topCm, double widthCm, double heightCm) {
            return AddLineChart(
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a line chart with default data using inch measurements.
        /// </summary>
        public PowerPointChart AddLineChartInches(double leftInches, double topInches, double widthInches,
            double heightInches) {
            return AddLineChart(
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a line chart with default data using point measurements.
        /// </summary>
        public PowerPointChart AddLineChartPoints(double leftPoints, double topPoints, double widthPoints,
            double heightPoints) {
            return AddLineChart(
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a line chart using the supplied data.
        /// </summary>
        public PowerPointChart AddLineChart(PowerPointChartData data, long left = 0L, long top = 0L, long width = 5486400L,
            long height = 3200400L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            return AddChartInternal(PowerPointChartKind.Line, data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a line chart using the supplied data with centimeter measurements.
        /// </summary>
        public PowerPointChart AddLineChartCm(PowerPointChartData data, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddLineChart(data,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a line chart using the supplied data with inch measurements.
        /// </summary>
        public PowerPointChart AddLineChartInches(PowerPointChartData data, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddLineChart(data,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a line chart using the supplied data with point measurements.
        /// </summary>
        public PowerPointChart AddLineChartPoints(PowerPointChartData data, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddLineChart(data,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a line chart using object data selectors.
        /// </summary>
        public PowerPointChart AddLineChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddLineChart(items, categorySelector, 0L, 0L, 5486400L, 3200400L, seriesDefinitions);
        }

        /// <summary>
        ///     Adds a line chart using object data selectors (centimeters).
        /// </summary>
        public PowerPointChart AddLineChartCm<T>(IEnumerable<T> items, Func<T, string> categorySelector, double leftCm,
            double topCm, double widthCm, double heightCm, params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddLineChart(items, categorySelector,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a line chart using object data selectors (inches).
        /// </summary>
        public PowerPointChart AddLineChartInches<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftInches, double topInches, double widthInches, double heightInches,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddLineChart(items, categorySelector,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a line chart using object data selectors (points).
        /// </summary>
        public PowerPointChart AddLineChartPoints<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddLineChart(items, categorySelector,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a line chart using object data selectors at a specific position.
        /// </summary>
        public PowerPointChart AddLineChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            long left, long top, long width, long height,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            var data = PowerPointChartData.From(items, categorySelector, seriesDefinitions);
            return AddChartInternal(PowerPointChartKind.Line, data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a scatter chart with default data.
        /// </summary>
        public PowerPointChart AddScatterChart() {
            return AddScatterChartInternal(null, 0L, 0L, 5486400L, 3200400L);
        }

        /// <summary>
        ///     Adds a scatter chart with default data at a specific position.
        /// </summary>
        public PowerPointChart AddScatterChart(long left, long top, long width, long height) {
            return AddScatterChartInternal(null, left, top, width, height);
        }

        /// <summary>
        ///     Adds a scatter chart with default data using centimeter measurements.
        /// </summary>
        public PowerPointChart AddScatterChartCm(double leftCm, double topCm, double widthCm, double heightCm) {
            return AddScatterChart(
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a scatter chart with default data using inch measurements.
        /// </summary>
        public PowerPointChart AddScatterChartInches(double leftInches, double topInches, double widthInches,
            double heightInches) {
            return AddScatterChart(
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a scatter chart with default data using point measurements.
        /// </summary>
        public PowerPointChart AddScatterChartPoints(double leftPoints, double topPoints, double widthPoints,
            double heightPoints) {
            return AddScatterChart(
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a scatter chart using the supplied data.
        /// </summary>
        public PowerPointChart AddScatterChart(PowerPointScatterChartData data, long left = 0L, long top = 0L,
            long width = 5486400L, long height = 3200400L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            return AddScatterChartInternal(data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a scatter chart using the supplied data with centimeter measurements.
        /// </summary>
        public PowerPointChart AddScatterChartCm(PowerPointScatterChartData data, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddScatterChart(data,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a scatter chart using the supplied data with inch measurements.
        /// </summary>
        public PowerPointChart AddScatterChartInches(PowerPointScatterChartData data, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddScatterChart(data,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a scatter chart using the supplied data with point measurements.
        /// </summary>
        public PowerPointChart AddScatterChartPoints(PowerPointScatterChartData data, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddScatterChart(data,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a scatter chart using object data selectors.
        /// </summary>
        public PowerPointChart AddScatterChart<T>(IEnumerable<T> items, Func<T, double> xSelector,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddScatterChart(items, xSelector, 0L, 0L, 5486400L, 3200400L, seriesDefinitions);
        }

        /// <summary>
        ///     Adds a scatter chart using object data selectors (centimeters).
        /// </summary>
        public PowerPointChart AddScatterChartCm<T>(IEnumerable<T> items, Func<T, double> xSelector, double leftCm,
            double topCm, double widthCm, double heightCm, params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddScatterChart(items, xSelector,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a scatter chart using object data selectors (inches).
        /// </summary>
        public PowerPointChart AddScatterChartInches<T>(IEnumerable<T> items, Func<T, double> xSelector,
            double leftInches, double topInches, double widthInches, double heightInches,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddScatterChart(items, xSelector,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a scatter chart using object data selectors (points).
        /// </summary>
        public PowerPointChart AddScatterChartPoints<T>(IEnumerable<T> items, Func<T, double> xSelector,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddScatterChart(items, xSelector,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a scatter chart using object data selectors at a specific position.
        /// </summary>
        public PowerPointChart AddScatterChart<T>(IEnumerable<T> items, Func<T, double> xSelector,
            long left, long top, long width, long height,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            var data = PowerPointScatterChartData.From(items, xSelector, seriesDefinitions);
            return AddScatterChartInternal(data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a pie chart with default data.
        /// </summary>
        public PowerPointChart AddPieChart() {
            return AddChartInternal(PowerPointChartKind.Pie, null, 0L, 0L, 5486400L, 3200400L);
        }

        /// <summary>
        ///     Adds a pie chart with default data at a specific position.
        /// </summary>
        public PowerPointChart AddPieChart(long left, long top, long width, long height) {
            return AddChartInternal(PowerPointChartKind.Pie, null, left, top, width, height);
        }

        /// <summary>
        ///     Adds a pie chart with default data using centimeter measurements.
        /// </summary>
        public PowerPointChart AddPieChartCm(double leftCm, double topCm, double widthCm, double heightCm) {
            return AddPieChart(
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a pie chart with default data using inch measurements.
        /// </summary>
        public PowerPointChart AddPieChartInches(double leftInches, double topInches, double widthInches,
            double heightInches) {
            return AddPieChart(
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a pie chart with default data using point measurements.
        /// </summary>
        public PowerPointChart AddPieChartPoints(double leftPoints, double topPoints, double widthPoints,
            double heightPoints) {
            return AddPieChart(
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a pie chart using the supplied data.
        /// </summary>
        public PowerPointChart AddPieChart(PowerPointChartData data, long left = 0L, long top = 0L, long width = 5486400L,
            long height = 3200400L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            return AddChartInternal(PowerPointChartKind.Pie, data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a pie chart using the supplied data with centimeter measurements.
        /// </summary>
        public PowerPointChart AddPieChartCm(PowerPointChartData data, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddPieChart(data,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a pie chart using the supplied data with inch measurements.
        /// </summary>
        public PowerPointChart AddPieChartInches(PowerPointChartData data, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddPieChart(data,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a pie chart using the supplied data with point measurements.
        /// </summary>
        public PowerPointChart AddPieChartPoints(PowerPointChartData data, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddPieChart(data,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a pie chart using object data selectors.
        /// </summary>
        public PowerPointChart AddPieChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddPieChart(items, categorySelector, 0L, 0L, 5486400L, 3200400L, seriesDefinitions);
        }

        /// <summary>
        ///     Adds a pie chart using object data selectors (centimeters).
        /// </summary>
        public PowerPointChart AddPieChartCm<T>(IEnumerable<T> items, Func<T, string> categorySelector, double leftCm,
            double topCm, double widthCm, double heightCm, params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddPieChart(items, categorySelector,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a pie chart using object data selectors (inches).
        /// </summary>
        public PowerPointChart AddPieChartInches<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftInches, double topInches, double widthInches, double heightInches,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddPieChart(items, categorySelector,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a pie chart using object data selectors (points).
        /// </summary>
        public PowerPointChart AddPieChartPoints<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddPieChart(items, categorySelector,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a pie chart using object data selectors at a specific position.
        /// </summary>
        public PowerPointChart AddPieChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            long left, long top, long width, long height,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            var data = PowerPointChartData.From(items, categorySelector, seriesDefinitions);
            return AddChartInternal(PowerPointChartKind.Pie, data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a doughnut chart with default data.
        /// </summary>
        public PowerPointChart AddDoughnutChart() {
            return AddChartInternal(PowerPointChartKind.Doughnut, null, 0L, 0L, 5486400L, 3200400L);
        }

        /// <summary>
        ///     Adds a doughnut chart with default data at a specific position.
        /// </summary>
        public PowerPointChart AddDoughnutChart(long left, long top, long width, long height) {
            return AddChartInternal(PowerPointChartKind.Doughnut, null, left, top, width, height);
        }

        /// <summary>
        ///     Adds a doughnut chart with default data using centimeter measurements.
        /// </summary>
        public PowerPointChart AddDoughnutChartCm(double leftCm, double topCm, double widthCm, double heightCm) {
            return AddDoughnutChart(
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a doughnut chart with default data using inch measurements.
        /// </summary>
        public PowerPointChart AddDoughnutChartInches(double leftInches, double topInches, double widthInches,
            double heightInches) {
            return AddDoughnutChart(
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a doughnut chart with default data using point measurements.
        /// </summary>
        public PowerPointChart AddDoughnutChartPoints(double leftPoints, double topPoints, double widthPoints,
            double heightPoints) {
            return AddDoughnutChart(
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a doughnut chart using the supplied data.
        /// </summary>
        public PowerPointChart AddDoughnutChart(PowerPointChartData data, long left = 0L, long top = 0L, long width = 5486400L,
            long height = 3200400L) {
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            return AddChartInternal(PowerPointChartKind.Doughnut, data, left, top, width, height);
        }

        /// <summary>
        ///     Adds a doughnut chart using the supplied data with centimeter measurements.
        /// </summary>
        public PowerPointChart AddDoughnutChartCm(PowerPointChartData data, double leftCm, double topCm, double widthCm,
            double heightCm) {
            return AddDoughnutChart(data,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm));
        }

        /// <summary>
        ///     Adds a doughnut chart using the supplied data with inch measurements.
        /// </summary>
        public PowerPointChart AddDoughnutChartInches(PowerPointChartData data, double leftInches, double topInches,
            double widthInches, double heightInches) {
            return AddDoughnutChart(data,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches));
        }

        /// <summary>
        ///     Adds a doughnut chart using the supplied data with point measurements.
        /// </summary>
        public PowerPointChart AddDoughnutChartPoints(PowerPointChartData data, double leftPoints, double topPoints,
            double widthPoints, double heightPoints) {
            return AddDoughnutChart(data,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints));
        }

        /// <summary>
        ///     Adds a doughnut chart using object data selectors.
        /// </summary>
        public PowerPointChart AddDoughnutChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddDoughnutChart(items, categorySelector, 0L, 0L, 5486400L, 3200400L, seriesDefinitions);
        }

        /// <summary>
        ///     Adds a doughnut chart using object data selectors (centimeters).
        /// </summary>
        public PowerPointChart AddDoughnutChartCm<T>(IEnumerable<T> items, Func<T, string> categorySelector, double leftCm,
            double topCm, double widthCm, double heightCm, params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddDoughnutChart(items, categorySelector,
                PowerPointUnits.FromCentimeters(leftCm),
                PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm),
                PowerPointUnits.FromCentimeters(heightCm),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a doughnut chart using object data selectors (inches).
        /// </summary>
        public PowerPointChart AddDoughnutChartInches<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftInches, double topInches, double widthInches, double heightInches,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddDoughnutChart(items, categorySelector,
                PowerPointUnits.FromInches(leftInches),
                PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches),
                PowerPointUnits.FromInches(heightInches),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a doughnut chart using object data selectors (points).
        /// </summary>
        public PowerPointChart AddDoughnutChartPoints<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            return AddDoughnutChart(items, categorySelector,
                PowerPointUnits.FromPoints(leftPoints),
                PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints),
                PowerPointUnits.FromPoints(heightPoints),
                seriesDefinitions);
        }

        /// <summary>
        ///     Adds a doughnut chart using object data selectors at a specific position.
        /// </summary>
        public PowerPointChart AddDoughnutChart<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            long left, long top, long width, long height,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            var data = PowerPointChartData.From(items, categorySelector, seriesDefinitions);
            return AddChartInternal(PowerPointChartKind.Doughnut, data, left, top, width, height);
        }

        private PowerPointChart AddChartInternal(PowerPointChartKind chartKind, PowerPointChartData? data, long left, long top, long width, long height) {
            PowerPointChartData chartData = data ?? PowerPointChartData.CreateDefault();
            byte[] workbookBytes = PowerPointUtils.BuildChartWorkbook(chartData);
            return AddChartInternal(workbookBytes,
                (chartPart, embeddedRelId) => PowerPointUtils.PopulateChart(chartPart, embeddedRelId, chartData, chartKind),
                left, top, width, height);
        }

        private PowerPointChart AddScatterChartInternal(PowerPointScatterChartData? data, long left, long top, long width, long height) {
            PowerPointScatterChartData chartData = data ?? PowerPointScatterChartData.CreateDefault();
            byte[] workbookBytes = PowerPointUtils.BuildChartWorkbook(chartData);
            return AddChartInternal(workbookBytes,
                (chartPart, embeddedRelId) => PowerPointUtils.PopulateChart(chartPart, embeddedRelId, chartData, PowerPointChartKind.Scatter),
                left, top, width, height);
        }

        private PowerPointChart AddChartInternal(byte[] workbookBytes, Action<ChartPart, string> populateChart,
            long left, long top, long width, long height) {
            string chartPartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/charts",
                "chart",
                ".xml",
                allowBaseWithoutIndex: false);
            ChartPart chartPart = PowerPointPartFactory.CreatePart<ChartPart>(
                _slidePart,
                contentType: null,
                chartPartUri);
            string chartRelId = _slidePart.GetIdOfPart(chartPart);

            // Embed workbook + styles/colors exactly like the template
            string stylePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/charts",
                "style",
                ".xml",
                allowBaseWithoutIndex: false);
            ChartStylePart stylePart = PowerPointPartFactory.CreatePart<ChartStylePart>(
                chartPart,
                contentType: null,
                stylePartUri);
            PowerPointUtils.PopulateChartStyle(stylePart);
            string colorStylePartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/charts",
                "colors",
                ".xml",
                allowBaseWithoutIndex: false);
            ChartColorStylePart colorStylePart = PowerPointPartFactory.CreatePart<ChartColorStylePart>(
                chartPart,
                contentType: null,
                colorStylePartUri);
            PowerPointUtils.PopulateChartColorStyle(colorStylePart);

            string embeddedPartUri = PowerPointPartFactory.GetIndexedPartUri(
                _slidePart.OpenXmlPackage,
                "ppt/embeddings",
                "Microsoft_Excel_Worksheet",
                ".xlsx",
                allowBaseWithoutIndex: false);
            EmbeddedPackagePart embedded = PowerPointPartFactory.CreatePart<EmbeddedPackagePart>(
                chartPart,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                embeddedPartUri);
            using (var ms = new MemoryStream(workbookBytes)) {
                embedded.FeedData(ms);
            }

            string embeddedRelId = chartPart.GetIdOfPart(embedded);
            populateChart(chartPart, embeddedRelId);

            string name = GenerateUniqueName("Chart");
            GraphicFrame frame = new(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = _nextShapeId++, Name = name },
                    new NonVisualGraphicFrameDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new Transform(new A.Offset { X = left, Y = top }, new A.Extents { Cx = width, Cy = height }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = chartRelId }) {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                })
            );

            CommonSlideData dataElement = _slidePart.Slide.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = dataElement.ShapeTree ??= new ShapeTree();
            tree.AppendChild(frame);
            PowerPointChart chart = new(frame, _slidePart);
            _shapes.Add(chart);
            return chart;
        }

    }
}
