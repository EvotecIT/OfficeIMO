using System;
using System.Collections.Generic;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Verifies accessibility and defaults of Excel-related enums.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void TableStyleHasValues() {
            Assert.True(Enum.IsDefined(typeof(TableStyle), nameof(TableStyle.TableStyleLight1)));
        }

        [Fact]
        public void ExecutionPolicyDefaultsToAutomatic() {
            var policy = new ExecutionPolicy();
            Assert.Equal(ExecutionMode.Automatic, policy.Mode);
        }

        [Fact]
        public void ObjectFlattenerOptionsDefaults() {
            var opts = new ObjectFlattenerOptions();
            Assert.Equal(HeaderCase.Raw, opts.HeaderCase);
            Assert.Equal(NullPolicy.NullLiteral, opts.NullPolicy);
            Assert.Equal(CollectionMode.JoinWith, opts.CollectionMode);
        }

        [Fact]
        public void ObjectFlattenerApplyOrderingPreservesPinsPrioritiesAndDiscoveryOrder() {
            var input = new List<string> {
                "Id",
                "Details.Score",
                "Name",
                "Details.Status",
                "Created",
                "Notes"
            };
            var opts = new ObjectFlattenerOptions()
                .PinFirst("Name")
                .PriorityOrder("Status", "Score")
                .PinLast("Notes");

            var ordered = ObjectFlattener.ApplyOrdering(input, opts);

            Assert.Equal(new[] {
                "Name",
                "Id",
                "Created",
                "Details.Status",
                "Details.Score",
                "Notes"
            }, ordered);
        }

        [Fact]
        public void ObjectFlattenerApplySelectionCombinesIgnoreExcludeAndInclude() {
            var input = new List<string> {
                "Id",
                "Details.Score",
                "Details.Status",
                "Details.Secret",
                "Ignored.Value",
                "Notes"
            };
            var opts = new ObjectFlattenerOptions {
                Ignore = new[] { "Ignored" },
                IncludeProperties = new[] { "Id", "Score", "Status", "Secret", "Value" },
                ExcludeProperties = new[] { "Secret" }
            };

            var selected = ObjectFlattener.ApplySelection(input, opts);

            Assert.Equal(new[] {
                "Id",
                "Details.Score",
                "Details.Status"
            }, selected);
        }

        [Fact]
        public void ObjectFlattenerGetPathsAppliesSelectionAndOrdering() {
            var flattener = new ObjectFlattener();
            var opts = new ObjectFlattenerOptions {
                IncludeProperties = new[] { "Id", "Score", "Status", "Secret", "Value" },
                ExcludeProperties = new[] { "Secret" },
                Ignore = new[] { "Ignored" }
            }.PinFirst("Status").PinLast("Id");
            opts.ExpandProperties.Add(nameof(ObjectFlattenerSelectionPathRow.Details));

            var paths = flattener.GetPaths(typeof(ObjectFlattenerSelectionPathRow), opts);

            Assert.Equal(new[] {
                "Details.Status",
                "Details.Score",
                "Id"
            }, paths);
        }

        [Fact]
        public void ObjectFlattenerJoinCollectionsPreservesNullAndEmptyItems() {
            var flattener = new ObjectFlattener();
            var values = flattener.Flatten(new ObjectFlattenerCollectionRow(), new ObjectFlattenerOptions());

            Assert.Equal("a,,b", values["Tags"]);
            Assert.Equal(string.Empty, values["Empty"]);
        }

        [Fact]
        public void ObjectFlattenerCollectionMapColumnsPreservesDynamicColumns() {
            var flattener = new ObjectFlattener();
            var options = new ObjectFlattenerOptions();
            options.CollectionMapColumns["Metrics"] = new CollectionColumnMapping {
                KeyProperty = nameof(ObjectFlattenerMetric.Name),
                ValueProperty = nameof(ObjectFlattenerMetric.Value)
            };

            var values = flattener.Flatten(new ObjectFlattenerMetricsRow(), options);

            Assert.Equal(2, values["Metrics.HasMX"]);
            Assert.Equal(4, values["Metrics.EffectiveSPFSends"]);
            Assert.False(values.ContainsKey("Metrics."));
        }

        [Fact]
        public void ObjectFlattenerValueTuplePreservesItemPaths() {
            var flattener = new ObjectFlattener();
            var options = new ObjectFlattenerOptions();

            var values = flattener.Flatten((Name: "Alice", Age: 30), options);
            var paths = flattener.GetPaths(typeof((string Name, int Age)), options);

            Assert.Equal("Alice", values["Item1"]);
            Assert.Equal(30, values["Item2"]);
            Assert.Equal(new[] { "Item1", "Item2" }, paths);
        }

        private sealed class ObjectFlattenerCollectionRow {
            public List<string?> Tags { get; } = new() { "a", null, "b" };

            public string[] Empty { get; } = Array.Empty<string>();
        }

        private sealed class ObjectFlattenerSelectionPathRow {
            public int Id { get; set; }

            public ObjectFlattenerSelectionDetails Details { get; set; } = new();

            public ObjectFlattenerIgnoredDetails Ignored { get; set; } = new();
        }

        private sealed class ObjectFlattenerSelectionDetails {
            public int Score { get; set; }

            public string Status { get; set; } = string.Empty;

            public string Secret { get; set; } = string.Empty;
        }

        private sealed class ObjectFlattenerIgnoredDetails {
            public string Value { get; set; } = string.Empty;
        }

        private sealed class ObjectFlattenerMetricsRow {
            public List<ObjectFlattenerMetric?> Metrics { get; } = new() {
                new ObjectFlattenerMetric("HasMX", 2),
                null,
                new ObjectFlattenerMetric(string.Empty, 3),
                new ObjectFlattenerMetric("EffectiveSPFSends", 4)
            };
        }

        private sealed class ObjectFlattenerMetric {
            public ObjectFlattenerMetric(string name, int value) {
                Name = name;
                Value = value;
            }

            public string Name { get; }

            public int Value { get; }
        }
    }
}
