using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Data;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Verifies accessibility and defaults of Excel-related enums.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void TableStyleHasValues() {
            Assert.True(Enum.IsDefined(typeof(ExcelTableStyle), nameof(ExcelTableStyle.TableStyleLight1)));
        }

        [Fact]
        public void ExecutionPolicyDefaultsToAutomatic() {
            var policy = new ExcelExecutionPolicy();
            Assert.Equal(ExcelExecutionMode.Automatic, policy.Mode);
        }

        [Fact]
        public void ObjectFlattenerOptionsDefaults() {
            var opts = new ObjectFlattenerOptions();
            Assert.Equal(HeaderCase.Raw, opts.HeaderCase);
            Assert.Equal(NullPolicy.NullLiteral, opts.NullPolicy);
            Assert.Equal(CollectionMode.JoinWith, opts.CollectionMode);
        }

        [Fact]
        public void ObjectFlattenerResolvePathsPreservesPinsPrioritiesAndDiscoveryOrder() {
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

            var ordered = new ObjectFlattener().ResolvePaths(input, opts);

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
        public void ObjectFlattenerResolvePathsCombinesIgnoreExcludeAndInclude() {
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

            var selected = new ObjectFlattener().ResolvePaths(input, opts);

            Assert.Equal(new[] {
                "Id",
                "Details.Score",
                "Details.Status"
            }, selected);
        }

        [Fact]
        public void ObjectFlattenerResolvePathsStopsEnumerationAtColumnLimit() {
            int enumeratedPaths = 0;

            IEnumerable<string> Paths() {
                while (true) yield return "Column" + enumeratedPaths++;
            }

            Assert.Throws<System.IO.InvalidDataException>(() =>
                new ObjectFlattener().ResolvePaths(
                    Paths(),
                    new ObjectFlattenerOptions { MaxColumns = 1 }));
            Assert.Equal(2, enumeratedPaths);
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
        public void ObjectFlattenerCollectionMapColumnsStopsAtColumnLimit() {
            var options = new ObjectFlattenerOptions { MaxColumns = 1 };
            options.CollectionMapColumns["Metrics"] = new CollectionColumnMapping {
                KeyProperty = nameof(ObjectFlattenerMetric.Name),
                ValueProperty = nameof(ObjectFlattenerMetric.Value)
            };

            Assert.Throws<System.IO.InvalidDataException>(() =>
                new ObjectFlattener().Flatten(new ObjectFlattenerMetricsRow(), options));
        }

        [Fact]
        public void ObjectFlattenerExplicitColumnsCountsResolvedDistinctPaths() {
            var options = new ObjectFlattenerOptions {
                Columns = new[] { "Name", "Name", "Missing" },
                MaxColumns = 1
            };

            Dictionary<string, object?> values = new ObjectFlattener().Flatten(
                new ObjectFlattenerSinglePropertyRow { Name = "Alice" }, options);

            Assert.Single(values);
            Assert.Equal("Alice", values["Name"]);
        }

        [Fact]
        public void ObjectFlattenerExplicitColumnsProjectsSubsetBeforeColumnLimit() {
            var options = new ObjectFlattenerOptions {
                Columns = new[] { "Name" },
                MaxColumns = 1
            };

            Dictionary<string, object?> values = new ObjectFlattener().Flatten(
                new ObjectFlattenerWideRow { Name = "Alice", Status = "Active" }, options);

            Assert.Single(values);
            Assert.Equal("Alice", values["Name"]);
        }

        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void ObjectFlattenerAppliesIncludeAndExcludeBeforeColumnLimit(bool useInclude) {
            var options = new ObjectFlattenerOptions { MaxColumns = 1 };
            if (useInclude) {
                options.IncludeProperties = new[] { "Name" };
            } else {
                options.ExcludeProperties = new[] { "Status" };
            }

            Dictionary<string, object?> values = new ObjectFlattener().Flatten(
                new ObjectFlattenerWideRow { Name = "Alice", Status = "Active" }, options);

            Assert.Single(values);
            Assert.Equal("Alice", values["Name"]);
        }

        [Fact]
        public void ObjectFlattenerExplicitColumnsProjectsDictionarySubsetBeforeColumnLimit() {
            var options = new ObjectFlattenerOptions {
                Columns = new[] { "Name" },
                MaxColumns = 1
            };
            var row = new Dictionary<string, object?> {
                ["Status"] = "Active",
                ["Name"] = "Alice"
            };

            Dictionary<string, object?> values = new ObjectFlattener().Flatten(row, options);

            Assert.Single(values);
            Assert.Equal("Alice", values["Name"]);
        }

        [Fact]
        public void ObjectFlattenerAppliesDictionaryIgnoreBeforeColumnLimit() {
            var options = new ObjectFlattenerOptions {
                Ignore = new[] { "Status" },
                MaxColumns = 1
            };
            var row = new Dictionary<string, object?> {
                ["Status"] = "Active",
                ["Name"] = "Alice"
            };

            Dictionary<string, object?> values = new ObjectFlattener().Flatten(row, options);

            Assert.Single(values);
            Assert.Equal("Alice", values["Name"]);
        }

        [Fact]
        public void ObjectFlattenerCountsDistinctResolvedDictionaryPathsAtColumnLimit() {
            var row = new Dictionary<string, object?>(StringComparer.Ordinal) {
                ["Name"] = "first",
                ["name"] = "last"
            };

            Dictionary<string, object?> values = new ObjectFlattener().Flatten(
                row,
                new ObjectFlattenerOptions {
                    MaxColumns = 1,
                    MaxCollectionItems = 2
                });

            Assert.Single(values);
            Assert.Equal("last", values["Name"]);
            Assert.Equal("Name", Assert.Single(values).Key);

            var heterogeneousKeys = new Dictionary<object, object?> {
                [1] = "number",
                ["1"] = "text"
            };
            Dictionary<string, object?> heterogeneousValues = new ObjectFlattener().Flatten(
                heterogeneousKeys,
                new ObjectFlattenerOptions {
                    MaxColumns = 1,
                    MaxCollectionItems = 2
                });

            Assert.Single(heterogeneousValues);
            Assert.True(heterogeneousValues.ContainsKey("1"));
        }

        [Fact]
        public void ObjectFlattenerIndexesMaximumWidthExplicitColumnSelection() {
            const int columnCount = ObjectFlattenerOptions.DefaultMaxColumns;
            string[] columns = Enumerable.Range(0, columnCount)
                .Select(index => "Column" + index)
                .ToArray();
            var row = columns.ToDictionary(
                column => column,
                column => (object?)column,
                StringComparer.Ordinal);

            Dictionary<string, object?> values = new ObjectFlattener().Flatten(
                row,
                new ObjectFlattenerOptions {
                    Columns = columns,
                    MaxColumns = columnCount,
                    MaxCollectionItems = columnCount
                });

            Assert.Equal(columnCount, values.Count);
            Assert.Equal("Column0", values["Column0"]);
            Assert.Equal("Column16383", values["Column16383"]);
        }

        [Fact]
        public void ObjectFlattenerStopsRecursiveProjectionAtColumnLimit() {
            var options = new ObjectFlattenerOptions { MaxColumns = 1 };
            options.ExpandProperties.Add(nameof(ObjectFlattenerBranch.Left));
            options.ExpandProperties.Add(nameof(ObjectFlattenerBranch.Right));
            var row = new ObjectFlattenerBranch {
                Value = 1,
                Left = new ObjectFlattenerBranch { Value = 2 },
                Right = new ObjectFlattenerBranch { Value = 3 }
            };

            Assert.Throws<System.IO.InvalidDataException>(() =>
                new ObjectFlattener().Flatten(row, options));
            Assert.False(row.WasRightRead());
        }

        [Fact]
        public void ObjectFlattenerStopsRecursiveTypePathDiscoveryAtColumnLimit() {
            var options = new ObjectFlattenerOptions {
                MaxColumns = 1,
                MaxDepth = 12
            };
            options.ExpandProperties.Add(nameof(ObjectFlattenerRecursiveType.Left));
            options.ExpandProperties.Add(nameof(ObjectFlattenerRecursiveType.Right));

            Assert.Throws<System.IO.InvalidDataException>(() =>
                new ObjectFlattener().GetPaths(typeof(ObjectFlattenerRecursiveType), options));
        }

        [Fact]
        public void ObjectFlattenerTypePathDiscoveryCountsCaseInsensitiveDistinctPaths() {
            List<string> paths = new ObjectFlattener().GetPaths(
                typeof(ObjectFlattenerCaseDistinctRow),
                new ObjectFlattenerOptions { MaxColumns = 1 });

            Assert.Equal("A", Assert.Single(paths));
        }

        [Fact]
        public void ObjectFlattenerExplicitColumnsRejectsResolvedDistinctPathsBeyondLimit() {
            var options = new ObjectFlattenerOptions {
                Columns = new[] { "Name", "Status" },
                MaxColumns = 1
            };

            Assert.Throws<System.IO.InvalidDataException>(() =>
                new ObjectFlattener().ResolvePaths(new[] { "Name", "Status" }, options));
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

        private sealed class ObjectFlattenerSinglePropertyRow {
            public string Name { get; set; } = string.Empty;
        }

        private sealed class ObjectFlattenerWideRow {
            public string Name { get; set; } = string.Empty;

            public string Status { get; set; } = string.Empty;
        }

        private sealed class ObjectFlattenerBranch {
            private ObjectFlattenerBranch? _right;
            private bool _rightWasRead;

            public int Value { get; set; }

            public ObjectFlattenerBranch? Left { get; set; }

            public ObjectFlattenerBranch? Right {
                get {
                    _rightWasRead = true;
                    return _right;
                }
                set => _right = value;
            }

            public bool WasRightRead() => _rightWasRead;
        }

        private sealed class ObjectFlattenerRecursiveType {
            public ObjectFlattenerRecursiveType? Left { get; set; }

            public ObjectFlattenerRecursiveType? Right { get; set; }

            public int Value { get; set; }
        }

        private sealed class ObjectFlattenerCaseDistinctRow {
            public int A { get; set; }

            public int a { get; set; }
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
