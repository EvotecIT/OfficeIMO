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
        public void ObjectFlattenerJoinCollectionsPreservesNullAndEmptyItems() {
            var flattener = new ObjectFlattener();
            var values = flattener.Flatten(new ObjectFlattenerCollectionRow(), new ObjectFlattenerOptions());

            Assert.Equal("a,,b", values["Tags"]);
            Assert.Equal(string.Empty, values["Empty"]);
        }

        private sealed class ObjectFlattenerCollectionRow {
            public List<string?> Tags { get; } = new() { "a", null, "b" };

            public string[] Empty { get; } = Array.Empty<string>();
        }
    }
}
