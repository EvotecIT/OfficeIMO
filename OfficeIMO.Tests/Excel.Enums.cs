using System;
using OfficeIMO.Excel;
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
    }
}

