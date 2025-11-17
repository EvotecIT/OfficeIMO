using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SdtIdAllocator_AssignsSequentialPositiveIds() {
            using var document = WordDocument.Create();

            var ids = Enumerable.Range(0, 5).Select(_ => document.GenerateSdtId()).ToArray();

            Assert.All(ids, id => Assert.InRange(id, 1, int.MaxValue - 1));
            Assert.True(ids.SequenceEqual(new[] { 1, 2, 3, 4, 5 }));
        }

        [Fact]
        public void SdtIdAllocator_WrapsAfterMaxValue() {
            using var document = WordDocument.Create();

            var nextField = typeof(WordDocument).GetField("_nextSdtId", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(nextField);
            nextField!.SetValue(document, int.MaxValue - 1);

            var highValue = document.GenerateSdtId();
            Assert.Equal(int.MaxValue - 1, highValue);

            var wrapped = document.GenerateSdtId();
            Assert.Equal(1, wrapped);
        }

        [Fact]
        public void SdtIdAllocator_RespectsExistingIdsOnReload() {
            using var document = WordDocument.Create();

            document._document.Body ??= new Body();
            document._document.Body.AppendChild(new SdtBlock(
                new SdtProperties(new SdtId() { Val = 7 }),
                new SdtContentBlock(new Paragraph(new Run(new Text("Hello"))))
            ));

            document._document.Body.AppendChild(new SdtBlock(
                new SdtProperties(new SdtId() { Val = -5 }),
                new SdtContentBlock(new Paragraph(new Run(new Text("Ignored"))))
            ));

            var initialize = CreateInitializeDelegate(document);
            initialize();

            var id = document.GenerateSdtId();

            Assert.Equal(8, id);
        }

        [Fact]
        public void SdtIdAllocator_IsThreadSafe() {
            using var document = WordDocument.Create();

            var bag = new ConcurrentBag<int>();

            Parallel.For(0, 64, _ => bag.Add(document.GenerateSdtId()));

            Assert.Equal(64, bag.Distinct().Count());
            Assert.All(bag, id => Assert.InRange(id, 1, int.MaxValue - 1));
        }

        private static Action CreateInitializeDelegate(WordDocument document) {
            var method = typeof(WordDocument).GetMethod("InitializeSdtIdState", BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(method);
            return (Action)method!.CreateDelegate(typeof(Action), document);
        }
    }
}
