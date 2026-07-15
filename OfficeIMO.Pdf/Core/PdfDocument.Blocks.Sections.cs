namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    private int nextGeneratedSectionId = 1;
    private int generatedSectionLayoutDepth;
    private int nextLayoutGeneratedSectionId;
    private Dictionary<object, List<int>>? layoutGeneratedSectionIdsByOwner;
    private object? generatedSectionMaterializationOwner;
    private int generatedSectionMaterializationOrdinal;

    /// <summary>Adds a semantic section backed by an outline entry and named destination.</summary>
    public PdfDocument Section(
        string title,
        Action<PdfItemCompose> compose,
        PdfSectionOptions? options = null) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));
        Guard.NotNull(compose, nameof(compose));
        PdfSectionOptions requested = options ?? new PdfSectionOptions();
        string destinationName = string.IsNullOrWhiteSpace(requested.DestinationName)
            ? "section-" + TakeNextGeneratedSectionId().ToString(System.Globalization.CultureInfo.InvariantCulture)
            : requested.DestinationName!.Trim();
        var blocks = new List<IPdfBlock>();
        using (PushBlockScope(blocks.Add)) {
            compose(new PdfItemCompose(this));
        }

        AddBlock(new SectionBlock(title, blocks, requested.Clone(destinationName)));
        return this;
    }

    /// <summary>Adds an automatically generated, internally linked table of contents.</summary>
    public PdfDocument TableOfContents(PdfTableOfContentsOptions? options = null) {
        AddBlock(new TableOfContentsBlock(options));
        return this;
    }

    internal System.IDisposable BeginGeneratedSectionLayout() {
        if (generatedSectionLayoutDepth++ == 0) {
            nextLayoutGeneratedSectionId = nextGeneratedSectionId;
            layoutGeneratedSectionIdsByOwner = new Dictionary<object, List<int>>();
            generatedSectionMaterializationOwner = null;
            generatedSectionMaterializationOrdinal = 0;
        }

        return new GeneratedSectionLayoutScope(this);
    }

    private GeneratedSectionMaterializationScope BeginGeneratedSectionMaterialization(object owner) {
        object? previousOwner = generatedSectionMaterializationOwner;
        int previousOrdinal = generatedSectionMaterializationOrdinal;
        generatedSectionMaterializationOwner = owner;
        generatedSectionMaterializationOrdinal = 0;
        return new GeneratedSectionMaterializationScope(this, previousOwner, previousOrdinal);
    }

    private int TakeNextGeneratedSectionId() {
        if (generatedSectionLayoutDepth <= 0) return nextGeneratedSectionId++;
        if (generatedSectionMaterializationOwner == null || layoutGeneratedSectionIdsByOwner == null) {
            return nextLayoutGeneratedSectionId++;
        }

        if (!layoutGeneratedSectionIdsByOwner.TryGetValue(generatedSectionMaterializationOwner, out List<int>? sectionIds)) {
            sectionIds = new List<int>();
            layoutGeneratedSectionIdsByOwner.Add(generatedSectionMaterializationOwner, sectionIds);
        }

        int ordinal = generatedSectionMaterializationOrdinal++;
        if (ordinal < sectionIds.Count) return sectionIds[ordinal];
        int sectionId = nextLayoutGeneratedSectionId++;
        sectionIds.Add(sectionId);
        return sectionId;
    }

    private void EndGeneratedSectionLayout() {
        if (generatedSectionLayoutDepth <= 0 || --generatedSectionLayoutDepth != 0) return;
        nextLayoutGeneratedSectionId = 0;
        layoutGeneratedSectionIdsByOwner = null;
        generatedSectionMaterializationOwner = null;
        generatedSectionMaterializationOrdinal = 0;
    }

    private sealed class GeneratedSectionLayoutScope : System.IDisposable {
        private PdfDocument? document;

        public GeneratedSectionLayoutScope(PdfDocument document) {
            this.document = document;
        }

        public void Dispose() {
            PdfDocument? current = document;
            if (current == null) return;
            document = null;
            current.EndGeneratedSectionLayout();
        }
    }

    private sealed class GeneratedSectionMaterializationScope : System.IDisposable {
        private PdfDocument? document;
        private readonly object? previousOwner;
        private readonly int previousOrdinal;

        public GeneratedSectionMaterializationScope(PdfDocument document, object? previousOwner, int previousOrdinal) {
            this.document = document;
            this.previousOwner = previousOwner;
            this.previousOrdinal = previousOrdinal;
        }

        public void Dispose() {
            PdfDocument? current = document;
            if (current == null) return;
            document = null;
            current.generatedSectionMaterializationOwner = previousOwner;
            current.generatedSectionMaterializationOrdinal = previousOrdinal;
        }
    }
}
