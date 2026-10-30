using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Current complete stories, frame definitions and resolved chain order, independent of native text records.</summary>
internal sealed class PublisherTextState {
    private readonly HashSet<uint> _tableStories;
    internal PublisherTextState(IReadOnlyDictionary<uint, PublisherTextStory> stories,
        IReadOnlyDictionary<uint, PublisherFrameState> frames, IReadOnlyDictionary<uint, IReadOnlyList<uint>> chains,
        IEnumerable<uint> tableStoryIds, IReadOnlyDictionary<uint, PublisherTableStory> tableStories,
        IReadOnlyDictionary<uint, PublisherWrapInfo> wrapInfo, IReadOnlyDictionary<uint, PublisherImportedWrapObject> obstacles) {
        Stories = ReadOnly(stories); Frames = ReadOnly(frames);
        Chains = ReadOnly(chains.ToDictionary(item => item.Key, item => (IReadOnlyList<uint>)Array.AsReadOnly(item.Value.ToArray())));
        _tableStories = new HashSet<uint>(tableStoryIds.Concat(tableStories.Keys));
        TableStories = ReadOnly(tableStories);
        WrapInfo = ReadOnly(wrapInfo); ImportedObstacles = ReadOnly(obstacles);
    }

    private PublisherTextState(PublisherTextState previous, IReadOnlyDictionary<uint, PublisherTextStory>? stories = null,
        IReadOnlyDictionary<uint, PublisherFrameState>? frames = null,
        IReadOnlyDictionary<uint, IReadOnlyList<uint>>? chains = null,
        IReadOnlyDictionary<uint, PublisherTableStory>? tableStories = null) {
        Stories = stories == null ? previous.Stories : ReadOnly(stories);
        Frames = frames == null ? previous.Frames : ReadOnly(frames);
        Chains = chains == null ? previous.Chains : ReadOnly(chains);
        TableStories = tableStories == null ? previous.TableStories : ReadOnly(tableStories);
        _tableStories = new HashSet<uint>(previous._tableStories.Concat(TableStories.Keys));
        WrapInfo = previous.WrapInfo; ImportedObstacles = previous.ImportedObstacles;
    }

    internal IReadOnlyDictionary<uint, PublisherTextStory> Stories { get; }
    internal IReadOnlyDictionary<uint, PublisherFrameState> Frames { get; }
    internal IReadOnlyDictionary<uint, IReadOnlyList<uint>> Chains { get; }
    internal IReadOnlyDictionary<uint, PublisherTableStory> TableStories { get; }
    internal IReadOnlyDictionary<uint, PublisherWrapInfo> WrapInfo { get; }
    internal IReadOnlyDictionary<uint, PublisherImportedWrapObject> ImportedObstacles { get; }
    internal bool IsTableStory(uint storyId) => _tableStories.Contains(storyId);
    internal int AuthoredFrameCount => Frames.Values.Count(frame => frame.IsAuthored);

    internal PublisherTextState ReplaceStory(PublisherTextStory story) {
        var stories = Stories.ToDictionary(item => item.Key, item => item.Value); stories[story.Id] = story;
        return new PublisherTextState(this, stories: stories);
    }

    internal PublisherTextState WithTableStory(PublisherTextStory story, PublisherTableStory partition) {
        var stories = Stories.ToDictionary(item => item.Key, item => item.Value); stories[story.Id] = story;
        var partitions = TableStories.ToDictionary(item => item.Key, item => item.Value); partitions[story.Id] = partition;
        return new PublisherTextState(this, stories: stories, tableStories: partitions);
    }

    internal PublisherTextState SetBounds(uint id, double x, double y, double width, double height) {
        if (!Frames.TryGetValue(id, out PublisherFrameState? frame) || !frame.CanEditBounds)
            throw new NotSupportedException("Frame bounds can only be edited for visible, resolved, untransformed top-level frames.");
        frame.EditRange.Validate(x, y, width, height);
        var frames = Frames.ToDictionary(item => item.Key, item => item.Value); frames[id] = frame.WithBounds(x, y, width, height);
        return new PublisherTextState(this, frames: frames);
    }

    internal PublisherTextState AppendFrame(PublisherFrameState previous, PublisherPageState page,
        double x, double y, double width, double height, uint frameId) {
        page.FrameEditRange.Validate(x, y, width, height);
        var frame = new PublisherFrameState(frameId, previous.StoryId, page.Id, x, y, width, height,
            OfficeTransform.Translate(x, y), true, true, false, true, true, true, previous.ColumnCount, previous.ColumnSpacing,
            previous.Padding, previous.VerticalAlignment, Array.Empty<uint>(), page.FrameEditRange, true);
        var frames = Frames.ToDictionary(item => item.Key, item => item.Value); frames.Add(frameId, frame);
        var chains = Chains.ToDictionary(item => item.Key, item => item.Value);
        chains[previous.StoryId] = Array.AsReadOnly(Chains[previous.StoryId].Concat(new[] { frameId }).ToArray());
        return new PublisherTextState(this, frames: frames, chains: chains);
    }

    private static IReadOnlyDictionary<TKey, TValue> ReadOnly<TKey, TValue>(IReadOnlyDictionary<TKey, TValue> values) where TKey : notnull =>
        new System.Collections.ObjectModel.ReadOnlyDictionary<TKey, TValue>(values.ToDictionary(item => item.Key, item => item.Value));
}

/// <summary>Imported wrap intent. Placement comes from the current frame/picture when that object has a semantic owner.</summary>
internal sealed record PublisherWrapInfo(OfficeTextPadding Margins, bool NegativeMargins, bool Tight);

/// <summary>Immutable wrap placement for object kinds whose editable geometry has not yet been migrated.</summary>
internal sealed record PublisherImportedWrapObject(uint? PageId, double X, double Y, double Width, double Height,
    OfficeTransform PageTransform, bool Available, bool Hidden);
