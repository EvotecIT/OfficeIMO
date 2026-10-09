using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>Alignment of text following a paragraph tab.</summary>
public enum OfficeTextTabAlignment {
    /// <summary>The following text starts at the stop.</summary>
    Left,
    /// <summary>The following text is centered on the stop.</summary>
    Center,
    /// <summary>The following text ends at the stop.</summary>
    Right,
    /// <summary>The first matching character starts at the stop; an absent character aligns the end.</summary>
    Character
}

/// <summary>A fixed paragraph tab position, measured before render scaling.</summary>
public sealed class OfficeTextTabStop {
    /// <summary>Creates a tab stop. Positions are nonnegative and the character delimiter is one Unicode scalar.</summary>
    public OfficeTextTabStop(double position, OfficeTextTabAlignment alignment = OfficeTextTabAlignment.Left, string character = ".") {
        if (position < 0 || double.IsNaN(position) || double.IsInfinity(position)) throw new ArgumentOutOfRangeException(nameof(position));
        if (!Enum.IsDefined(typeof(OfficeTextTabAlignment), alignment)) throw new ArgumentOutOfRangeException(nameof(alignment));
        if (character == null || !(character.Length == 1 && !char.IsSurrogate(character[0]) ||
            character.Length == 2 && char.IsSurrogatePair(character, 0)) || char.IsControl(character, 0))
            throw new ArgumentException("A tab delimiter must be one non-control Unicode scalar.", nameof(character));
        Position = position; Alignment = alignment; Character = character;
    }
    /// <summary>Position relative to the tab settings' origin.</summary>
    public double Position { get; }
    /// <summary>Alignment of content up to the next tab or hard break.</summary>
    public OfficeTextTabAlignment Alignment { get; }
    /// <summary>Delimiter for character alignment.</summary>
    public string Character { get; }
    /// <summary>Optional single Unicode character repeated in the tab gap.</summary>
    public string? LeaderText { get; private set; }
    /// <summary>Optional vector line paint. A declared <see cref="LeaderText"/> takes precedence, including SPACE.</summary>
    public OfficeTextTabLineLeader? LineLeader { get; private set; }
    /// <summary>Optional overrides for a textual leader; null uses the tab run's formatting.</summary>
    public OfficeTextTabLeaderStyle? LeaderStyle { get; private set; }

    /// <summary>Returns an independent stop with a textual leader; null removes the leader.</summary>
    /// <remarks>The value must be one non-control Unicode scalar. Whitespace leaders retain spacing without paint. Generated leader text is bounded to 100,000 UTF-16 characters per drawing text-frame layout; exhaustion reports clipping without moving the following field.</remarks>
    public OfficeTextTabStop WithLeader(string? text) {
        if (text != null && (!(text.Length == 1 && !char.IsSurrogate(text[0]) || text.Length == 2 && char.IsSurrogatePair(text, 0)) || char.IsControl(text, 0)))
            throw new ArgumentException("A tab leader must be one non-control Unicode scalar.", nameof(text));
        return new OfficeTextTabStop(Position, Alignment, Character) { LeaderText = text, LineLeader = LineLeader, LeaderStyle = LeaderStyle };
    }
    /// <summary>Returns an independent stop with line paint; null removes line paint without changing a textual leader.</summary>
    /// <remarks>Paint is bounded to 8,192 vertices per tab and 100,000 vertices per drawing text-frame layout. Exhaustion reports clipping without moving the following field.</remarks>
    public OfficeTextTabStop WithLineLeader(OfficeTextTabLineLeader? leader) =>
        new OfficeTextTabStop(Position, Alignment, Character) { LeaderText = LeaderText, LineLeader = leader, LeaderStyle = LeaderStyle };
    /// <summary>Returns an independent stop with textual formatting overrides; null removes the overrides without changing the glyph or line declarations.</summary>
    public OfficeTextTabStop WithLeaderStyle(OfficeTextTabLeaderStyle? style) =>
        new OfficeTextTabStop(Position, Alignment, Character) { LeaderText = LeaderText, LineLeader = LineLeader, LeaderStyle = style };
}

/// <summary>Immutable paragraph tab settings shared by SVG, raster and PDF layout.</summary>
public sealed class OfficeTextTabStops {
    /// <summary>Creates sorted, unique stops with a repeating default interval after the last explicit stop.</summary>
    /// <param name="stops">At most 256 explicit stops; the collection is copied.</param>
    /// <param name="defaultInterval">Positive distance between default stops.</param>
    /// <param name="origin">Tab origin relative to the paragraph's inner left margin. May be negative.</param>
    public OfficeTextTabStops(IReadOnlyList<OfficeTextTabStop> stops, double defaultInterval = 36, double origin = 0) {
        if (stops == null) throw new ArgumentNullException(nameof(stops));
        if (stops.Count > 256) throw new ArgumentException("A paragraph supports at most 256 explicit tab stops.", nameof(stops));
        if (defaultInterval <= 0 || double.IsNaN(defaultInterval) || double.IsInfinity(defaultInterval)) throw new ArgumentOutOfRangeException(nameof(defaultInterval));
        if (double.IsNaN(origin) || double.IsInfinity(origin)) throw new ArgumentOutOfRangeException(nameof(origin));
        var copy = new List<OfficeTextTabStop>(stops.Count);
        foreach (OfficeTextTabStop stop in stops) copy.Add(stop ?? throw new ArgumentException("Tab stops cannot contain null.", nameof(stops)));
        copy.Sort((a, b) => a.Position.CompareTo(b.Position));
        for (int i = 1; i < copy.Count; i++) if (copy[i].Position == copy[i - 1].Position) throw new ArgumentException("Tab positions must be unique.", nameof(stops));
        Stops = new ReadOnlyCollection<OfficeTextTabStop>(copy); DefaultInterval = defaultInterval; Origin = origin;
    }
    /// <summary>Sorted snapshot of explicit stops.</summary>
    public IReadOnlyList<OfficeTextTabStop> Stops { get; }
    /// <summary>Distance between repeating default stops.</summary>
    public double DefaultInterval { get; }
    /// <summary>Origin relative to the paragraph's inner left margin.</summary>
    public double Origin { get; }
    /// <summary>Whether tabbed lines follow the paragraph's center or right alignment. False keeps the fixed tab grid.</summary>
    /// <remarks>Tab advances and field alignment remain unchanged within the line. Tabbed lines are never justified.</remarks>
    public bool AlignWithParagraph { get; }

    /// <summary>Returns independent settings that opt tabbed lines into or out of paragraph alignment.</summary>
    /// <param name="align">True applies paragraph center or right alignment to the entire tabbed line; false keeps the fixed tab grid.</param>
    /// <returns>A snapshot with the same stops, default interval and origin.</returns>
    public OfficeTextTabStops WithParagraphAlignment(bool align = true) =>
        new OfficeTextTabStops(Stops, DefaultInterval, Origin, align);

    private OfficeTextTabStops(IReadOnlyList<OfficeTextTabStop> stops, double defaultInterval, double origin, bool alignWithParagraph)
        : this(stops, defaultInterval, origin) {
        AlignWithParagraph = alignWithParagraph;
    }

    internal OfficeTextTabStops Scale(double scale, double originAdjustment = 0, double fontScale = 1) {
        var stops = new List<OfficeTextTabStop>(Stops.Count);
        foreach (OfficeTextTabStop stop in Stops) stops.Add(new OfficeTextTabStop(stop.Position * scale, stop.Alignment, stop.Character)
            .WithLeader(stop.LeaderText).WithLineLeader(stop.LineLeader?.Scale(scale)).WithLeaderStyle(stop.LeaderStyle?.Scale(scale * fontScale)));
        return new OfficeTextTabStops(stops, DefaultInterval * scale, Origin * scale + originAdjustment, AlignWithParagraph);
    }
}
