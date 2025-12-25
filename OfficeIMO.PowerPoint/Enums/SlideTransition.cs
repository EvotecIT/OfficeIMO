namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents simple slide transitions.
    /// </summary>
    public enum SlideTransition {
        /// <summary>No transition is applied.</summary>
        None,

        /// <summary>Fade between slides.</summary>
        Fade,

        /// <summary>Wipe transition between slides.</summary>
        Wipe,

        /// <summary>Blinds transition (vertical).</summary>
        BlindsVertical,

        /// <summary>Blinds transition (horizontal).</summary>
        BlindsHorizontal,

        /// <summary>Comb transition (horizontal).</summary>
        CombHorizontal,

        /// <summary>Comb transition (vertical).</summary>
        CombVertical,

        /// <summary>Push transition (up).</summary>
        PushUp,

        /// <summary>Push transition (down).</summary>
        PushDown,

        /// <summary>Push transition (left).</summary>
        PushLeft,

        /// <summary>Push transition (right).</summary>
        PushRight,

        /// <summary>Cut transition.</summary>
        Cut,

        /// <summary>Flash transition.</summary>
        Flash,

        /// <summary>Warp transition (in).</summary>
        WarpIn,

        /// <summary>Warp transition (out).</summary>
        WarpOut,

        /// <summary>Prism transition.</summary>
        Prism,

        /// <summary>Ferris transition (left).</summary>
        FerrisLeft,

        /// <summary>Ferris transition (right).</summary>
        FerrisRight,

        /// <summary>Morph transition (by object).</summary>
        Morph
    }
}
