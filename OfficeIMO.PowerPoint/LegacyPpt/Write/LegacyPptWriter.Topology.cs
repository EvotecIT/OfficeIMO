namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        /// <summary>
        /// Allocates the contiguous persist-object and OfficeArt drawing identifiers
        /// used by a freshly generated binary presentation.
        /// </summary>
        internal sealed class LegacyPptWriterTopology {
            private const int EmbeddedMasterSlotCount = 11;
            private const int MaxPersistObjectCount = 0x0FFF;
            private const uint MaxDrawingId = 0x0FFFU;

            internal LegacyPptWriterTopology(int masterCount, int slideCount,
                int notesCount, bool hasHandoutMaster) {
                if (masterCount <= 0) {
                    throw new NotSupportedException(
                        "The presentation has no slide master to encode.");
                }
                if (masterCount > MaxNativeMasterCount) {
                    throw new NotSupportedException(
                        $"The binary PowerPoint persist-directory format supports at most {MaxNativeMasterCount} slide masters; the presentation contains {masterCount}.");
                }
                if (slideCount < 0) throw new ArgumentOutOfRangeException(nameof(slideCount));
                if (notesCount < 0) throw new ArgumentOutOfRangeException(nameof(notesCount));

                MasterCount = masterCount;
                MasterSlotCount = Math.Max(EmbeddedMasterSlotCount, masterCount);
                SlideCount = slideCount;
                NotesCount = notesCount;
                HasHandoutMaster = hasHandoutMaster;

                NotesMasterPersistId = checked((uint)(MasterSlotCount + 2));
                FirstSlidePersistId = checked(NotesMasterPersistId + 1U);
                FirstNotesPersistId = checked(FirstSlidePersistId
                    + unchecked((uint)SlideCount));
                HandoutMasterPersistId = HasHandoutMaster
                    ? checked(FirstNotesPersistId + unchecked((uint)NotesCount))
                    : 0U;
                FirstAdditionalPersistId = checked(FirstNotesPersistId
                    + unchecked((uint)NotesCount)
                    + (HasHandoutMaster ? 1U : 0U));
                BasePersistObjectCount = checked((int)FirstAdditionalPersistId - 1);

                NotesMasterDrawingId = checked(unchecked((uint)MasterSlotCount) + 1U);
                FirstSlideDrawingId = checked(NotesMasterDrawingId + 1U);
                FirstNotesDrawingId = checked(FirstSlideDrawingId
                    + unchecked((uint)SlideCount));
                HandoutMasterDrawingId = HasHandoutMaster
                    ? checked(FirstNotesDrawingId + unchecked((uint)NotesCount))
                    : 0U;

                EnsurePersistObjectCapacity(additionalPersistObjectCount: 0);
                uint lastDrawingId = HasHandoutMaster
                    ? HandoutMasterDrawingId
                    : NotesCount > 0
                        ? checked(FirstNotesDrawingId
                            + unchecked((uint)NotesCount) - 1U)
                        : SlideCount > 0
                            ? checked(FirstSlideDrawingId
                                + unchecked((uint)SlideCount) - 1U)
                            : NotesMasterDrawingId;
                if (lastDrawingId > MaxDrawingId) {
                    throw new NotSupportedException(
                        $"The presentation requires OfficeArt drawing identifier {lastDrawingId}, but binary PowerPoint supports identifiers through {MaxDrawingId}.");
                }
            }

            internal int MasterCount { get; }
            internal int MasterSlotCount { get; }
            internal int SlideCount { get; }
            internal int NotesCount { get; }
            internal bool HasHandoutMaster { get; }
            internal int BasePersistObjectCount { get; }
            internal uint NotesMasterPersistId { get; }
            internal uint FirstSlidePersistId { get; }
            internal uint FirstNotesPersistId { get; }
            internal uint HandoutMasterPersistId { get; }
            internal uint FirstAdditionalPersistId { get; }
            internal uint NotesMasterDrawingId { get; }
            internal uint FirstSlideDrawingId { get; }
            internal uint FirstNotesDrawingId { get; }
            internal uint HandoutMasterDrawingId { get; }

            internal uint GetMasterDrawingId(int index) {
                ValidateIndex(index, MasterCount, nameof(index));
                return checked(unchecked((uint)index) + 1U);
            }

            internal uint GetSlidePersistId(int index) {
                ValidateIndex(index, SlideCount, nameof(index));
                return checked(FirstSlidePersistId + unchecked((uint)index));
            }

            internal uint GetSlideDrawingId(int index) {
                ValidateIndex(index, SlideCount, nameof(index));
                return checked(FirstSlideDrawingId + unchecked((uint)index));
            }

            internal uint GetNotesPersistId(int index) {
                ValidateIndex(index, NotesCount, nameof(index));
                return checked(FirstNotesPersistId + unchecked((uint)index));
            }

            internal uint GetNotesDrawingId(int index) {
                ValidateIndex(index, NotesCount, nameof(index));
                return checked(FirstNotesDrawingId + unchecked((uint)index));
            }

            internal void EnsurePersistObjectCapacity(
                int additionalPersistObjectCount) {
                if (additionalPersistObjectCount < 0) {
                    throw new ArgumentOutOfRangeException(
                        nameof(additionalPersistObjectCount));
                }
                int total = checked(BasePersistObjectCount
                    + additionalPersistObjectCount);
                if (total > MaxPersistObjectCount) {
                    throw new NotSupportedException(
                        $"The presentation requires {total} persist objects, but a binary PowerPoint persist-directory run supports at most {MaxPersistObjectCount}.");
                }
            }

            private static void ValidateIndex(int index, int count,
                string parameterName) {
                if (index < 0 || index >= count) {
                    throw new ArgumentOutOfRangeException(parameterName);
                }
            }
        }
    }
}
