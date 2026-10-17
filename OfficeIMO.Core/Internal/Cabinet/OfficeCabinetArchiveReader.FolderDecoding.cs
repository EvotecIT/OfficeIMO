using System;
using System.Collections.Generic;

namespace OfficeIMO.Core.Internal.Cabinet {
    internal static partial class OfficeCabinetArchiveReader {
        private sealed partial class Reader {
            /// <summary>Verifies the complete folder while retaining only the requested member ranges.</summary>
            internal byte[][] ReadFolderMembers(FolderHeader folder, int dataReserve,
                IReadOnlyList<FileHeader> files, long maxExpandedBytes) =>
                ReadFolderMembers(ReadFolderFrames(folder, dataReserve, maxExpandedBytes), files);

            private byte[][] ReadFolderMembers(FolderData folder, IReadOnlyList<FileHeader> files) {
                var result = new byte[files.Count][];
                var order = new int[files.Count];
                for (int index = 0; index < files.Count; index++) {
                    _cancellationToken.ThrowIfCancellationRequested();
                    FileHeader file = files[index];
                    if ((long)file.FolderOffset + file.Length > folder.ExpandedBytes)
                        throw Error("CAB_ENTRY_RANGE", "A CAB member extends beyond its decoded folder.");
                    result[index] = new byte[(int)file.Length];
                    order[index] = index;
                }
                Array.Sort(order, (left, right) => files[left].FolderOffset.CompareTo(files[right].FolderOffset));
                var active = new List<int>();
                int next = 0;
                DecodeFolder(folder, CopyFrame);
                _cancellationToken.ThrowIfCancellationRequested();
                return result;

                // Sweep sorted ranges instead of visiting every member for every CFDATA frame.
                // Overlapping authored records and aliases receive independent complete copies.
                void CopyFrame(byte[] buffer, int bufferOffset, int count, long frameOffset) {
                    long frameEnd = frameOffset + count;
                    while (next < order.Length && files[order[next]].FolderOffset < frameEnd) {
                        int index = order[next++];
                        if (files[index].Length != 0) active.Add(index);
                    }
                    int retained = 0;
                    for (int index = 0; index < active.Count; index++) {
                        _cancellationToken.ThrowIfCancellationRequested();
                        int member = active[index];
                        FileHeader file = files[member];
                        long memberEnd = (long)file.FolderOffset + file.Length;
                        long start = Math.Max(frameOffset, file.FolderOffset);
                        long end = Math.Min(frameEnd, memberEnd);
                        if (end > start) Buffer.BlockCopy(buffer, bufferOffset + (int)(start - frameOffset),
                            result[member], (int)(start - file.FolderOffset), (int)(end - start));
                        if (memberEnd > frameEnd) active[retained++] = member;
                    }
                    if (retained < active.Count) active.RemoveRange(retained, active.Count - retained);
                }
            }

            /// <summary>Consumes every admitted frame, with at most the codec window and reusable frame buffers.
            /// Receivers are private staging buffers; final LZX framing must succeed before results escape.</summary>
            private void DecodeFolder(FolderData folder, Action<byte[], int, int, long> receive) {
                int method = folder.Compression & 15;
                byte[]? msZip = method == 1 && folder.Blocks.Length != 0 ? new byte[2 * OfficeMsZipDecoder.WindowBytes] : null;
                var lzx = method == 3 && folder.Blocks.Length != 0
                    ? new global::OfficeIMO.OfficeLzxDecoder(checked((int)folder.ExpandedBytes), (folder.Compression >> 8) & 31, _cancellationToken) : null;
                long expanded = 0;
                try {
                    foreach (BlockHeader block in folder.Blocks) {
                        _cancellationToken.ThrowIfCancellationRequested();
                        if (method == 0) receive(_data, block.DataOffset, block.ExpandedLength, expanded);
                        else if (method == 1) {
                            int dictionary = (int)Math.Min(expanded, OfficeMsZipDecoder.WindowBytes);
                            OfficeMsZipDecoder.Decompress(_data, block.DataOffset, block.CompressedLength,
                                msZip!, dictionary, block.ExpandedLength, _cancellationToken);
                            receive(msZip!, dictionary, block.ExpandedLength, expanded);
                            int keep = Math.Min(dictionary + block.ExpandedLength, OfficeMsZipDecoder.WindowBytes);
                            Buffer.BlockCopy(msZip!, dictionary + block.ExpandedLength - keep, msZip!, 0, keep);
                        } else {
                            byte[] frame = lzx!.DecodeFrame(_data, block.DataOffset, block.CompressedLength, block.ExpandedLength);
                            receive(frame, 0, block.ExpandedLength, expanded);
                        }
                        expanded += block.ExpandedLength;
                    }
                    lzx?.Complete();
                } catch (global::OfficeIMO.OfficeLzxException exception) {
                    throw new OfficeCabinetException("CAB_" + exception.Code, exception.Message, null, exception);
                }
            }

            private sealed class FolderData {
                internal FolderData(BlockHeader[] blocks, long expandedBytes, ushort compression) {
                    Blocks = blocks; ExpandedBytes = expandedBytes; Compression = compression;
                }
                internal BlockHeader[] Blocks { get; }
                internal long ExpandedBytes { get; }
                internal ushort Compression { get; }
            }
        }
    }
}
