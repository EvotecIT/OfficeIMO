using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Remaps legacy note shapes alongside comment references during range mutations.</summary>
        internal void RemapMutationCommentVml(Func<ExcelReference, ExcelReference?> transform) {
            CommentList? comments = _worksheetPart.WorksheetCommentsPart?.Comments?.CommentList;
            if (comments == null) return;

            var removed = new List<(int Row, int Col)>();
            var moved = new List<((int Row, int Col) OldCell, (int Row, int Col) NewCell)>();
            foreach (Comment comment in comments.Elements<Comment>()) {
                if (comment.Reference?.Value is not string value
                    || !ExcelReference.TryParse(value, out ExcelReference? reference)) continue;
                ExcelReference? mapped = transform(reference!);
                var oldCell = (reference!.Start.Row, reference.Start.Column);
                if (mapped == null) {
                    removed.Add(oldCell);
                } else if (mapped.Start.Row != oldCell.Row || mapped.Start.Column != oldCell.Column) {
                    moved.Add((oldCell, (mapped.Start.Row, mapped.Start.Column)));
                }
            }

            if (removed.Count > 0 || moved.Count > 0) {
                RemapCommentVmlShapes(removed, moved);
            }
        }
    }
}
