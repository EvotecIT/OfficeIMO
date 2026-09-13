using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private void DrawSummaryCaps(OfficeDrawing drawing, ProjectViewRow row, int offset, int count, double x, double y, double cell) {
        if (!row.Start.HasValue || !row.Finish.HasValue || row.Finish <= row.Start) return;
        void Cap(DateTime date, bool end) {
            if (date < Buckets[offset].Start || date >= Buckets[offset + count - 1].Finish) return;
            double position = x + TimePosition(date, offset, count, cell);
            var shape = OfficeShape.Polygon(new OfficePoint(0, 0), new OfficePoint(6, 0), new OfficePoint(end ? 6 : 0, 7));
            shape.FillColor = Ink; shape.StrokeColor = null;
            drawing.AddShape(shape, position - (end ? 6 : 0), y);
        }
        Cap(row.Start.Value, false); Cap(row.Finish.Value, true);
    }
}
