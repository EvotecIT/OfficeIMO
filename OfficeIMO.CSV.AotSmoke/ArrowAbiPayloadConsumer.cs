using System.Data;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text;
using OfficeIMO.Data.Arrow;

/// <summary>Consumes raw C Data Interface buffers without Apache's managed importer.</summary>
internal static unsafe class ArrowAbiPayloadConsumer {
    internal static void Verify() {
        using var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Rows.Add(7, "Ada", 1.2300m);
        table.Rows.Add(8, DBNull.Value, DBNull.Value);
        table.Rows.Add(9, "Gräce🙂", -999.9900m);
        table.Rows.Add(0, "", 0.0000m);
        table.Rows.Add(DBNull.Value, "Last", 2.5000m);
        using var reader = table.CreateDataReader();
        using var owner = reader.ExportArrowCStream(new ArrowReadOptions { BatchSize = 2, DecimalPrecision = 5, DecimalScale = 2 });
        using var lease = owner.AcquireLease();
        var stream = (NativeArrowArrayStream*)lease.Address;
        var schema = new NativeArrowSchema();
        CheckStatus(stream->get_schema(stream, &schema));
        try {
            Require(schema.n_children == 3 && Text(schema.format) == "+s", "record schema");
            Require(Text(schema.children[0]->format) == "i" && Text(schema.children[1]->format) == "u" &&
                Text(schema.children[2]->format) == "d:5,2", "column formats");
            Require(Text(schema.children[0]->name) == "Id" && Text(schema.children[1]->name) == "Name" && Text(schema.children[2]->name) == "Amount", "column names");
        } finally {
            if (schema.release != null) schema.release(&schema);
        }
        Require(schema.release == null, "schema release");
        int row = 0, batches = 0;
        while (true) {
            var batch = new NativeArrowArray();
            CheckStatus(stream->get_next(stream, &batch));
            if (batch.release == null) break;
            try {
                Require(batch.length is > 0 and <= 2 && batch.n_children == 3, "bounded batch");
                for (long local = 0; local < batch.length; local++, row++) {
                    Require(row < table.Rows.Count, "row count");
                    ValidateInteger(batch.children[0], local + batch.offset, table.Rows[row][0]);
                    ValidateString(batch.children[1], local + batch.offset, table.Rows[row][1]);
                    ValidateDecimal(batch.children[2], local + batch.offset, table.Rows[row][2]);
                }
                batches++;
            } finally {
                if (batch.release != null) batch.release(&batch);
            }
            Require(batch.release == null, "batch release");
        }
        Require(row == 5 && batches == 3, "complete stream");
        stream->release(stream);
        Require(stream->release == null, "stream release");
    }

    private static bool IsNull(NativeArrowArray* array, long row, object expected, long buffers) {
        Require(array->n_buffers == buffers && array->buffers != null, "buffer layout");
        long index = array->offset + row;
        Require(index >= array->offset && row < array->length, "array bounds");
        byte* bitmap = array->buffers[0];
        bool isNull = bitmap != null && (bitmap[index / 8] & (1 << (int)(index % 8))) == 0;
        Require(isNull == ReferenceEquals(expected, DBNull.Value), "null bitmap");
        return isNull;
    }

    private static void ValidateInteger(NativeArrowArray* array, long row, object expected) {
        if (IsNull(array, row, expected, 2)) return;
        Require(array->buffers[1] != null && ((int*)array->buffers[1])[array->offset + row] == (int)expected, "integer payload");
    }

    private static void ValidateString(NativeArrowArray* array, long row, object expected) {
        if (IsNull(array, row, expected, 3)) return;
        Require(array->buffers[1] != null, "string offsets");
        long index = array->offset + row;
        int start = ((int*)array->buffers[1])[index], end = ((int*)array->buffers[1])[index + 1];
        Require(start >= 0 && end >= start && (end == start || array->buffers[2] != null), "string range");
        string value = end == start ? string.Empty : Encoding.UTF8.GetString(array->buffers[2] + start, end - start);
        Require(value == (string)expected, "UTF-8 payload");
    }

    private static void ValidateDecimal(NativeArrowArray* array, long row, object expected) {
        if (IsNull(array, row, expected, 2)) return;
        Require(array->buffers[1] != null, "decimal values");
        var coefficient = new BigInteger(new ReadOnlySpan<byte>(array->buffers[1] + (array->offset + row) * 16, 16), isUnsigned: false, isBigEndian: !BitConverter.IsLittleEndian);
        Require(coefficient == new BigInteger((decimal)expected * 100m), "signed decimal coefficient");
    }

    private static string? Text(byte* value) => Marshal.PtrToStringUTF8((nint)value);
    private static void CheckStatus(int status) => Require(status == 0, "callback status " + status);
    private static void Require(bool condition, string contract) {
        if (!condition) throw new InvalidOperationException("Arrow raw ABI consumer failed: " + contract + ".");
    }
}
