"""Exercise the published DLL through Windows' ODBC manager without driver registration."""
import argparse
import ctypes as c
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
import uuid
import winreg


def workbook(path):
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Customers" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:D3"/><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Country</t></is></c><c r="C1" t="inlineStr"><is><t>Amount</t></is></c><c r="D1" t="inlineStr"><is><t>Optional</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Alicja Żółć</t></is></c><c r="B2" t="inlineStr"><is><t>Poland</t></is></c><c r="C2"><v>12.5</v></c><c r="D2" t="inlineStr"><is><t></t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Bob</t></is></c><c r="B3" t="inlineStr"><is><t>Germany</t></is></c><c r="C3"><v>20</v></c></row></sheetData></worksheet>',
    }
    sheet = parts["xl/worksheets/sheet1.xml"].replace("A1:D3", "A1:F3")
    sheet = sheet.replace('</row>', '<c r="E1" t="inlineStr"><is><t>Active</t></is></c><c r="F1" t="inlineStr"><is><t>Recorded</t></is></c></row>', 1)
    sheet = sheet.replace('</row><row r="3">', '<c r="E2" t="b"><v>1</v></c><c r="F2" t="d"><v>2026-10-12T01:02:03.1234567</v></c></row><row r="3">')
    parts["xl/worksheets/sheet1.xml"] = sheet
    # Workbook order deliberately differs from the required catalog order.
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace('<sheets>', '<sheets><sheet name="Z" sheetId="2" r:id="rId2"/>').replace('</sheets>', '<sheet name="A" sheetId="3" r:id="rId3"/></sheets>')
    for number in (2, 3):
        parts[f"xl/worksheets/sheet{number}.xml"] = sheet
        parts["xl/_rels/workbook.xml.rels"] = parts["xl/_rels/workbook.xml.rels"].replace('</Relationships>', f'<Relationship Id="rId{number}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{number}.xml"/></Relationships>')
        parts["[Content_Types].xml"] = parts["[Content_Types].xml"].replace('</Types>', f'<Override PartName="/xl/worksheets/sheet{number}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, xml in parts.items():
            archive.writestr(name, xml.encode("utf-8"))


class Manager:
    def __init__(self, library="odbc32.dll"):
        self.dll = c.WinDLL(str(library))
        h, s, u, i, p, n, w = c.c_void_p, c.c_short, c.c_ushort, c.c_int, c.c_void_p, c.c_ssize_t, c.c_wchar_p
        signatures = {
            "SQLAllocHandle": [s, h, p], "SQLFreeHandle": [s, h],
            "SQLSetEnvAttr": [h, i, h, i], "SQLDriverConnectW": [h, h, w, s, p, s, p, u],
            "SQLGetDiagRecW": [s, h, s, p, p, p, s, p], "SQLDisconnect": [h],
            "SQLExecDirectW": [h, w, i], "SQLFetch": [h], "SQLGetData": [h, u, s, p, n, p],
            "SQLNumResultCols": [h, p], "SQLNumParams": [h, p], "SQLCloseCursor": [h], "SQLPrepareW": [h, w, i],
            "SQLBindParameter": [h, u, s, s, s, c.c_size_t, s, p, n, p], "SQLExecute": [h],
            "SQLTablesW": [h, w, s, w, s, w, s, w, s], "SQLColumnsW": [h, w, s, w, s, w, s, w, s],
            "SQLBindCol": [h, u, s, p, n, p], "SQLFreeStmt": [h, u], "SQLCancel": [h],
            "SQLGetInfoW": [h, u, p, s, p], "SQLGetFunctions": [h, u, p],
            "SQLSetStmtAttrW": [h, i, h, i],
            "SQLGetStmtAttrW": [h, i, p, i, p], "SQLGetEnvAttr": [h, i, p, i, p],
            "SQLGetDescFieldW": [h, s, s, p, i, p], "SQLGetDiagFieldW": [s, h, s, s, p, s, p],
            "SQLGetConnectAttrW": [h, i, p, i, p], "SQLEndTran": [s, h, s],
            "SQLFetchScroll": [h, s, n], "SQLGetTypeInfoW": [h, s],
            "SQLColAttributeW": [h, u, u, p, s, p, p],
            "SQLDescribeColW": [h, u, p, s, p, p, p, p, p],
        }
        for name, args in signatures.items():
            function = getattr(self.dll, name)
            function.argtypes, function.restype = args, s
            setattr(self, name, function)

    def diagnostics(self, kind, handle):
        records = []
        for record in range(1, 20):
            state, message = c.create_unicode_buffer(6), c.create_unicode_buffer(1024)
            native, length = c.c_int(), c.c_short()
            result = self.SQLGetDiagRecW(kind, handle, record, state, c.byref(native), message, 1024, c.byref(length))
            if result == 100:
                break
            assert result in (0, 1), result
            records.append((state.value, message.value))
        return records

    def check(self, result, kind, handle):
        assert result in (0, 1), (result, self.diagnostics(kind, handle))
        return result

    def allocate(self, kind, parent=None):
        handle = c.c_void_p()
        self.check(self.SQLAllocHandle(kind, parent, c.byref(handle)), kind - 1, parent)
        return handle

    def text(self, statement, column, capacity=256):
        chunks = []
        while True:
            buffer, length = c.create_unicode_buffer(capacity), c.c_ssize_t()
            result = self.SQLGetData(statement, column, -8, buffer, c.sizeof(buffer), c.byref(length))
            self.check(result, 3, statement)
            if length.value == -1:
                return None
            chunks.append(buffer.value)
            if result == 0:
                return "".join(chunks)
            assert self.diagnostics(3, statement)[0][0] == "01004"


def verify(driver_name, path, ado_consumer, hive):
    m = Manager()
    original = hashlib.sha256(path.read_bytes()).digest()
    connection_string = "DRIVER={" + driver_name + "};DBQ={" + str(path).replace("}", "}}") + "};ReadOnly=1;"
    # Repeated load/disconnect/free checks the native library's pinned lifetime.
    for iteration in range(3):
        environment = m.allocate(1)
        connection = statement = None
        connected = False
        try:
            m.check(m.SQLSetEnvAttr(environment, 200, 3, 0), 1, environment)
            connection = m.allocate(2, environment)
            output, length = c.create_unicode_buffer(2048), c.c_short()
            m.check(m.SQLDriverConnectW(connection, None, connection_string, -3, output, 2048, c.byref(length), (0, 1, 3)[iteration]), 2, connection)
            assert output.value == connection_string, "Completed string must select the same driver on reconnect"
            connection_string = output.value
            connected = True
            statement = m.allocate(3, connection)

            def execute(sql):
                m.check(m.SQLExecDirectW(statement, sql, -3), 3, statement)

            def close():
                m.check(m.SQLFreeStmt(statement, 0), 3, statement)

            # SQL-to-C conversion truncates fractions and reports loss rather than rounding.
            for target, scalar in [(5, c.c_short), (-17, c.c_ushort), (4, c.c_int),
                                   (-18, c.c_uint), (-25, c.c_longlong), (-27, c.c_ulonglong), (-7, c.c_ubyte)]:
                for number in ([1.9, -1.9] if target in (5, 4, -25) else [1.9]):
                    execute('SELECT ' + str(number))
                    m.check(m.SQLFetch(statement), 3, statement)
                    value, indicator = scalar(), c.c_ssize_t()
                    assert m.SQLGetData(statement, 1, target, c.byref(value), c.sizeof(value), c.byref(indicator)) == 1
                    assert value.value == int(number) and m.diagnostics(3, statement)[0][0] == '01S07'
                    close()
            for number, target, scalar in [(32768, 5, c.c_short), (-1, -17, c.c_ushort),
                                           (2147483648, 4, c.c_int), (4294967296, -18, c.c_uint),
                                           (9223372036854775808, -25, c.c_longlong),
                                           (18446744073709551616, -27, c.c_ulonglong), (2, -7, c.c_ubyte)]:
                execute('SELECT ' + str(number))
                m.check(m.SQLFetch(statement), 3, statement)
                value, indicator = scalar(), c.c_ssize_t()
                assert m.SQLGetData(statement, 1, target, c.byref(value), c.sizeof(value), c.byref(indicator)) == -1
                assert m.diagnostics(3, statement)[0][0] == '22003'
                close()
            for text, target, scalar, expected, status in [
                    ('1.9', 4, c.c_int, 1, 1), ('-1.9', 5, c.c_short, -1, 1),
                    ('42', 4, c.c_int, 42, 0), ('1.99999999999999999999999999999999', 4, c.c_int, 1, 1),
                    ('1.99999999999999999999999999999999', -7, c.c_ubyte, 1, 1),
                    ('9223372036854775807.9', -25, c.c_longlong, 9223372036854775807, 1),
                    ('18446744073709551615', -27, c.c_ulonglong, 18446744073709551615, 0),
                    ('1e2', -17, c.c_ushort, 100, 0), ('1e-1000000000', -18, c.c_uint, 0, 1)]:
                execute("SELECT '" + text + "'")
                m.check(m.SQLFetch(statement), 3, statement)
                value, indicator = scalar(), c.c_ssize_t()
                actual = m.SQLGetData(statement, 1, target, c.byref(value), c.sizeof(value), c.byref(indicator))
                assert actual == status, (text, target, actual, m.diagnostics(3, statement))
                assert value.value == expected
                if status == 1:
                    assert m.diagnostics(3, statement)[0][0] == '01S07'
                close()
            for text, state in [('invalid', '22018'), ('1e30', '22003')]:
                execute("SELECT '" + text + "'")
                m.check(m.SQLFetch(statement), 3, statement)
                value, indicator = c.c_longlong(), c.c_ssize_t()
                assert m.SQLGetData(statement, 1, -25, c.byref(value), 8, c.byref(indicator)) == -1
                assert m.diagnostics(3, statement)[0][0] == state
                close()
            execute("SELECT 'abcdef', 'other', NULL, 42")
            m.check(m.SQLFetch(statement), 3, statement)
            chunk, indicator = c.create_unicode_buffer(3), c.c_ssize_t()
            assert m.SQLGetData(statement, 1, -8, chunk, c.sizeof(chunk), c.byref(indicator)) == 1
            assert chunk.value == 'ab' and m.text(statement, 2) == 'other'
            assert m.text(statement, 1) == 'abcdef'
            assert m.text(statement, 2) == 'other' and m.text(statement, 1) == 'abcdef'
            assert m.text(statement, 3) is None and m.text(statement, 4) == '42' and m.text(statement, 3) is None
            assert m.text(statement, 4) == '42'
            close()
            bound_integer, indicator = c.c_int(), c.c_ssize_t()
            row_status = c.c_ushort()
            m.check(m.SQLSetStmtAttrW(statement, 25, c.cast(c.byref(row_status), c.c_void_p), 0), 3, statement)
            m.check(m.SQLBindCol(statement, 1, 4, c.byref(bound_integer), 4, c.byref(indicator)), 3, statement)
            execute('SELECT 1.9')
            assert m.SQLFetch(statement) == 1 and bound_integer.value == 1
            assert row_status.value == 6, "A warned row must have SQL_ROW_SUCCESS_WITH_INFO"
            assert m.diagnostics(3, statement)[0][0] == '01S07'
            close()
            m.check(m.SQLFreeStmt(statement, 2), 3, statement)
            m.check(m.SQLSetStmtAttrW(statement, 25, None, 0), 3, statement)

            execute('SELECT "Name", "Amount", "Optional" FROM "Customers" ORDER BY "Name"')
            count = c.c_short()
            m.check(m.SQLNumResultCols(statement, c.byref(count)), 3, statement)
            assert count.value == 3
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.text(statement, 1, 4) == "Alicja Żółć"
            amount, indicator = c.c_double(), c.c_ssize_t()
            m.check(m.SQLGetData(statement, 2, 8, c.byref(amount), 8, c.byref(indicator)), 3, statement)
            assert amount.value == 12.5 and indicator.value == 8
            assert m.text(statement, 3) == ""
            end = c.create_unicode_buffer(8)
            assert m.SQLGetData(statement, 3, -8, end, c.sizeof(end), c.byref(indicator)) == 100
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.text(statement, 1) == "Bob" and m.text(statement, 3) is None
            assert m.SQLFetch(statement) == 100
            close()
            assert m.SQLExecDirectW(statement, "SELECT " + ",".join("1" for _ in range(257)), -3) == -1
            assert m.diagnostics(3, statement)[0][0] == "HY001"
            m.check(m.SQLSetStmtAttrW(statement, 1, 1, 0), 3, statement)
            execute('SELECT "Name" FROM "Customers"')
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.SQLFetch(statement) == 100
            close()
            m.check(m.SQLSetStmtAttrW(statement, 1, 0, 0), 3, statement)

            m.check(m.SQLPrepareW(statement, 'SELECT "Name" FROM "Customers" WHERE "Country" = ?', -3), 3, statement)
            parameter, indicator = c.create_unicode_buffer("Poland"), c.c_ssize_t(12)
            m.check(m.SQLBindParameter(statement, 1, 1, -8, -9, 6, 0, parameter, c.sizeof(parameter), c.byref(indicator)), 3, statement)
            m.check(m.SQLExecute(statement), 3, statement)
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.text(statement, 1) == "Alicja Żółć" and m.SQLFetch(statement) == 100
            close()
            m.check(m.SQLFreeStmt(statement, 3), 3, statement)

            for char_type, buffer in [(-8, c.create_unicode_buffer("Poland\0padding")), (1, c.create_string_buffer(b"Poland\0padding"))]:
                for capacity in (0, c.sizeof(buffer)):
                    m.check(m.SQLBindParameter(statement, 1, 1, char_type, -9 if char_type == -8 else 12, 6, 0, buffer, capacity, None), 3, statement)
                    m.check(m.SQLExecute(statement), 3, statement)
                    m.check(m.SQLFetch(statement), 3, statement)
                    assert m.text(statement, 1) == "Alicja Żółć" and m.SQLFetch(statement) == 100
                    close()
            m.check(m.SQLFreeStmt(statement, 3), 3, statement)

            m.check(m.SQLTablesW(statement, None, 0, None, 0, "%", -3, None, 0), 3, statement)
            tables = []
            while m.SQLFetch(statement) == 0:
                tables.append(m.text(statement, 3))
                assert m.text(statement, 4) == "TABLE"
            assert tables == ["A", "Customers", "Z"]
            close()
            m.check(m.SQLColumnsW(statement, None, 0, None, 0, "Customers", -3, "%", -3), 3, statement)
            columns = []
            while m.SQLFetch(statement) == 0:
                columns.append(m.text(statement, 4))
            assert columns == ["Name", "Country", "Amount", "Optional", "Active", "Recorded"]
            close()
            m.check(m.SQLColumnsW(statement, None, 0, None, 0, "%", -3, "%", -3), 3, statement)
            keys = []
            while m.SQLFetch(statement) == 0:
                keys.append((m.text(statement, 3), int(m.text(statement, 17))))
            assert keys == sorted(keys) and len(keys) == 18
            close()

            if iteration == 0:
                m.check(m.SQLSetStmtAttrW(statement, 0, 1, 0), 3, statement)
                execute('SELECT "Name" FROM "Customers"')
                m.check(m.SQLFetch(statement), 3, statement)
                time.sleep(1.2)  # Deliberate caller time, exceeding the requested execution deadline.
                m.check(m.SQLFetch(statement), 3, statement)
                assert m.text(statement, 1) == "Bob"
                close()
                m.check(m.SQLSetStmtAttrW(statement, 0, 30, 0), 3, statement)

            m.check(m.SQLTablesW(statement, None, 0, None, 0, "", 0, None, 0), 3, statement)
            assert m.SQLFetch(statement) == 100, "An empty pattern must not match every table"
            close()

            execute('SELECT "Active", "Recorded" FROM "Customers" WHERE "Country" = \'Poland\'')
            m.check(m.SQLFetch(statement), 3, statement)
            active, indicator = c.c_ubyte(), c.c_ssize_t()
            m.check(m.SQLGetData(statement, 1, -7, c.byref(active), 1, c.byref(indicator)), 3, statement)
            assert active.value == 1
            class Timestamp(c.Structure):
                _fields_ = [("year", c.c_short), ("month", c.c_ushort), ("day", c.c_ushort), ("hour", c.c_ushort),
                            ("minute", c.c_ushort), ("second", c.c_ushort), ("fraction", c.c_uint)]
            date = Timestamp()
            m.check(m.SQLGetData(statement, 2, 93, c.byref(date), c.sizeof(date), c.byref(indicator)), 3, statement)
            assert (date.year, date.month, date.day, date.hour, date.minute, date.second, date.fraction) == (2026, 10, 12, 1, 2, 3, 123456700)
            for target, expected in [(91, (2026, 10, 12)), (92, (1, 2, 3))]:
                assert m.text(statement, 1) == '1'  # Switching columns restarts the timestamp.
                parts = (c.c_ushort * 3)()
                assert m.SQLGetData(statement, 2, target, parts, c.sizeof(parts), c.byref(indicator)) == 1
                assert tuple(parts) == expected and m.diagnostics(3, statement)[0][0] == '01S07'
            close()

            execute('SELECT "Amount" FROM "Customers" ORDER BY "Amount"')
            bound, indicator = c.c_double(), c.c_ssize_t()
            m.check(m.SQLBindCol(statement, 1, 8, c.byref(bound), 8, c.byref(indicator)), 3, statement)
            m.check(m.SQLFetch(statement), 3, statement)
            assert bound.value == 12.5
            m.check(m.SQLFetch(statement), 3, statement)
            assert bound.value == 20
            close()
            m.check(m.SQLFreeStmt(statement, 2), 3, statement)

            for sql in ['DELETE FROM "Customers"', 'ATTACH DATABASE \'test.db\' AS other', 'SELECT 1; DELETE FROM "Customers"']:
                assert m.SQLExecDirectW(statement, sql, -3) == -1
                assert m.diagnostics(3, statement)[0][0] == "HYC00"
            execute('SELECT COUNT(*) FROM "Customers"')
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.text(statement, 1) == "2"
            close()
            # A running native SQLite statement must be interrupted, not merely canceled between rows.
            query = 'SELECT COUNT(*) FROM ' + ','.join('"Customers" t' + str(i) for i in range(30))
            result = []
            independent = m.allocate(2, environment)
            m.check(m.SQLDriverConnectW(independent, None, connection_string, -3, None, 0, None, 0), 2, independent)
            worker = threading.Thread(target=lambda: result.append(m.SQLExecDirectW(statement, query, -3)), daemon=True)
            worker.start()
            time.sleep(0.2)
            probe_result = []
            def probe():
                info = c.c_ushort()
                probe_result.append(m.SQLGetInfoW(independent, 0, c.byref(info), 2, None))
            probe_worker = threading.Thread(target=probe, daemon=True)
            probe_worker.start()
            try:
                probe_worker.join(2)
                assert not probe_worker.is_alive() and probe_result == [0], "Independent connection stalled"
            finally:
                m.check(m.SQLCancel(statement), 3, statement)
                worker.join(5)
                probe_worker.join(5)
                m.check(m.SQLDisconnect(independent), 2, independent)
                m.check(m.SQLFreeHandle(2, independent), 2, independent)
            assert not worker.is_alive() and result == [-1], result
            assert m.diagnostics(3, statement)[0][0] == "HY008"
            close()
            execute('SELECT 1')
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.text(statement, 1) == "1"
            close()
        finally:
            if statement:
                m.check(m.SQLFreeHandle(3, statement), 3, statement)
            if connection:
                if connected:
                    m.check(m.SQLDisconnect(connection), 2, connection)
                m.check(m.SQLFreeHandle(2, connection), 2, connection)
            m.check(m.SQLFreeHandle(1, environment), 1, environment)
    if ado_consumer:
        subprocess.run(["dotnet", str(ado_consumer), connection_string, hive], check=True, timeout=60)
    assert original == hashlib.sha256(path.read_bytes()).digest(), "Source workbook changed"
    print(json.dumps({"windows_odbc_manager": "passed", "unicode_and_chunking": "passed", "nulls_and_numbers": "passed",
                      "parameters_and_bound_columns": "passed", "metadata": "passed", "read_only": "passed",
                      "native_cancellation": "passed", "reconnects": 3, "source_unchanged": True}))


def verify_native_widths(driver):
    # Call driver exports directly: the manager must not mask a caller-buffer overwrite.
    m = Manager(driver)
    environment = m.allocate(1)
    m.check(m.SQLSetEnvAttr(environment, 200, 380, 0), 1, environment)
    version = c.c_int()
    m.check(m.SQLGetEnvAttr(environment, 200, c.byref(version), 0, None), 1, environment)
    assert version.value == 380
    connection = m.allocate(2, environment)
    statement = m.allocate(3, connection)
    descriptor = c.c_void_p()
    canary = (c.c_int * 2)(0, 0x12345678)
    try:
        assert m.SQLSetStmtAttrW(statement, 999, None, 0) == -1
        m.check(m.SQLGetDiagFieldW(3, statement, 1, -1247, canary, 0, None), 3, statement)
        assert canary[0] == -1 and canary[1] == 0x12345678
        # Identifier 3 is SQL_DIAG_ROW_COUNT (SQLLEN), not SQL_DIAG_RETURNCODE (1).
        row_count = (c.c_ssize_t * 2)(0, 0x12345678)
        m.check(m.SQLGetDiagFieldW(3, statement, 0, 3, row_count, 0, None), 3, statement)
        assert row_count[0] == -1 and row_count[1] == 0x12345678
        return_code = (c.c_short * 4)(123, 456, 789, 123)
        assert m.SQLGetDiagFieldW(3, statement, 0, 1, return_code, 0, None) == -1
        assert list(return_code) == [123, 456, 789, 123]
        m.check(m.SQLGetStmtAttrW(statement, 10010, c.byref(descriptor), 0, None), 3, statement)
        length = c.c_int()
        m.check(m.SQLGetDescFieldW(descriptor, 0, 25, canary, 0, c.byref(length)), 4, descriptor)
        assert canary[0] == 0 and canary[1] == 0x12345678 and length.value == 4
        for attribute, expected in [(0, 30), (2, 1), (4, 0), (6, 0), (7, 1), (11, 1)]:
            value = c.c_size_t(-1)
            m.check(m.SQLGetStmtAttrW(statement, attribute, c.byref(value), 0, None), 3, statement)
            assert value.value == expected
    finally:
        try:
            m.check(m.SQLFreeHandle(3, statement), 3, statement)
            if descriptor.value:
                assert m.SQLGetDescFieldW(descriptor, 0, 25, canary, 0, None) == -2
        finally:
            try:
                m.check(m.SQLFreeHandle(2, connection), 2, connection)
            finally:
                m.check(m.SQLFreeHandle(1, environment), 1, environment)
    print("Native ABI: diagnostic/descriptor canaries, native-width attributes, ODBC version and stale handles passed.")


def verify_native_compatibility(driver, path):
    m = Manager(driver)
    environment = m.allocate(1)
    connection = m.allocate(2, environment)
    statement = None
    connected = False
    try:
        m.check(m.SQLDriverConnectW(connection, None, f"DBQ={{{path}}};ReadOnly=1;", -3, None, 0, None, 0), 2, connection)
        connected = True
        statement = m.allocate(3, connection)
        assert m.SQLCloseCursor(statement) == -1
        assert m.diagnostics(3, statement)[0][0] == "24000"
        m.check(m.SQLFreeStmt(statement, 0), 3, statement)
        assert m.SQLSetStmtAttrW(statement, 2, 0, 0) == -1
        m.check(m.SQLSetStmtAttrW(statement, 2, 1, 0), 3, statement)
        m.check(m.SQLSetStmtAttrW(statement, 4, 0, 0), 3, statement)
        assert m.SQLSetStmtAttrW(statement, 4, 1, 0) == -1
        catalog, length = c.create_unicode_buffer(2), c.c_int()
        assert m.SQLGetConnectAttrW(connection, 109, catalog, c.sizeof(catalog), c.byref(length)) == 1
        assert m.diagnostics(2, connection)[0][0] == "01004"
        for kind, handle in [(1, environment), (2, connection)]:
            for completion in (0, 1):
                m.check(m.SQLEndTran(kind, handle, completion), kind, handle)
        for info, expected in [(84, 2), (93, 4), (95, 23)]:
            value = c.c_uint()
            m.check(m.SQLGetInfoW(connection, info, c.byref(value), 4, None), 2, connection)
            assert value.value == expected
        m.check(m.SQLGetTypeInfoW(statement, 8), 3, statement)
        m.check(m.SQLFetch(statement), 3, statement)
        assert m.text(statement, 3) == "53" and m.text(statement, 18) == "2"
        m.check(m.SQLCloseCursor(statement), 3, statement)
        m.check(m.SQLGetTypeInfoW(statement, 0), 3, statement)
        types = []
        while m.SQLFetch(statement) == 0:
            types.append(int(m.text(statement, 2)))
        assert types == [-9, -7, -5, 4, 5, 8, 93]
        m.check(m.SQLCloseCursor(statement), 3, statement)
        parameter = c.create_unicode_buffer("001")
        assert m.SQLBindParameter(statement, 1, 1, -8, 4, 10, 0, parameter, c.sizeof(parameter), None) == -1
        assert m.diagnostics(3, statement)[0][0] == "07006"
        integer = c.c_int(1)
        m.check(m.SQLBindParameter(statement, 1, 1, 99, 4, 10, 0, c.byref(integer), c.sizeof(integer), None), 3, statement)
        m.check(m.SQLPrepareW(statement, 'SELECT ? = 1', -3), 3, statement)
        m.check(m.SQLExecute(statement), 3, statement)
        m.check(m.SQLFetch(statement), 3, statement)
        assert m.text(statement, 1) == "1"
        m.check(m.SQLCloseCursor(statement), 3, statement)
        m.check(m.SQLFreeStmt(statement, 3), 3, statement)
        for catalog_call in (m.SQLTablesW, m.SQLColumnsW):
            m.check(catalog_call(statement, path.name, -3, None, 0, None, 0, None, 0), 3, statement)
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.text(statement, 1) == path.name
        m.check(m.SQLTablesW(statement, "%", -3, "", 0, "", 0, None, 0), 3, statement)
        m.check(m.SQLFetch(statement), 3, statement)
        assert m.text(statement, 1) == path.name and m.SQLFetch(statement) == 100
        m.check(m.SQLFreeStmt(statement, 0), 3, statement)
        # Metadata must not read an application buffer or evaluate a value-sensitive LIMIT or expression.
        m.check(m.SQLBindParameter(statement, 1, 1, 4, 4, 10, 0, c.c_void_p(1), 4, None), 3, statement)
        m.check(m.SQLPrepareW(statement, 'SELECT Name AS same, Amount AS same, abs(-9223372036854775808) AS unevaluated FROM Customers LIMIT ?', -3), 3, statement)
        count = c.c_short()
        m.check(m.SQLNumResultCols(statement, c.byref(count)), 3, statement)
        assert count.value == 3
        for ordinal, expected_name, expected_type in [(1, 'same', -9), (2, 'same', 8), (3, 'unevaluated', 0)]:
            name, sql_type = c.create_unicode_buffer(32), c.c_short()
            m.check(m.SQLDescribeColW(statement, ordinal, name, 32, None, c.byref(sql_type), None, None, None), 3, statement)
            assert name.value == expected_name and sql_type.value == expected_type
        result_descriptor = c.c_void_p()
        m.check(m.SQLGetStmtAttrW(statement, 10012, c.byref(result_descriptor), 0, None), 3, statement)
        m.check(m.SQLGetDescFieldW(result_descriptor, 0, 1001, c.byref(count), 2, None), 4, result_descriptor)
        assert count.value == 3
        assert m.SQLFetch(statement) == -1 and m.diagnostics(3, statement)[0][0] == '24000'
        m.check(m.SQLPrepareW(statement, 'SELECT Name FROM Customers LIMIT ?', -3), 3, statement)
        m.check(m.SQLNumResultCols(statement, c.byref(count)), 3, statement)
        assert count.value == 1
        m.check(m.SQLBindParameter(statement, 1, 1, 4, 4, 10, 0, c.byref(integer), 4, None), 3, statement)
        m.check(m.SQLExecute(statement), 3, statement)
        m.check(m.SQLFetch(statement), 3, statement)
        assert m.text(statement, 1) == 'Alicja Żółć' and m.SQLFetch(statement) == 100
        for date_type, fields in [(91, (2026, 13, 1, 0, 0, 0, 0, 0)), (93, (2026, 1, 1, 24, 0, 0, 0, 0))]:
            date = (c.c_short * 8)(*fields)
            m.check(m.SQLBindParameter(statement, 1, 1, date_type, date_type, 27, 7, date, 16, None), 3, statement)
            m.check(m.SQLPrepareW(statement, 'SELECT ?', -3), 3, statement)
            assert m.SQLExecute(statement) == -1 and m.diagnostics(3, statement)[0][0] == '22008'
        m.check(m.SQLFreeStmt(statement, 3), 3, statement)
        for sql, expected in [("SELECT '1,2'", "22018"), ("SELECT '1e999'", "22003")]:
            for target, scalar in [(7, c.c_float), (8, c.c_double)]:
                m.check(m.SQLExecDirectW(statement, sql, -3), 3, statement)
                m.check(m.SQLFetch(statement), 3, statement)
                value = scalar()
                assert m.SQLGetData(statement, 1, target, c.byref(value), c.sizeof(value), None) == -1
                assert m.diagnostics(3, statement)[0][0] == expected
        for target in (1, -8):
            m.check(m.SQLExecDirectW(statement, "SELECT X'00ABFF'", -3), 3, statement)
            m.check(m.SQLFetch(statement), 3, statement)
            buffer = c.create_string_buffer(4) if target == 1 else c.create_unicode_buffer(4)
            chunks = []
            for expected in ('00', 'AB', 'FF'):
                status = m.SQLGetData(statement, 1, target, buffer, c.sizeof(buffer), None)
                assert status == (0 if expected == 'FF' else 1)
                chunks.append(buffer.value.decode('ascii') if target == 1 else buffer.value)
            assert ''.join(chunks) == '00ABFF'
            m.check(m.SQLExecDirectW(statement, 'SELECT 12345', -3), 3, statement)
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.SQLGetData(statement, 1, target, buffer, 3 if target == 1 else 6, None) == -1
            assert m.diagnostics(3, statement)[0][0] == '22003'
            m.check(m.SQLExecDirectW(statement, 'SELECT "Active" FROM "Customers"', -3), 3, statement)
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.SQLGetData(statement, 1, target, buffer, 1 if target == 1 else 2, None) == -1
            assert m.diagnostics(3, statement)[0][0] == '22003'
            m.check(m.SQLFreeStmt(statement, 0), 3, statement)
            m.check(m.SQLBindCol(statement, 1, target, buffer, 3 if target == 1 else 6, None), 3, statement)
            m.check(m.SQLExecDirectW(statement, 'SELECT 12345', -3), 3, statement)
            assert m.SQLFetch(statement) == -1
            assert m.diagnostics(3, statement)[0][0] == '22003'
            m.check(m.SQLFreeStmt(statement, 2), 3, statement)
        m.check(m.SQLExecDirectW(statement, 'SELECT "name", "amount" FROM "customers"', -3), 3, statement)
        value = c.c_ssize_t()
        for ordinal, field, expected in [(1, 4, 0), (2, 4, 15), (2, 1005, 53), (2, 6, 24), (1, 6, -4), (1, 1013, -4)]:
            m.check(m.SQLColAttributeW(statement, ordinal, field, None, 0, None, c.byref(value)), 3, statement)
            assert value.value == expected
        result_descriptor = c.c_void_p()
        m.check(m.SQLGetStmtAttrW(statement, 10012, c.byref(result_descriptor), 0, None), 3, statement)
        m.check(m.SQLGetDescFieldW(result_descriptor, 1, 1011, None, 0, c.byref(length)), 4, result_descriptor)
        assert length.value == 8
        m.check(m.SQLFetchScroll(statement, 1, 123), 3, statement)
        m.check(m.SQLCloseCursor(statement), 3, statement)
        m.check(m.SQLExecDirectW(statement, "SELECT x'000102030405060708'", -3), 3, statement)
        for field, expected in [(1013, -4), (6, -4), (4, 0)]:
            m.check(m.SQLColAttributeW(statement, 1, field, None, 0, None, c.byref(value)), 3, statement)
            assert value.value == expected
        m.check(m.SQLFetch(statement), 3, statement)
        binary = (c.c_ubyte * 9)()
        size = c.c_ssize_t()
        m.check(m.SQLGetData(statement, 1, -2, binary, 9, c.byref(size)), 3, statement)
        assert size.value == 9 and list(binary) == list(range(9))
        m.check(m.SQLCloseCursor(statement), 3, statement)
        for replace in (m.SQLPrepareW, m.SQLExecDirectW):
            m.check(m.SQLPrepareW(statement, "SELECT 1", -3), 3, statement)
            assert replace(statement, "DELETE FROM Customers", -3) == -1
            assert m.SQLExecute(statement) == -1
            assert m.diagnostics(3, statement)[0][0] == "HY010"
            assert replace(statement, "SELECT ?1", -3) == -1
            parameter_count = c.c_short(-1)
            m.check(m.SQLNumParams(statement, c.byref(parameter_count)), 3, statement)
            assert parameter_count.value == 0
        descriptor = c.c_void_p()
        m.check(m.SQLGetStmtAttrW(statement, 10010, c.byref(descriptor), 0, None), 3, statement)
        m.check(m.SQLDisconnect(connection), 2, connection)
        connected = False
        stale = statement
        statement = None
        assert m.SQLFreeHandle(3, stale) == -2
        assert m.SQLGetDescFieldW(descriptor, 0, 25, c.byref(value), 0, None) == -2
    finally:
        try:
            if statement:
                m.check(m.SQLFreeHandle(3, statement), 3, statement)
            if connected:
                m.check(m.SQLDisconnect(connection), 2, connection)
        finally:
            try:
                m.check(m.SQLFreeHandle(2, connection), 2, connection)
            finally:
                m.check(m.SQLFreeHandle(1, environment), 1, environment)
    print("Native compatibility: metadata, failed prepare, fetch offsets, transactionless completion and disconnect lifetime passed.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--driver", type=Path, required=True)
    parser.add_argument("--ado-consumer", type=Path)
    parser.add_argument("--registered-name", help=argparse.SUPPRESS)
    parser.add_argument("--hive", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.registered_name:
        verify_native_widths(args.driver)
        override = c.WinDLL("advapi32.dll").RegOverridePredefKey
        override.argtypes, override.restype = [c.c_void_p, c.c_void_p], c.c_long
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, args.hive, 0, winreg.KEY_ALL_ACCESS) as key:
            assert override(c.c_void_p(0xFFFFFFFF80000002), c.c_void_p(int(key))) == 0
            try:
                with tempfile.TemporaryDirectory(prefix="OfficeIMO-Odbc-") as directory:
                    path = Path(directory) / "customers;Żółć.xlsx"
                    workbook(path)
                    verify_native_compatibility(args.driver, path)
                    verify(args.registered_name, path, args.ado_consumer, args.hive)
            finally:
                assert override(c.c_void_p(0xFFFFFFFF80000002), None) == 0
    else:
        # Redirect this child's HKLM reads to a private HKCU hive. This exercises
        # Windows' real driver manager without administrator rights or global drivers.
        name = "OfficeIMO verification " + uuid.uuid4().hex
        root = r"SOFTWARE\ODBC\ODBCINST.INI"
        hive = r"Software\OfficeIMO.Odbc.Verification." + uuid.uuid4().hex
        created = []
        try:
            for relative in ["", "SOFTWARE", r"SOFTWARE\ODBC", root, root + r"\ODBC Drivers", root + "\\" + name]:
                path = hive + ("\\" + relative if relative else "")
                with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE):
                    created.append(path)
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, hive + "\\" + root + "\\" + name, 0, winreg.KEY_WRITE) as key:
                winreg.SetValueEx(key, "Driver", 0, winreg.REG_SZ, str(args.driver.resolve(strict=True)))
                winreg.SetValueEx(key, "APILevel", 0, winreg.REG_SZ, "0")
                winreg.SetValueEx(key, "FileUsage", 0, winreg.REG_SZ, "2")
                winreg.SetValueEx(key, "ConnectFunctions", 0, winreg.REG_SZ, "NYN")
                winreg.SetValueEx(key, "DriverODBCVer", 0, winreg.REG_SZ, "03.00")
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, hive + "\\" + root + r"\ODBC Drivers", 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, name, 0, winreg.REG_SZ, "Installed")
            command = [sys.executable, __file__, "--driver", str(args.driver), "--registered-name", name, "--hive", hive]
            if args.ado_consumer:
                command += ["--ado-consumer", str(args.ado_consumer)]
            subprocess.run(command, check=True, timeout=90)
        finally:
            for path in reversed(created):
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, path)
