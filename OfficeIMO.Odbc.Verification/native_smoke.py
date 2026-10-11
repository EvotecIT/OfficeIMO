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
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, xml in parts.items():
            archive.writestr(name, xml.encode("utf-8"))


class Manager:
    def __init__(self):
        self.dll = c.WinDLL("odbc32.dll")
        h, s, u, i, p, n, w = c.c_void_p, c.c_short, c.c_ushort, c.c_int, c.c_void_p, c.c_ssize_t, c.c_wchar_p
        signatures = {
            "SQLAllocHandle": [s, h, p], "SQLFreeHandle": [s, h],
            "SQLSetEnvAttr": [h, i, h, i], "SQLDriverConnectW": [h, h, w, s, p, s, p, u],
            "SQLGetDiagRecW": [s, h, s, p, p, p, s, p], "SQLDisconnect": [h],
            "SQLExecDirectW": [h, w, i], "SQLFetch": [h], "SQLGetData": [h, u, s, p, n, p],
            "SQLNumResultCols": [h, p], "SQLCloseCursor": [h], "SQLPrepareW": [h, w, i],
            "SQLBindParameter": [h, u, s, s, s, c.c_size_t, s, p, n, p], "SQLExecute": [h],
            "SQLTablesW": [h, w, s, w, s, w, s, w, s], "SQLColumnsW": [h, w, s, w, s, w, s, w, s],
            "SQLBindCol": [h, u, s, p, n, p], "SQLFreeStmt": [h, u], "SQLCancel": [h],
            "SQLGetInfoW": [h, u, p, s, p], "SQLGetFunctions": [h, u, p],
            "SQLSetStmtAttrW": [h, i, h, i],
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
            m.check(m.SQLDriverConnectW(connection, None, connection_string, -3, output, 2048, c.byref(length), 0), 2, connection)
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
            m.check(m.SQLBindCol(statement, 1, 4, c.byref(bound_integer), 4, c.byref(indicator)), 3, statement)
            execute('SELECT 1.9')
            assert m.SQLFetch(statement) == 1 and bound_integer.value == 1
            assert m.diagnostics(3, statement)[0][0] == '01S07'
            close()
            m.check(m.SQLFreeStmt(statement, 2), 3, statement)

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

            m.check(m.SQLTablesW(statement, None, 0, None, 0, "%", -3, None, 0), 3, statement)
            m.check(m.SQLFetch(statement), 3, statement)
            assert m.text(statement, 3) == "Customers" and m.text(statement, 4) == "TABLE"
            assert m.SQLFetch(statement) == 100
            close()
            m.check(m.SQLColumnsW(statement, None, 0, None, 0, "Customers", -3, "%", -3), 3, statement)
            columns = []
            while m.SQLFetch(statement) == 0:
                columns.append(m.text(statement, 4))
            assert columns == ["Name", "Country", "Amount", "Optional", "Active", "Recorded"]
            close()

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
            worker = threading.Thread(target=lambda: result.append(m.SQLExecDirectW(statement, query, -3)), daemon=True)
            worker.start()
            time.sleep(0.2)
            m.check(m.SQLCancel(statement), 3, statement)
            worker.join(5)
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


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--driver", type=Path, required=True)
    parser.add_argument("--ado-consumer", type=Path)
    parser.add_argument("--registered-name", help=argparse.SUPPRESS)
    parser.add_argument("--hive", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.registered_name:
        override = c.WinDLL("advapi32.dll").RegOverridePredefKey
        override.argtypes, override.restype = [c.c_void_p, c.c_void_p], c.c_long
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, args.hive, 0, winreg.KEY_ALL_ACCESS) as key:
            assert override(c.c_void_p(0xFFFFFFFF80000002), c.c_void_p(int(key))) == 0
            try:
                with tempfile.TemporaryDirectory(prefix="OfficeIMO-Odbc-") as directory:
                    path = Path(directory) / "customers;Żółć.xlsx"
                    workbook(path)
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
                winreg.SetValueEx(key, "APILevel", 0, winreg.REG_SZ, "1")
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
