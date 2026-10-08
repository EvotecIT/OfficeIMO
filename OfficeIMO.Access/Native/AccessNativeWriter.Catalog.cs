using System.Text;
namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        private static object?[][] PermissionRows(IEnumerable<Table> userTables, int relationships) {
            object?[] Permission(int id, int sid, int rights, bool inherited = false) => new object?[] { id, new byte[] { (byte)sid, (byte)(sid >> 8) }, rights, inherited };
            List<object?[]> rows = new List<object?[]> {
                // Plain built-in IDs observed by decoding two independently produced synthetic databases.
                Permission(2,0x0103,393216),Permission(3,0x0103,393216),Permission(4,0x0103,393216),Permission(5,0x0103,917504),
                Permission(0x0f000001,0x0402,983294,true),Permission(0x0f000001,0x0103,393217),
                Permission(0x0f000003,0x0402,983294,true),Permission(0x0f000003,0x0103,393217),
                Permission(0x0f000002,0x0103,393216),Permission(0x10000000,0x0103,393230),Permission(0x10000000,0x0102,14),
                Permission(4,0x0102,20),Permission(5,0x0102,20),Permission(2,0x0102,20),
                Permission(0x0f000001,0x0102,1048319,true),Permission(0x0f000003,0x0102,1048575,true),
            };
            foreach (Table table in userTables) { rows.Add(Permission(table.DefinitionPage,0x0103,983294)); rows.Add(Permission(table.DefinitionPage,0x0102,1048319)); }
            for (int i=0;i<relationships;i++) { rows.Add(Permission(int.MinValue+i,0x0103,983294)); rows.Add(Permission(int.MinValue+i,0x0102,1048575)); }
            return rows.ToArray();
        }
        private void CreateSystemTables(AccessDocument document, List<Table> users, object?[][] relationships) {
            Column[] catalogColumns = new[] {
                    new Column("Id", 4, 4), new Column("ParentId", 4, 4), new Column("Name", 10, 510, true),
                    new Column("Type", 3, 2), new Column("DateCreate", 8, 8), new Column("DateUpdate", 8, 8),
                    new Column("Owner", 9, 255, true) { IsSystemSid = true }, new Column("Flags", 4, 4),
                    new Column("Database", 12, 0, true), new Column("Connect", 12, 0, true), new Column("ForeignName", 10, 510, true),
                    new Column("RmtInfoShort", 9, 255, true), new Column("RmtInfoLong", 11, 0, true), new Column("Lv", 11, 0, true),
                    new Column("LvProp", 11, 0, true), new Column("LvModule", 11, 0, true), new Column("LvExtra", 11, 0, true)
                };
                object?[] CatalogRow(int id, int parent, string name, short type) =>
                    new object?[] { id, parent, name, type, document.CreatedAt.ToOADate(), document.CreatedAt.ToOADate(), new byte[] { 3, 1 }, 0, null, null, null, null, null, null, null, null, null };
            List<object?[]> catalogRows = new List<object?[]> {
                CatalogRow(0x0f000001,0x0f000000,"Tables",3), CatalogRow(0x0f000002,0x0f000000,"Databases",3),
                CatalogRow(0x0f000003,0x0f000000,"Relationships",3), CatalogRow(0x10000000,0x0f000002,"MSysDb",2),
                CatalogRow(2,0x0f000001,"MSysObjects",1), CatalogRow(3,0x0f000001,"MSysACEs",1),
                CatalogRow(4,0x0f000001,"MSysQueries",1), CatalogRow(5,0x0f000001,"MSysRelationships",1)
            };
            for (int i=0;i<catalogRows.Count;i++) catalogRows[i][7]=int.MinValue;
            catalogRows[3][14] = Properties(document.Properties);
            foreach (Table table in users) {
                object?[] row = CatalogRow(table.DefinitionPage,0x0f000001,table.Name,1);
                row[14] = TextColumnProperties(table); catalogRows.Add(row);
            }
            for(int i=0;i<relationships.Length;i++) catalogRows.Add(CatalogRow(int.MinValue+i,0x0f000003,(string)relationships[i][0]!,8));
            AddSystem(new Table(2,"MSysObjects",catalogColumns,catalogRows.ToArray()),new[] { new Index("ParentIdName",new[]{1,2},129),new Index("Id",new[]{0},129,1) });
            AddSystem(new Table(3, "MSysACEs", new[] { new Column("ObjectId",4,4), new Column("SID",9,510,true) { IsSystemSid = true }, new Column("ACM",4,4), new Column("FInheritable",1,1) }, PermissionRows(users, relationships.Length)), new[] {new Index("ObjectId",new[]{0},136)});
            AddSystem(new Table(4, "MSysQueries", new[] { new Column("ObjectId",4,4), new Column("Attribute",2,1), new Column("Order",9,510,true), new Column("Name1",10,510,true), new Column("Name2",10,510,true), new Column("Expression",12,0,true), new Column("Flag",3,2), new Column("LvExtra",4,4) }, Array.Empty<object?[]>()), new[] {new Index("ObjectIdAttribute",new[]{0,1,2},129,1)});
            AddSystem(new Table(5, "MSysRelationships", new[] { new Column("szRelationship",10,510,true), new Column("grbit",4,4), new Column("ccolumn",4,4), new Column("icolumn",4,4), new Column("szObject",10,510,true), new Column("szColumn",10,510,true), new Column("szReferencedObject",10,510,true), new Column("szReferencedColumn",10,510,true) }, relationships), new[] {new Index("szRelationship",new[]{0},130),new Index("szObject",new[]{4},130),new Index("szReferencedObject",new[]{6},130)});
            byte[] mask=SidMask(_pages[0]);
            foreach(Table table in _tables) foreach(object?[] row in table.Rows) for(int c=0;c<table.Columns.Length;c++) {
                if(table.Columns[c].IsSystemSid && row[c] is byte[] sid)
                    row[c]=sid.Select((value,i)=>(byte)(value^mask[i])).ToArray();
            }
        }
        private void AddSystem(Table table, Index[] indexes) {
            table.MapPage=Allocate();
            foreach(Index index in indexes) { index.RootPage=Allocate(); table.Indexes.Add(index); }
            _tables.Add(table);
        }
    }
}
