"""Independently decode OfficeIMO's native Numbers scalar-table creation profile.

Opt-in validation only: pinned protobuf schemas and numbers-parser 4.19.0.
No OfficeIMO assemblies or native application calls are used.
"""
import argparse
import hashlib
import json
import plistlib
import uuid
from importlib.metadata import version
from pathlib import Path
from zipfile import ZipFile
from google.protobuf import descriptor_pb2, descriptor_pool, message_factory
from iwork_protobuf_evidence import fields, creation_records, creation_references
from iwork_table_creation_evidence import verify_tables

SCHEMAS = {1: 'TN.DocumentArchive', 2: 'TN.SheetArchive', 12009: 'TN.ThemeArchive',
    11006: 'TSP.PackageMetadata', 401: 'TSS.StylesheetArchive', 205: 'TSK.TreeNode',
    4000: 'TSCE.CalculationEngineArchive', 4003: 'TSCE.NamedReferenceManagerArchive',
    4004: 'TSCE.TrackedReferenceStoreArchive', 4008: 'TSCE.FormulaOwnerDependenciesArchive',
    222: 'TSK.CustomFormatListArchive',
    601: 'TSA.FunctionBrowserStateArchive', 12050: 'TN.SheetStyleArchive',
    2021: 'TSWP.CharacterStyleArchive', 2022: 'TSWP.ParagraphStyleArchive',
    2023: 'TSWP.ListStyleArchive', 2025: 'TSWP.ShapeStyleArchive', 2050: 'TSWP.TextStylePresetArchive',
    2001: 'TSWP.StorageArchive', 3016: 'TSD.MediaStyleArchive',
    6000: 'TST.TableInfoArchive', 6001: 'TST.TableModelArchive', 6002: 'TST.Tile',
    6003: 'TST.TableStyleArchive', 6004: 'TST.CellStyleArchive', 6005: 'TST.TableDataList',
    6006: 'TST.HeaderStorageBucket', 6008: 'TST.TableStylePresetArchive', 6247: 'TST.TableStyleNetworkArchive'}


def verify(path, descriptor, require_calculation_owners=False):
    pool = descriptor_pool.DescriptorPool()
    descriptor_bytes = descriptor.read_bytes()
    for file in descriptor_pb2.FileDescriptorSet.FromString(descriptor_bytes).file:
        pool.Add(file)
    def decode(name, payload):
        message = message_factory.GetMessageClass(pool.FindMessageTypeByName(name))()
        message.ParseFromString(payload)
        assert not message.FindInitializationErrors(), (name, message.FindInitializationErrors())
        return message
    with ZipFile(path) as package:
        graph = creation_records(package)
        assert not ({71,72} & graph.keys())
        decoded, objects = {}, []
        for identifier, (kind, payload, header) in graph.items():
            message = decode(SCHEMAS[kind], payload)
            decoded[identifier] = message
            refs = creation_references(message)
            if kind == 1:
                # This pinned schema retains the legacy TST custom-format slot (7).
                # Native Numbers uses TSA slot 12 for the modern TSK custom-format list.
                tsa = fields(fields(payload)[8][0])
                for value in tsa.get(12, []): refs.update(creation_references(decode('TSP.Reference', value)))
            if kind == 12009:
                theme = fields(fields(payload)[1][0])
                for field, schema in [(100,'TSD.ThemePresetsArchive'),(110,'TSWP.ThemePresetsArchive'),(200,'TST.ThemePresetsArchive')]:
                    for value in theme.get(field,[]): refs.update(creation_references(decode(schema,value)))
            assert refs == set(header.message_infos[0].object_references), (identifier, refs, set(header.message_infos[0].object_references))
            assert refs.issubset(graph.keys()), (identifier,refs-graph.keys())
            assert not header.message_infos[0].data_references
            objects.append({'id':identifier,'type':kind,'schema':SCHEMAS[kind],'references':sorted(refs)})
        root = decoded[1]
        assert graph[1][0] == 1 and len(root.sheets)>0
        sidebar=decoded[root.sidebar_order.identifier]
        assert len(sidebar.children)==len(root.sheets)
        tables=verify_tables(graph,decoded)
        calculation=decoded[getattr(root,'super').calculation_engine.identifier]
        tracker=calculation.dependency_tracker
        table_infos={identifier:message for identifier,message in decoded.items() if graph[identifier][0]==6000}
        assert tracker.number_of_formulas==0
        if require_calculation_owners or tracker.formula_owner_dependencies:
            assert len(tracker.formula_owner_dependencies)==len(table_infos)
        owner_map={entry.internal_owner_id:entry.owner_id for entry in tracker.owner_id_map.map_entry}
        assert len(owner_map)==len(tracker.formula_owner_dependencies)
        calculation_owners=[]
        seen_tables=set()
        for ref in tracker.formula_owner_dependencies:
            owner=decoded[ref.identifier]
            info_id=owner.formula_owner.identifier
            assert info_id in table_infos and info_id not in seen_tables
            seen_tables.add(info_id)
            info=table_infos[info_id]; model=decoded[info.tableModel.identifier]
            assert info.formula_coord_space==0 and owner.owner_kind==1
            identity=uuid.UUID(model.table_id).bytes
            assert owner.formula_owner_uid.lower==int.from_bytes(identity[:8],'little')
            assert owner.formula_owner_uid.upper==int.from_bytes(identity[8:],'little')
            mapped=owner_map[owner.internal_formula_owner_id]
            assert [mapped.uuid_w0,mapped.uuid_w1,mapped.uuid_w2,mapped.uuid_w3]==[
                int.from_bytes(identity[offset:offset+4],'little') for offset in range(0,16,4)]
            for spanning in (owner.spanning_column_dependencies,owner.spanning_row_dependencies):
                for area in (spanning.total_range_for_table,spanning.body_range_for_table):
                    assert (area.top_left_column,area.top_left_row,area.bottom_right_column,area.bottom_right_row)==(
                        0,0,model.number_of_columns-1,model.number_of_rows-1)
            calculation_owners.append({'tableInfo':info_id,'tableModel':info.tableModel.identifier,
                'uuid':model.table_id,'ownerIndex':owner.internal_formula_owner_id,
                'rows':model.number_of_rows,'columns':model.number_of_columns,'coordinateSpace':info.formula_coord_space})
        sheets=[]
        for index, ref in enumerate(root.sheets):
            sheet=decoded[ref.identifier]
            node=decoded[sidebar.children[index].identifier]
            assert node.object.identifier == ref.identifier
            assert [child.object.identifier for child in (decoded[x.identifier] for x in node.children)] == [x.identifier for x in sheet.drawable_infos]
            assert len(sheet.headers)==len(sheet.footers)==3
            for header in list(sheet.headers)+list(sheet.footers):
                storage=decoded[header.identifier]
                assert storage.kind==1 and ''.join(storage.text)=='\n' and storage.in_document
            table_ids=[decoded[x.identifier].tableModel.identifier for x in sheet.drawable_infos]
            sheets.append({'name':sheet.name,'table_ids':table_ids,'hidden_header_footer_storages_complete':True})
        names=package.namelist(); assert names==sorted(names) and len(names)==5
        assert package.testzip() is None
        assert all(item.compress_type==0 and item.date_time==(1980,1,1,0,0,0) for item in package.infolist())
        properties=plistlib.loads(package.read('Metadata/Properties.plist'))
        assert properties['documentUUID']==package.read('Metadata/DocumentIdentifier').decode('ascii')
        plistlib.loads(package.read('Metadata/BuildVersionHistory.plist'))
    return {'source':str(path),'sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'bytes':path.stat().st_size,
        'descriptor_sha256':hashlib.sha256(descriptor_bytes).hexdigest(),'numbers_parser_version':version('numbers-parser'),
        'required_fields_complete':True,'references_complete':True,'deterministic_zip_profile':True,
        'records':len(objects),'objects':objects,'sheets':sheets,'native_tables':tables,
        'table_calculation_owners':calculation_owners,
        'table_calculation_owners_complete':len(calculation_owners)==len(table_infos)}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('package',type=Path);parser.add_argument('--descriptor',required=True,type=Path);parser.add_argument('--out',required=True,type=Path)
    parser.add_argument('--require-calculation-owners',action='store_true',help='Require the current formula-entry-ready table graph; historical scalar-only references can omit it.')
    args=parser.parse_args();assert version('numbers-parser')=='4.19.0'
    result=verify(args.package,args.descriptor,args.require_calculation_owners);args.out.write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n')
    print({k:v for k,v in result.items() if k not in ('objects','sheets','native_tables')})
