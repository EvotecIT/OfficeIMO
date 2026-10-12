"""Independent checks for the owned two-level decimal Pages creation profile."""


def verify_decimal_lists(graph, decoded, body, text):
    selected = {entry.character_index: entry.object.identifier for entry in body.table_list_style.entries}
    styles = {identifier: message for identifier, message in decoded.items() if graph[identifier][0] == 2023}
    numbered = {identifier: style for identifier, style in styles.items() if any(style.label_types)}
    if not numbered:
        return []
    for style in numbered.values():
        assert list(style.label_types) == [3, 3] + [0] * 7
        assert list(style.number_types) == [0] * 9 and list(style.tiered_numbers) == [False] * 9
        assert list(style.text_indents) == [1.5, 1.5] + [0] * 7
        assert list(style.indents) == [0, 18] + [0] * 7
        assert len(style.geometries) == 9 and all(
            geometry.scale == 1 and geometry.baseline_offset == 0 and geometry.scale_with_text
            for geometry in style.geometries)
        assert not style.font_name_null and style.font_name
    data = {entry.character_index: entry for entry in body.table_para_data.entries}
    starts = {entry.character_index: entry for entry in body.table_para_starts.entries}
    paragraph_styles = {entry.character_index: decoded[entry.object.identifier]
                        for entry in body.table_para_style.entries if entry.HasField('object')}
    result, offset, scope, counters = [], 0, None, {}
    for raw in text.splitlines(keepends=True):
        assert offset in selected and offset in starts and offset in data
        identifier = selected[offset]
        assert paragraph_styles[offset].para_properties.list_style.identifier == identifier
        if identifier in numbered:
            level, start = data[offset].first, starts[offset].first
            assert level in (0, 1) and starts[offset].second == 0
            assert data[offset].second == (1 if start else 0)
            if identifier != scope:
                counters, scope = {}, identifier
                assert start > 0
            number = start if start else counters.get(level, 0) + 1
            assert 0 < number <= 2147483647
            if level == 0:
                counters.pop(1, None)
            elif not start:
                assert 0 in counters or 1 in counters
            counters[level] = number
            result.append({'offset': offset, 'styleId': identifier, 'level': level,
                           'number': number, 'start': start, 'text': raw.removesuffix('\n'),
                           'font': numbered[identifier].font_name})
        else:
            assert not any(styles[identifier].label_types)
            assert data[offset].first == data[offset].second == starts[offset].first == starts[offset].second == 0
        offset += len(raw.encode('utf-16-le')) // 2
    assert offset in selected and not any(styles[selected[offset]].label_types)
    assert data[offset].first == data[offset].second == starts[offset].first == starts[offset].second == 0
    assert set(selected) == set(data) == set(starts) == set(paragraph_styles) | {offset}
    return result
