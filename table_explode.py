from constants import NSMAP


def _get_tables(element_tree):
    return element_tree.findall('.//w:tbl', namespaces=NSMAP)


def _remove_table(table):
    for row in table.xpath('.//w:tr', namespaces=NSMAP):
        # print('row', row)
        for cell in row.xpath('.//w:tc', namespaces=NSMAP):
            # print('cell', cell)
            for element in cell.xpath('.//w:p | .//w:tbl', namespaces=NSMAP):
                # print('element', element)
                table.addprevious(element)
    table.getparent().remove(table)


def explode_all_tables(content_tree):
    tables = _get_tables(content_tree)
    for table in tables:
        _remove_table(table)
