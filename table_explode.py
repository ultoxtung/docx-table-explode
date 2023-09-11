from copy import deepcopy

from constants import NSMAP
from helper import get_first_sub_element


class Table():
    def _read_table(self, table):
        self.table = []
        row_list = table.xpath('./w:tr', namespaces=NSMAP)
        if not len(row_list):
            return

        for row in row_list:
            cells = []
            for cell in row.xpath('./w:tc', namespaces=NSMAP):
                cells.append(cell)

                grid_span = get_first_sub_element(cell, './w:tcPr/w:gridSpan/@w:val')
                cell_span = int(str(grid_span)) if grid_span is not None else 1

                for _ in range(cell_span - 1):
                    cells.append(None)

            self.table.append(cells)

        self.column_count = len(self.table[0])
        self.row_count = len(self.table)

    def __init__(self, table):
        self.old_table_element = table
        self._read_table(table)

    def _append_cell(self, cell):
        if cell is None:
            return

        empty_element = []
        for element in cell.xpath('./*[not(self::w:tcPr)]', namespaces=NSMAP):
            if len(element.getchildren()) == 1 and bool(element.xpath('./w:pPr', namespaces=NSMAP)):
                empty_element.append(element)
                continue
            self.old_table_element.addprevious(deepcopy(element))

        if len(empty_element):
            for element in empty_element:
                element.getparent().remove(element)

    def remove_table(self, with_column_title, with_row_title, column_first, ignore_cols, ignore_rows):
        if column_first:
            start_col_idx = 0
            if with_row_title:
                start_col_idx = 1
            for col_idx in range(start_col_idx, self.column_count):
                if col_idx in ignore_cols:
                    continue

                start_row_idx = 0
                if with_column_title:
                    self._append_cell(self.table[0][col_idx])
                    start_row_idx = 1

                for row_idx in range(start_row_idx, self.row_count):
                    if row_idx in ignore_rows:
                        continue
                    if with_row_title:
                        self._append_cell(self.table[row_idx][0])
                    self._append_cell(self.table[row_idx][col_idx])
        else:
            start_row_idx = 0
            if with_column_title:
                start_row_idx = 1
            for row_idx in range(start_row_idx, self.row_count):
                if row_idx in ignore_rows:
                    continue

                start_col_idx = 0
                if with_row_title:
                    self._append_cell(self.table[row_idx][0])
                    start_col_idx = 1

                for col_idx in range(start_col_idx, self.column_count):
                    if col_idx in ignore_cols:
                        continue
                    if with_column_title:
                        self._append_cell(self.table[0][col_idx])
                    self._append_cell(self.table[row_idx][col_idx])

        self.old_table_element.getparent().remove(self.old_table_element)


def _get_tables(element_tree):
    return [Table(table) for table in element_tree.xpath('.//w:tbl[not(ancestor::w:tbl)]', namespaces=NSMAP)]


def explode_all_tables(content_tree, with_column_title, with_row_title, column_first, ignore_cols, ignore_rows, limit):
    tables = _get_tables(content_tree)

    processed_count = 0
    for table in tables:
        table.remove_table(
            with_column_title=with_column_title,
            with_row_title=with_row_title,
            column_first=column_first,
            ignore_cols=[int(col_idx) - 1 for col_idx in ignore_cols.split(',')] if ignore_cols is not None else [],
            ignore_rows=[int(row_idx) - 1 for row_idx in ignore_rows.split(',')] if ignore_rows is not None else [],
        )
        processed_count += 1
        if limit > 0 and processed_count >= limit:
            break
