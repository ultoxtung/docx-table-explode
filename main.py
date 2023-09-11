import argparse

from file_handler import FileHandler
from table_explode import explode_all_tables


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('path', help='Path to the file that need to be processed')

    parser.add_argument('--with-column-title', action='store_true',
                        help='Use this flag if table has first row used as column title')
    parser.add_argument('--with-row-title', action='store_true',
                        help='Use this flag if table has first column used as row title')
    parser.add_argument('--column-first', action='store_true',
                        help='Use this flag if you want to explode by each column, not row')
    parser.add_argument('--ignore-column', action='store',
                        help='Ignore these columns, separated by comma')
    parser.add_argument('--ignore-row', action='store',
                        help='Ignore these rows, separated by comma')
    parser.add_argument('--limit', action='store', default=0, type=int,
                        help='Only process the first n table(s)')
    args = parser.parse_args()

    file_handler = FileHandler(args.path)
    content_tree = file_handler.read_docx_content()

    explode_all_tables(
        content_tree,
        with_column_title=args.with_column_title,
        with_row_title=args.with_row_title,
        column_first=args.column_first,
        ignore_cols=args.ignore_column,
        ignore_rows=args.ignore_row,
        limit=args.limit,
    )

    new_filepath = args.path.replace('.docx', '_new.docx')
    file_handler.export_docx(content_tree, new_filepath)


if __name__ == "__main__":
    main()
