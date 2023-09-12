import argparse
import os

from file_handler import FileHandler
from table_explode import explode_all_tables


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('path', help='Path to the file or folder that need to be processed')

    parser.add_argument('--with-column-title', action='store_true',
                        help='Use this flag if table has first row used as column title')
    parser.add_argument('--with-row-title', action='store_true',
                        help='Use this flag if table has first column used as row title')
    parser.add_argument('--column-first', action='store_true',
                        help='Use this flag for exploding by column first then row, default is row before column')
    parser.add_argument('--ignore-column', action='store', default='',
                        help='Ignore these columns, separated by comma (count from 1)')
    parser.add_argument('--ignore-row', action='store', default='',
                        help='Ignore these rows, separated by comma (count from 1)')
    parser.add_argument('--limit', action='store', default=0, type=int,
                        help='Only process the first n table(s)')
    parser.add_argument('--remove-old-file', action='store_true',
                        help='User this flag if you want to delete the old file after processing')
    args = parser.parse_args()

    paths = []
    if os.path.isdir(args.path):
        for (root, _, files) in os.walk(args.path):
            paths.extend([os.path.join(root, file) for file in files if file.endswith('.docx')])
    else:
        paths = [args.path]

    for path in paths:
        file_handler = FileHandler(path)
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

        new_filepath = path.replace('.docx', '_new.docx')
        file_handler.export_docx(content_tree, new_filepath)

        if args.remove_old_file:
            os.remove(path)
            os.rename(new_filepath, path)


if __name__ == "__main__":
    main()
