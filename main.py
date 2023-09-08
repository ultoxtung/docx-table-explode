import argparse

from file_handler import FileHandler
from table_explode import explode_all_tables


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('path', help='Path to the file that need to be processed')
    args = parser.parse_args()

    file_handler = FileHandler(args.path)
    content_tree = file_handler.read_docx_content()
    explode_all_tables(content_tree)
    new_filepath = args.path.replace('.docx', '_new.docx')
    file_handler.export_docx(content_tree, new_filepath)


if __name__ == "__main__":
    main()
