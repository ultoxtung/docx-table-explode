import unittest

from lxml import etree


from table_explode import explode_all_tables
from file_handler import FileHandler


class TestDocxTableExplode(unittest.TestCase):
    class _TestParams():
        def __init__(self, **kwargs):
            self.with_column_title = kwargs.get('with_column_title', False)
            self.with_row_title = kwargs.get('with_row_title', False)
            self.column_first = kwargs.get('column_first', False)
            self.ignore_cols = kwargs.get('ignore_cols', '')
            self.ignore_rows = kwargs.get('ignore_rows', '')
            self.limit = kwargs.get('limit', 0)

    def _assert_success(self, test_params, expected_file_path):
        file_handler = FileHandler('./testdata/test_input.docx')
        actual_content_tree = file_handler.read_docx_content()
        explode_all_tables(
            actual_content_tree,
            with_column_title=test_params.with_column_title,
            with_row_title=test_params.with_row_title,
            column_first=test_params.column_first,
            ignore_cols=test_params.ignore_cols,
            ignore_rows=test_params.ignore_rows,
            limit=test_params.limit,
        )
        file_handler.close()

        file_handler = FileHandler(expected_file_path)
        expected_content_tree = file_handler.read_docx_content()
        file_handler.close()

        self.assertEqual(
            etree.tostring(actual_content_tree, with_tail=False, pretty_print=True),
            etree.tostring(expected_content_tree, with_tail=False, pretty_print=True),
        )

    def test_with_column_title(self):
        self._assert_success(
            test_params=self._TestParams(with_column_title=True),
            expected_file_path='./testdata/with_column_title_output.docx',
        )

    def test_with_row_title(self):
        self._assert_success(
            test_params=self._TestParams(with_row_title=True),
            expected_file_path='./testdata/with_row_title_output.docx',
        )

    def test_column_first(self):
        self._assert_success(
            test_params=self._TestParams(with_column_title=True, with_row_title=True, column_first=True),
            expected_file_path='./testdata/column_first_output.docx',
        )

    def test_ignore_cols(self):
        self._assert_success(
            test_params=self._TestParams(ignore_cols='1,3'),
            expected_file_path='./testdata/ignore_column_output.docx',
        )

    def test_empty_ignore_cols(self):
        self._assert_success(
            test_params=self._TestParams(with_column_title=True, ignore_cols=''),
            expected_file_path='./testdata/with_column_title_output.docx',
        )

    def test_ignore_rows(self):
        self._assert_success(
            test_params=self._TestParams(ignore_rows='2,3'),
            expected_file_path='./testdata/ignore_row_output.docx',
        )

    def test_empty_ignore_rows(self):
        self._assert_success(
            test_params=self._TestParams(with_column_title=True, ignore_rows=''),
            expected_file_path='./testdata/with_column_title_output.docx',
        )

    def test_limit(self):
        self._assert_success(
            test_params=self._TestParams(limit=1),
            expected_file_path='./testdata/limit_output.docx',
        )

    def test_limit_less_than_zero(self):
        self._assert_success(
            test_params=self._TestParams(with_column_title=True, limit=-2),
            expected_file_path='./testdata/with_column_title_output.docx',
        )


if __name__ == '__main__':
    unittest.main()
