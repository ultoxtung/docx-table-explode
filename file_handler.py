from lxml import etree
import os
import shutil
import tempfile
import zipfile


class FileHandler:
    def __init__(self, filepath):
        self._file = open(filepath, 'rb')
        self._zip_file = zipfile.ZipFile(self._file)

    def read_docx_content(self):
        content = self._zip_file.read('word/document.xml')
        return etree.fromstring(content)

    def export_docx(self, content, output_filename):
        """ Create a temp directory, expand the original docx zip.
            Write the modified xml to word/document.xml
            Zip it up as the new docx
        """

        tmp_dir = tempfile.mkdtemp()
        self._zip_file.extractall(tmp_dir)

        # Get a list of all the files in the original docx zipfile
        filenames = self._zip_file.namelist()

        with open(os.path.join(tmp_dir, 'word/document.xml'), 'wb') as f:
            xmlstr = etree.tostring(content, pretty_print=True)
            f.write(xmlstr)

        # Now, create the new zip file and add all the filex into the archive
        zip_copy_filename = output_filename
        with zipfile.ZipFile(zip_copy_filename, "w") as docx:
            for filename in filenames:
                docx.write(os.path.join(tmp_dir, filename), filename)

        # Clean up the temp dir
        shutil.rmtree(tmp_dir)
        self._file.close()
