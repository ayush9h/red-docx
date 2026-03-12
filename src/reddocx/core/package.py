import zipfile
from io import BytesIO
from typing import Union

from lxml import etree


class DocxPackage:
    def __init__(self, input_type: Union[str, bytes, BytesIO]):
        if isinstance(input_type, str):
            with open(input_type, "rb") as f:
                self._raw = f.read()

        elif isinstance(input_type, bytes):
            self._raw = input_type

        elif isinstance(input_type, BytesIO):
            self._raw = input_type.getvalue()

        else:
            raise TypeError("Input Type must be path | bytes | BytesIO")

        self.zip = zipfile.ZipFile(BytesIO(self._raw), "r")

    def read_xml(self, part: str) -> etree._Element:
        xml_bytes = self.zip.read(part)
        return etree.fromstring(xml_bytes)

    def list_parts(self):
        return self.zip.namelist()
