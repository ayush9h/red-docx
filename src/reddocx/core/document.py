import os
import shutil
import tempfile
import zipfile
from io import BytesIO
from typing import List, Union, cast

from lxml import etree

from ..xml.namespace import NS
from .package import DocxPackage


class DocxDocument:
    def __init__(self, source: Union[str, bytes, BytesIO]):
        self.pkg = DocxPackage(source)
        self.document = self.pkg.read_xml("word/document.xml")
        self.settings = self.pkg.read_xml("word/settings.xml")

    def paragraphs(self) -> List[etree._Element]:
        return cast(
            List[etree._Element],
            self.document.xpath(".//w:p", namespaces=NS),
        )

    def runs(self) -> List[etree._Element]:
        return cast(
            List[etree._Element],
            self.document.xpath(".//w:r", namespaces=NS),
        )

    def get_paragraph_text(
        self,
        paragraph: etree._Element,
    ) -> str:
        texts = cast(List, paragraph.xpath(".//w:t", namespaces=NS))
        return "".join(t.text for t in texts if t.text)

    def save(self) -> bytes:

        buffer = BytesIO()

        with zipfile.ZipFile(BytesIO(self.pkg._raw), "r") as r_file:
            with zipfile.ZipFile(
                buffer, "w", compression=zipfile.ZIP_DEFLATED
            ) as o_file:
                for item in r_file.infolist():
                    if item.filename == "word/document.xml":
                        o_file.writestr(
                            item,
                            etree.tostring(
                                self.document,
                                xml_declaration=True,
                                encoding="UTF-8",
                            ),
                        )

                    elif item.filename == "word/settings.xml":
                        o_file.writestr(
                            item,
                            etree.tostring(
                                self.settings,
                                xml_declaration=True,
                                encoding="UTF-8",
                            ),
                        )
                    else:
                        o_file.writestr(item, r_file.read(item.filename))

        return buffer.getvalue()
