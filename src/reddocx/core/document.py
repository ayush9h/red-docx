import zipfile
from datetime import datetime
from io import BytesIO
from typing import List, Optional, Union, cast

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

    def track_replace_words(
        self, replacements: dict[str, str], author: Optional[str] = "John DOE"
    ):

        results = {}
        revision_id = 1
        now = datetime.utcnow().isoformat() + "Z"

        paragraphs = self.paragraphs()

        for idx, p in enumerate(paragraphs):
            text = self.get_paragraph_text(p)

            for old, new in replacements.items():
                if old in text:

                    results.setdefault(old, []).append(idx)

                    self._apply_revision(
                        paragraph=p,
                        old_word=old,
                        new_word=new,
                        rev_id=revision_id,
                        date=now,
                        author=author,
                    )
                    revision_id += 1

        return results

    def _apply_revision(self, paragraph, old_word, new_word, rev_id, date, author):
        ns = NS["w"]

        for run in paragraph.xpath(".//w:r", namespaces=NS):
            t = run.find(".//w:t", namespaces=NS)

            if t is None or t.text is None:
                continue

            text = t.text

            if old_word not in text:
                continue

            before, match, after = text.partition(old_word)

            parent = run.getparent()
            idx = parent.index(run)

            if before:
                before_run = etree.Element(f"{{{ns}}}r")
                before_t = etree.Element(f"{{{ns}}}t")
                before_t.text = before
                before_run.append(before_t)
                parent.insert(idx, before_run)
                idx += 1

            # delete revision
            del_el = etree.Element(f"{{{ns}}}del")
            del_el.set(f"{{{ns}}}id", str(rev_id))
            del_el.set(f"{{{ns}}}author", author)
            del_el.set(f"{{{ns}}}date", date)

            del_run = etree.Element(f"{{{ns}}}r")
            del_text = etree.Element(f"{{{ns}}}delText")
            del_text.text = old_word
            del_run.append(del_text)
            del_el.append(del_run)

            parent.insert(idx, del_el)
            idx += 1

            ins_el = etree.Element(f"{{{ns}}}ins")
            ins_el.set(f"{{{ns}}}id", str(rev_id))
            ins_el.set(f"{{{ns}}}author", author)
            ins_el.set(f"{{{ns}}}date", date)

            ins_run = etree.Element(f"{{{ns}}}r")
            ins_text = etree.Element(f"{{{ns}}}t")
            ins_text.text = new_word
            ins_run.append(ins_text)
            ins_el.append(ins_run)

            parent.insert(idx, ins_el)
            idx += 1

            if after:
                after_run = etree.Element(f"{{{ns}}}r")
                after_t = etree.Element(f"{{{ns}}}t")
                after_t.text = after
                after_run.append(after_t)
                parent.insert(idx, after_run)

            parent.remove(run)
            break

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
