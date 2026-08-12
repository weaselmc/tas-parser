import html
from lxml import etree

class WordHtmlConverter:

    NS = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    }

    def word_xml_to_html(self, xml):

        if xml is None:
            return ""

        #
        # Handle python-docx objects
        #
        if hasattr(xml, "xml"):
            xml = xml.xml

        if isinstance(xml, str):
            root = etree.fromstring(
                xml.encode("utf-8")
            )
        else:
            root = xml

        return self._process_element(root)

    def _process_element(self, element):

        tag = etree.QName(element).localname

        if tag == "tbl":
            return self._table_to_html(element)

        if tag == "tc":
            return self._cell_to_html(element)

        if tag == "p":
            return self._paragraph_to_html(element)

        parts = []

        for child in element:
            parts.append(
                self._process_element(child)
            )

        return "".join(parts)

    def _table_to_html(self, table):

        rows_html = []

        rows_html.append(
            '<table border="1">'
        )

        rows = table.xpath(
            "./w:tr",
            namespaces=self.NS
        )

        for row in rows:

            rows_html.append("<tr>")

            cells = row.xpath(
                "./w:tc",
                namespaces=self.NS
            )

            for cell in cells:

                rows_html.append(
                    (
                        "<td>"
                        f"{self._cell_to_html(cell)}"
                        "</td>"
                    )
                )

            rows_html.append("</tr>")

        rows_html.append("</table>")

        return "\n".join(rows_html)

    def _cell_to_html(self, cell):

        parts = []

        in_list = False

        paragraphs = cell.xpath(
            "./w:p",
            namespaces=self.NS
        )

        for paragraph in paragraphs:

            is_list = bool(
                paragraph.xpath(
                    "./w:pPr/w:numPr",
                    namespaces=self.NS
                )
            )

            html = self._paragraph_to_html(
                paragraph
            )

            if not html:
                continue

            if is_list:

                if not in_list:
                    parts.append("<ul>")
                    in_list = True

                parts.append(html)

            else:

                if in_list:
                    parts.append("</ul>")
                    in_list = False

                parts.append(html)

        if in_list:
            parts.append("</ul>")

        return "\n".join(parts)

    def _paragraph_to_html(self, paragraph):

        is_list = bool(
            paragraph.xpath(
                "./w:pPr/w:numPr",
                namespaces=self.NS
            )
        )

        runs = paragraph.xpath(
            "./w:r",
            namespaces=self.NS
        )

        #
        # Build formatting groups
        #
        groups = []

        for run in runs:

            text = "".join(
                run.xpath(
                    ".//w:t/text()",
                    namespaces=self.NS
                )
            )

            if not text:
                continue

            fmt = (
                bool(run.xpath(
                    "./w:rPr/w:b",
                    namespaces=self.NS
                )),
                bool(run.xpath(
                    "./w:rPr/w:i",
                    namespaces=self.NS
                )),
                bool(run.xpath(
                    "./w:rPr/w:u",
                    namespaces=self.NS
                ))
            )

            #
            # Merge adjacent runs with same formatting
            #
            if groups and groups[-1]["fmt"] == fmt:

                groups[-1]["text"] += text

            else:

                groups.append({
                    "fmt": fmt,
                    "text": text
                })

        if not groups:
            return ""

        #
        # Is entire paragraph bold?
        #
        all_bold = all(
            g["fmt"][0]
            for g in groups
            if g["text"].strip()
        )

        #
        # Heading detection
        #
        plain_text = "".join(
            g["text"]
            for g in groups
        ).strip()

        if (
            all_bold
            and plain_text
        ):
            return f"<h4>{html.escape(plain_text)}</h4>"

        #
        # Normal paragraph
        #
        parts = []

        for group in groups:

            text = html.escape(
                group["text"]
            )

            bold, italic, underline = (
                group["fmt"]
            )

            if underline:
                text = f"<u>{text}</u>"

            if italic:
                text = f"<em>{text}</em>"

            if bold:
                text = f"<strong>{text}</strong>"

            parts.append(text)

        content = "".join(parts).strip()

        if not content:
            return ""

        if is_list:
            return f"<li>{content}</li>"

        return f"<p>{content}</p>"


    def _run_to_html(self, run):

        #
        # Checkbox
        #
        checked = run.xpath(
            ".//w14:checked[@w14:val='1']",
            namespaces=self.NS
        )

        if checked:
            return "☑"

        text = "".join(
            run.xpath(
                ".//w:t/text()",
                namespaces=self.NS
            )
        )

        if not text:

            br = run.xpath(
                ".//w:br",
                namespaces=self.NS
            )

            if br:
                return "<br>"

            return ""

        text = html.escape(text)

        #
        # Formatting
        #
        if run.xpath(
            "./w:rPr/w:b",
            namespaces=self.NS
        ):
            text = (
                f"<strong>{text}</strong>"
            )

        if run.xpath(
            "./w:rPr/w:i",
            namespaces=self.NS
        ):
            text = (
                f"<em>{text}</em>"
            )

        if run.xpath(
            "./w:rPr/w:u",
            namespaces=self.NS
        ):
            text = (
                f"<u>{text}</u>"
            )

        return text

    def _hyperlink_to_html(self, hyperlink):

        text = ""

        for run in hyperlink.xpath(
            ".//w:r",
            namespaces=self.NS
        ):
            text += self._run_to_html(
                run
            )

        #
        # Relationship lookup could
        # be added later.
        #
        return text