import html
import json
import re

from urllib.parse import urlencode
from urllib.request import Request
from urllib.request import urlopen

from lxml import html as lxml_html


class DtwdClient:

    BASE_URL = "https://tps.dtwd.wa.gov.au"

    def get_metadata(
        self,
        code
    ):

        code = code.strip().upper()

        search_html = self.search(
            code
        )

        detail_url = self.extract_detail_url(
            search_html
        )

        detail_html = self.get_html(
            detail_url
        )

        if "/course/" in detail_url:

            return self.get_course_metadata(
                detail_url,
                detail_html
            )

        if "/qualification/" in detail_url:

            return self.get_qualification_metadata(
                detail_url,
                detail_html
            )

        if (
            "/unit-of-competency/" in detail_url
            or "/module/" in detail_url
        ):

            return self.get_unit_metadata(
                detail_url,
                detail_html
            )

        raise Exception(
            f"Unknown TPS type [{detail_url}]"
        )

    #
    # SEARCH
    #
    def search(
        self,
        code
    ):

        url = (
            f"{self.BASE_URL}"
            f"/?{urlencode({'Code': code})}"
        )

        return self.get_html(url)

    #
    # QUALIFICATIONS
    #
    def get_qualification_metadata(
        self,
        detail_url,
        detail_html
    ):

        units = self.extract_units(
            detail_html
        )

        packaging = self.extract_packaging_rules(
            detail_html
        )

        return {
            "type": "Qualification",
            "qualificationId": self.extract_guid(
                detail_url
            ),
            "nationalCode": self.extract_value(
                detail_html,
                "National Code"
            ),
            "title": self.extract_title(
                detail_html
            ),
            "stateCode": self.extract_value(
                detail_html,
                "State Code"
            ),
            "tgaStatus": self.extract_value(
                detail_html,
                "TGA Status"
            ),
            "dtwdStatus": self.extract_value(
                detail_html,
                "DTWD Status"
            ),
            "approvedDate": self.extract_value(
                detail_html,
                "Approved Date"
            ),
            "nominalHours": self.extract_value(
                detail_html,
                "Nominal Hours"
            ),
            "fieldOfEducation": self.extract_value(
                detail_html,
                "Field of Education"
            ),
            "detailUrl": detail_url,

            "pathways": self.extract_pathways(
                detail_html
            ),

            "packagingRules":
                packaging["packagingRules"],

            "totalUnitCount":
                packaging["totalUnits"],

            "coreUnitCount":
                packaging["coreUnits"],

            "electiveUnitCount":
                packaging["electiveUnits"],

            "requiredSpecialisationUnitCount":
                packaging["requiredSpecialisationUnits"],

            "coreUnits": [
                u for u in units
                if u["section"] == "Core"
            ],

            "electiveUnits": [
                u for u in units
                if u["section"] == "Elective"
            ]
        }

    #
    # ACCREDITED COURSES
    #
    def get_course_metadata(
        self,
        detail_url,
        detail_html
    ):

        units = self.extract_units(
            detail_html
        )

        packaging = self.extract_packaging_rules(
            detail_html
        )

        return {
            "type": "AccreditedCourse",
            "courseId": self.extract_guid(
                detail_url
            ),
            "nationalCode": self.extract_value(
                detail_html,
                "National Code"
            ),
            "title": self.extract_title(
                detail_html
            ),
            "stateCode": self.extract_value(
                detail_html,
                "State Code"
            ),
            "tgaStatus": self.extract_value(
                detail_html,
                "TGA Status"
            ),
            "dtwdStatus": self.extract_value(
                detail_html,
                "DTWD Status"
            ),
            "approvedDate": self.extract_value(
                detail_html,
                "Approved Date"
            ),
            "nominalHours": self.extract_value(
                detail_html,
                "Nominal Hours"
            ),
            "fieldOfEducation": self.extract_value(
                detail_html,
                "Field of Education"
            ),
            "detailUrl": detail_url,

            "pathways": self.extract_pathways(
                detail_html
            ),

           "packagingRules":
                packaging["packagingRules"],

            "totalUnitCount":
                packaging["totalUnits"],

            "coreUnitCount":
                packaging["coreUnits"],

            "electiveUnitCount":
                packaging["electiveUnits"],

            "requiredSpecialisationUnitCount":
                packaging["requiredSpecialisationUnits"],

            "coreUnits": [
                u for u in units
                if u["section"] == "Core"
            ],

            "electiveUnits": [
                u for u in units
                if u["section"] == "Elective"
            ]

        }

    #
    # UNITS / MODULES
    #
    def get_unit_metadata(
        self,
        detail_url,
        detail_html
    ):

        return {
            "type": "Unit",
            "guid": self.extract_guid(
                detail_url
            ),
            "nationalCode": self.extract_value(
                detail_html,
                "National Code"
            ),
            "title": self.extract_title(
                detail_html
            ),
            "stateCode": self.extract_value(
                detail_html,
                "State Code"
            ),
            "tgaStatus": self.extract_value(
                detail_html,
                "TGA Status"
            ),
            "dtwdStatus": self.extract_value(
                detail_html,
                "DTWD Status"
            ),
            "approvedDate": self.extract_value(
                detail_html,
                "Approved Date"
            ),
            "nominalHours": self.extract_value(
                detail_html,
                "Nominal Hours"
            ),
            "fieldOfEducation": self.extract_value(
                detail_html,
                "Field of Education"
            ),
            "detailUrl": detail_url
        }

    #
    # PATHWAYS
    #
    def extract_pathways(
        self,
        detail_html
    ):

        doc = lxml_html.fromstring(
            detail_html
        )

        pathways = []
        seen = set()

        rows = doc.xpath(
            "//table//tr[.//a[contains(@href,'/pathway/')]]"
        )

        for row in rows:

            cells = row.xpath("./td")

            if len(cells) < 2:
                continue

            state_code = (
                cells[0]
                .text_content()
                .strip()
            )

            stream_name = (
                cells[1]
                .text_content()
                .strip()
            )

            href = cells[0].xpath(
                ".//a/@href"
            )

            pathway_url = (
                href[0]
                if href else ""
            )

            if pathway_url.startswith("/"):

                pathway_url = (
                    self.BASE_URL
                    + pathway_url
                )

            key = (
                state_code,
                stream_name
            )

            if key in seen:
                continue

            seen.add(key)

            pathways.append(
                {
                    "pathwayId": self.extract_guid(
                        pathway_url
                    ),
                    "stateCode": state_code,
                    "streamName": stream_name,
                    "pathwayUrl": pathway_url
                }
            )

        return pathways

    #
    # CORE / ELECTIVE UNITS
    #
    def extract_units(
        self,
        detail_html
    ):

        doc = lxml_html.fromstring(
            detail_html
        )

        units = []

        headings = doc.xpath(
            "//h5"
        )

        for heading in headings:

            heading_text = (
                heading.text_content()
                .strip()
            )

            section = None
            stream = None

            if re.search(
                r"core units",
                heading_text,
                re.IGNORECASE
            ):

                section = "Core"

            elif (
                re.search(
                    r"elective units",
                    heading_text,
                    re.IGNORECASE
                )
                or
                re.search(
                    r"group\s+[A-Z]",
                    heading_text,
                    re.IGNORECASE
                )
            ):

                section = "Elective"

                stream = re.sub(
                    r"^\d+\.\s*",
                    "",
                    heading_text
                ).strip()

            else:
                continue

            table = heading.xpath(
                "following::table[1]"
            )

            rows = table[0].xpath(".//tr")

            if not table:
                continue

            table = table[0]

            rows = table.xpath(
                ".//tr"
            )

            for row in rows:

                cells = row.xpath("./td")

                if len(cells) < 5:
                    continue

                units.append(
                    {
                        "section": section,
                        "stream": stream,

                        "stateCode":
                            cells[0]
                            .text_content()
                            .strip(),

                        "nationalCode":
                            cells[1]
                            .text_content()
                            .strip(),

                        "title":
                            cells[2]
                            .text_content()
                            .strip(),

                        "hours":
                            cells[3]
                            .text_content()
                            .strip(),

                        "type":
                            cells[4]
                            .text_content()
                            .strip()
                    }
                )

        return units

    #
    # HTTP
    #
    def get_html(
        self,
        url
    ):

        request = Request(
            url,
            headers={
                "User-Agent":
                    "Mozilla/5.0"
            }
        )

        with urlopen(
            request
        ) as response:

            return response.read().decode(
                "utf-8",
                errors="ignore"
            )

    #
    # HELPERS
    #
    def extract_detail_url(
        self,
        html_text
    ):

        patterns = [

            r'href="([^"]*/course/[^"]+)"',

            r'href="([^"]*/qualification/[^"]+)"',

            r'href="([^"]*/unit-of-competency/[^"]+)"',

            r'href="([^"]*/module/[^"]+)"'
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                html_text,
                re.IGNORECASE
            )

            if match:

                href = html.unescape(
                    match.group(1)
                )

                if href.startswith("/"):

                    return (
                        self.BASE_URL
                        + href
                    )

                return href

        raise Exception(
            "Detail link not found"
        )

    def extract_guid(
        self,
        url
    ):

        match = re.search(
            r'([a-f0-9\-]{36})',
            url,
            re.IGNORECASE
        )

        return (
            match.group(1)
            if match
            else None
        )

    def extract_title(
        self,
        html_text
    ):

        doc = lxml_html.fromstring(
            html_text
        )

        title = doc.xpath(
            "//h4[1]/text()"
        )

        return (
            title[0].strip()
            if title
            else ""
        )

    def extract_value(
        self,
        html_text,
        label
    ):

        doc = lxml_html.fromstring(
            html_text
        )

        elements = doc.xpath(
            f"//strong[contains(text(),'{label}')]"
        )

        if not elements:
            return ""

        parent = elements[0].getparent()

        text = (
            parent.text_content()
            .replace(label, "", 1)
            .lstrip(":")
            .strip()
        )

        return text

    def extract_packaging_rules(
        self,
        detail_html
    ):

        doc = lxml_html.fromstring(
            detail_html
        )

        result = {
            "packagingRules": None,
            "totalUnits": None,
            "coreUnits": None,
            "electiveUnits": None,
            "requiredSpecialisationUnits": None
        }

        heading = doc.xpath(
            "//h5[contains(normalize-space(.), 'Packaging Rules')]"
        )

        if not heading:
            return result

        heading = heading[0]

        rule_div = heading.getnext()

        if rule_div is None:
            return result

        rules_text = " ".join(
            text.strip()
            for text in rule_div.xpath(".//text()")
            if text.strip()
        )

        result["packagingRules"] = rules_text

        #
        # Total Units
        #
        match = re.search(
            r'complete.*?\((\d+)\)\s*units',
            rules_text,
            re.IGNORECASE
        )

        if match:
            result["totalUnits"] = int(
                match.group(1)
            )

        #
        # Core Units
        #
        match = re.search(
            r'\((\d+)\)\s*core units',
            rules_text,
            re.IGNORECASE
        )

        if match:
            result["coreUnits"] = int(
                match.group(1)
            )

        #
        # Elective Units
        #
        match = re.search(
            r'\((\d+)\)\s*elective units',
            rules_text,
            re.IGNORECASE
        )

        if match:
            result["electiveUnits"] = int(
                match.group(1)
            )

        #
        # Accredited Course wording
        #
        match = re.search(
            r'minimum of.*?\((\d+)\)\s*units\s*must\s*be\s*selected',
            rules_text,
            re.IGNORECASE
        )

        if match:
            result["requiredSpecialisationUnits"] = int(
                match.group(1)
            )

        #
        # Qualification Pathway wording
        #
        if result["requiredSpecialisationUnits"] is None:

            match = re.search(
                r'select all.*?\((\d+)\)\s*elective units',
                rules_text,
                re.IGNORECASE
            )

            if match:
                result["requiredSpecialisationUnits"] = int(
                    match.group(1)
                )

        return result