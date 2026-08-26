import json
import re
import html
from dtwdclient import DtwdClient

from urllib.request import Request
from urllib.request import urlopen


class UocParser:

    API_BASE = "https://training.gov.au/api"

    def __init__(self):
        self.dtwd = DtwdClient()

    #
    # Public Entry Point
    #
    def extract(self, unit_code):

        metadata = self.get_metadata(unit_code)

        dtwd = self.dtwd.get_unit_metadata(
            unit_code
        )

        release_number = self.get_latest_release_number(
            metadata
        )

        release = self.get_release(
            unit_code,
            release_number
        )

        items = []

        for bundle in release.get(
            "contentBundles",
            []
        ):

            bundle_data = self.get_bundle(
                bundle["id"]
            )

            items.extend(
                self.parse_bundle(
                    unit_code,
                    bundle_data
                )
            )

        order = {
            "Element": 1,
            "PC": 2,
            "KE": 3,
            "PE": 4,
            "AC": 5
        }

        def sort_key(item):
            return (
                order.get(
                    item["itemType"],
                    99
                ),
                [
                    int(x)
                    if x.isdigit()
                    else x
                    for x in re.split(
                        r'(\d+)',
                        item["id"]
                    )
                    if x
                ]
            )

        items.sort(key=sort_key)

        return {
            "type": "Unit",

            "unitCode":
                metadata["code"],

            "unitTitle":
                metadata["title"],

            "stateCode":
                dtwd.get(
                    "stateCode"
                ),

            "dtwdStatus":
                dtwd.get(
                    "dtwdStatus"
                ),

            "approvedDate":
                dtwd.get(
                    "approvedDate"
                ),

            "nominalHours":
                dtwd.get(
                    "nominalHours"
                ),

            "fieldOfEducation":
                dtwd.get(
                    "fieldOfEducation"
                ),

            "detailUrl":
                dtwd.get(
                    "detailUrl"
                ),

            "items":
                items
        }
    
    #
    # API
    #
    def get_json(self, url):

        request = Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0"
            }
        )

        with urlopen(request) as response:

            return json.loads(
                response.read().decode(
                    "utf-8"
                )
            )

    def get_metadata(
        self,
        unit_code
    ):

        return self.get_json(
            f"{self.API_BASE}/training/"
            f"{unit_code}"
            f"?include=all"
            f"&api-version=1.0"
        )

    def get_latest_release_number(
        self,
        metadata
    ):

        releases = metadata.get(
            "releases",
            []
        )

        if not releases:

            raise Exception(
                "No releases found"
            )

        return releases[0][
            "releaseNumber"
        ]

    def get_release(
        self,
        unit_code,
        release_number
    ):

        return self.get_json(
            f"{self.API_BASE}/training/"
            f"{unit_code}/releases/"
            f"{release_number}"
            f"?include=all"
            f"&api-version=1.0"
        )

    def get_bundle(
        self,
        bundle_id
    ):

        return self.get_json(
            f"{self.API_BASE}/content/bundle/"
            f"{bundle_id}"
        )

    #
    # Bundle Processing
    #
    def parse_bundle(
        self,
        unit_code,
        bundle
    ):

        items = []

        for bundle_item in bundle.get(
            "items",
            []
        ):

            content_type = bundle_item.get(
                "contentType"
            )

            content = html.unescape(
                bundle_item.get(
                    "content",
                    ""
                )
            )

            if (
                content_type
                == "PerformanceCriteria"
            ):

                items.extend(
                    self.parse_performance_criteria(
                        unit_code,
                        content
                    )
                )

            elif (
                content_type
                == "KnowledgeEvidence"
            ):

                items.extend(
                    self.parse_hierarchy(
                        unit_code,
                        content,
                        "KE"
                    )
                )

            elif (
                content_type
                == "PerformanceEvidence"
            ):

                items.extend(
                    self.parse_hierarchy(
                        unit_code,
                        content,
                        "PE"
                    )
                )

            elif (
                content_type
                == "AssessmentConditions"
            ):

                items.extend(
                    self.parse_hierarchy(
                        unit_code,
                        content,
                        "AC"
                    )
                )

        return items

    #
    # Elements + PCs
    #
    def parse_performance_criteria(
        self,
        unit_code,
        content
    ):

        items = []

        rows = re.findall(
            r"<tr>(.*?)</tr>",
            content,
            re.DOTALL
        )

        current_element_id = None

        for row in rows:

            paragraphs = re.findall(
                r"<p>(.*?)</p>",
                row,
                re.DOTALL
            )

            cleaned = []

            for p in paragraphs:

                text = self.clean_text(p)

                if text:

                    cleaned.append(text)

            if not cleaned:
                continue

            first = cleaned[0]

            element_match = re.match(
                r"^(\d+)\.\s+(.*)",
                first
            )

            if element_match:

                element_no = (
                    element_match.group(1)
                )

                element_text = (
                    element_match.group(2)
                )

                current_element_id = (
                    f"{unit_code}."
                    f"{element_no}"
                )

                items.append(
                    {
                        "id":
                            current_element_id,
                        "parentId":
                            unit_code,
                        "itemType":
                            "Element",
                        "referenceNo":
                            element_no,
                        "description":
                            element_text
                    }
                )

                for line in cleaned[1:]:

                    pc_match = re.match(
                        r"^(\d+\.\d+)\s+(.*)",
                        line
                    )

                    if pc_match:

                        pc_no = (
                            pc_match.group(1)
                        )

                        pc_text = (
                            pc_match.group(2)
                        )

                        items.append(
                            {
                                "id":
                                    f"{unit_code}.{pc_no}",
                                "parentId":
                                    current_element_id,
                                "itemType":
                                    "PC",
                                "referenceNo":
                                    pc_no,
                                "description":
                                    pc_text
                            }
                        )

        return items

    #
    # KE / PE / AC
    #
    def parse_hierarchy(
        self,
        unit_code,
        content,
        prefix
    ):

        roots = self.extract_nested_list(
            content
        )

        items = []

        counter = 1

        for node in roots:

            id_val = (
                f"{unit_code}."
                f"{prefix}"
                f"{counter}"
            )

            items.append(
                {
                    "id": id_val,
                    "parentId":
                        unit_code,
                    "itemType":
                        prefix,
                    "referenceNo":
                        f"{prefix}{counter}",
                    "description":
                        self.cleanup_heading(
                            node["text"]
                        )
                }
            )

            self.add_children(
                items,
                id_val,
                node["children"],
                prefix
            )

            counter += 1

        return items

    def add_children(
        self,
        items,
        parent_id,
        children,
        prefix
    ):

        child_no = 1

        for child in children:

            child_id = (
                f"{parent_id}."
                f"{child_no}"
            )

            items.append(
                {
                    "id":
                        child_id,
                    "parentId":
                        parent_id,
                    "itemType":
                        prefix,
                    "referenceNo": ".".join(
                            child_id.split(".")[1:]
                        ),
                    "description":
                        self.cleanup_heading(
                            child["text"]
                        )
                }
            )

            self.add_children(
                items,
                child_id,
                child["children"],
                prefix
            )

            child_no += 1

    #
    # Nested UL Parsing
    #
    def cleanup_heading(self, text):
        text = re.sub(
            r',?\s*including:\s*$',
            '',
            text,
            flags=re.IGNORECASE
        )
        
        text = re.sub(
            r':\s*$',
            '',
            text
        )        

        text = text.strip()
        return text

    def extract_nested_list(
        self,
        html_text
    ):

        html_text = re.sub(
            r"</p>",
            "",
            html_text
        )

        html_text = re.sub(
            r"<p[^>]*>",
            "",
            html_text
        )

        token_pattern = (
            r"<ul>|</ul>|<li>|</li>|[^<>]+"
        )

        tokens = re.findall(
            token_pattern,
            html_text,
            re.DOTALL
        )

        roots = []
        stack = []

        for token in tokens:

            token = token.strip()

            if not token:
                continue

            if token == "<li>":

                node = {
                    "text": "",
                    "children": []
                }

                if stack:

                    stack[-1][
                        "children"
                    ].append(node)

                else:

                    roots.append(node)

                stack.append(node)

            elif token == "</li>":

                if stack:
                    stack.pop()

            elif token.startswith("<"):

                continue

            else:

                if stack:

                    text = self.clean_text(
                        token
                    )

                    if text:

                        if stack[-1][
                            "text"
                        ]:

                            stack[-1][
                                "text"
                            ] += (
                                " "
                                + text
                            )

                        else:

                            stack[-1][
                                "text"
                            ] = text

        return roots

    #
    # Helpers
    #
    def clean_text(
        self,
        text
    ):

        text = re.sub(
            r"<.*?>",
            "",
            text
        )

        text = html.unescape(text)

        text = re.sub(
            r"\s+",
            " ",
            text
        )

        return text.strip()