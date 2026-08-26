import html
import re

from lxml import html as lxml_html
from lxml import etree

from http.cookiejar import CookieJar
from urllib.parse import urlencode
from urllib.request import (
    Request,
    build_opener,
    HTTPCookieProcessor
)


class DtwdClient:

    SEARCH_URL = (
        "https://tps.dtwd.wa.gov.au/apps/tps/Pages/default.aspx"
    )

    BASE_URL = (
        "https://tps.dtwd.wa.gov.au"
    )

    IDENTIFIER_FIELD = (
        "ctl00$SPWebPartManager1$g_72b3d09c_3e10_4bd3_8b9c_f4643f192177"
        "$ctl00$IdentifierTextBox"
    )

    SEARCH_BUTTON = (
        "ctl00$SPWebPartManager1$g_72b3d09c_3e10_4bd3_8b9c_f4643f192177"
        "$ctl00$SearchButton"
    )

    def __init__(self):

        self.cookie_jar = CookieJar()

        self.opener = build_opener(
            HTTPCookieProcessor(
                self.cookie_jar
            )
        )

    #
    # CODE TYPES
    #
    def is_qualification_code(
        self,
        code
    ):

        return bool(
            re.match(
                r"^[A-Z]{3}\d{5}$",
                code
            )
        )

    def is_unit_code(
        self,
        code
    ):

        return bool(
            re.match(
                r"^[A-Z]{6}\d{3}$",
                code
            )
        )

    #
    # QUALIFICATIONS
    #
    def get_qualification_metadata(
        self,
        qualification_code
    ):

        search_html = self.search(
            qualification_code
        )

        detail_url = self.extract_detail_url(
            search_html
        )

        detail_html = self.get_html(
            detail_url
        )

        result = {
            "type": "Qualification",
            "guid": self.extract_guid(
                detail_url
            ),
            "nationalCode": self.extract_span(
                detail_html,
                "NationalCodeLabel"
            ),
            "title": self.extract_span(
                detail_html,
                "NameLabel"
            ),
            "stateCode": self.extract_span(
                detail_html,
                "StateCodeLabel"
            ),
            "tgaStatus": self.extract_span(
                detail_html,
                "TgaStatusLabel"
            ),
            "dtwdStatus": self.extract_span(
                detail_html,
                "DtwdStatusLabel"
            ),
            "approvedDate": self.extract_span(
                detail_html,
                "ApprovedDateLabel"
            ),
            "nominalHours": self.extract_span(
                detail_html,
                "NominalHourLabel"
            ),
            "fieldOfEducation": self.extract_span(
                detail_html,
                "FieldOfEducationLabel"
            ),
            "detailUrl": detail_url,
            "pathways": []
        }

        try:

            schedule_html = self.get_schedule_html(
                detail_url
            )
            
            result["pathways"] = (
                self.extract_pathways(
                    schedule_html
                )
            )

        except Exception as ex:

            print(
                f"Schedule extraction failed: {ex}"
            )

        return result

    #
    # UNITS
    #
    def get_unit_metadata(
        self,
        unit_code
    ):

        search_html = self.search(
            unit_code
        )

        detail_url = self.extract_detail_url(
            search_html
        )

        detail_html = self.get_html(
            detail_url
        )

        return {
            "type": "Unit",
            "guid": self.extract_guid(
                detail_url
            ),
            "nationalCode": self.extract_span(
                detail_html,
                "NationalCodeLabel"
            ),
            "title": self.extract_span(
                detail_html,
                "NameLabel"
            ),
            "stateCode": self.extract_span(
                detail_html,
                "StateCodeLabel"
            ),
            "tgaStatus": self.extract_span(
                detail_html,
                "TgaStatusLabel"
            ),
            "dtwdStatus": self.extract_span(
                detail_html,
                "DtwdStatusLabel"
            ),
            "approvedDate": self.extract_span(
                detail_html,
                "ApprovedDateLabel"
            ),
            "nominalHours": self.extract_span(
                detail_html,
                "NominalHourLabel"
            ),
            "fieldOfEducation": self.extract_span(
                detail_html,
                "FieldOfEducationLabel"
            ),
            "detailUrl": detail_url
        }

    #
    # SEARCH
    #
    def search(
        self,
        identifier
    ):

        page_html = self.get_html(
            self.SEARCH_URL
        )

        payload = {

            "__VIEWSTATE":
                self.extract_hidden(
                    page_html,
                    "__VIEWSTATE"
                ),

            "__VIEWSTATEGENERATOR":
                self.extract_hidden(
                    page_html,
                    "__VIEWSTATEGENERATOR"
                ),

            "__EVENTVALIDATION":
                self.extract_hidden(
                    page_html,
                    "__EVENTVALIDATION"
                ),

            "__REQUESTDIGEST":
                self.extract_hidden(
                    page_html,
                    "__REQUESTDIGEST"
                ),

            "__EVENTTARGET": "",
            "__EVENTARGUMENT": "",

            "ctl00$SPWebPartManager1$g_72b3d09c_3e10_4bd3_8b9c_f4643f192177$ctl00$AllCheckBox":
                "on",

            "ctl00$SPWebPartManager1$g_72b3d09c_3e10_4bd3_8b9c_f4643f192177$ctl00$IncludeSupersededCheckBox":
                "on",

            self.IDENTIFIER_FIELD:
                identifier,

            self.SEARCH_BUTTON:
                "Search"
        }

        return self.post_html(
            self.SEARCH_URL,
            payload
        )

    #
    # SCHEDULE TAB
    #
    def get_schedule_html(
        self,
        qualification_url
    ):
        page_html = self.get_html(
            qualification_url
        )

        match = re.search(
            r'WebForm_PostBackOptions\(&quot;([^"]*ScheduleLinkButton)&quot;',
            page_html,
            re.IGNORECASE
        )

        if not match:
            raise Exception(
                "ScheduleLinkButton postback target not found"
            )

        event_target = match.group(1)

        payload = {

            "__VIEWSTATE":
                self.extract_hidden(
                    page_html,
                    "__VIEWSTATE"
                ),

            "__VIEWSTATEGENERATOR":
                self.extract_hidden(
                    page_html,
                    "__VIEWSTATEGENERATOR"
                ),

            "__EVENTVALIDATION":
                self.extract_hidden(
                    page_html,
                    "__EVENTVALIDATION"
                ),

            "__REQUESTDIGEST":
                self.extract_hidden(
                    page_html,
                    "__REQUESTDIGEST"
                ),

            "__EVENTTARGET":
                match.group(1),

            "__EVENTARGUMENT":
                ""
        }

        return self.post_html(
            qualification_url,
            payload
        )

    #
    # PATHWAYS
    #

    def extract_pathways(
        self,
        schedule_html
    ):

        doc = lxml_html.fromstring(
            schedule_html
        )

        pathways = []

        links = doc.xpath(
            "//a[contains(@href,'PathwayDetails.aspx')]"
        )

        for link in links:

            text = (
                link.text_content()
                .strip()
            )

            lines = [
                line.strip()
                for line in text.splitlines()
                if line.strip()
            ]

            if len(lines) < 2:
                continue

            state_code = lines[0]
            stream_name = lines[1]
            pathway_url = link.get("href")
            if pathway_url.startswith("/"):
                pathway_url = (
                self.BASE_URL
                + pathway_url
                )

            pathways.append(
                {
                    "stateCode": state_code,
                    "streamName": stream_name,
                    "pathwayUrl": pathway_url
                }
            )

        return pathways

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

        with self.opener.open(
            request
        ) as response:

            return response.read().decode(
                "utf-8",
                errors="ignore"
            )

    def post_html(
        self,
        url,
        payload
    ):

        encoded = urlencode(
            payload
        ).encode(
            "utf-8"
        )

        request = Request(
            url,
            data=encoded,
            headers={
                "User-Agent":
                    "Mozilla/5.0",

                "Content-Type":
                    "application/x-www-form-urlencoded"
            }
        )

        with self.opener.open(
            request
        ) as response:

            return response.read().decode(
                "utf-8",
                errors="ignore"
            )

    #
    # HELPERS
    #
    def extract_hidden(
        self,
        html_text,
        field_name
    ):

        match = re.search(
            rf'id="{field_name}"[^>]*value="([^"]*)"',
            html_text,
            re.I
        )

        if not match:
            raise Exception(
                f"{field_name} not found"
            )

        return html.unescape(
            match.group(1)
        )

    def extract_detail_url(
        self,
        html_text
    ):

        patterns = [

            r'href="([^"]*UoCEndorsement\.aspx[^"]+)"',

            r'href="([^"]*QualificationEndorsement\.aspx[^"]+)"'
        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                html_text,
                re.I
            )

            if match:

                url = html.unescape(
                    match.group(1)
                )

                if url.startswith("/"):

                    return (
                        self.BASE_URL
                        + url
                    )

                return url

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
            re.I
        )

        return (
            match.group(1)
            if match
            else None
        )

    def extract_span(
        self,
        html_text,
        id_fragment
    ):

        match = re.search(
            rf'id="[^"]*{id_fragment}"[^>]*>(.*?)</span>',
            html_text,
            re.I | re.S
        )

        if not match:
            return ""

        text = re.sub(
            r"<.*?>",
            "",
            match.group(1)
        )

        return html.unescape(
            text
        ).strip()