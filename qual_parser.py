from dtwdclient import DtwdClient

import json

from urllib.request import Request
from urllib.request import urlopen


class QualParser:

    API_BASE = "https://training.gov.au/api"

    def __init__(self):

        self.dtwd = DtwdClient()

    #
    # Public Entry Point
    #
    def extract(
        self,
        qualification_code
    ):

        qualification_code = (
            qualification_code
            .strip()
            .upper()
        )

        #
        # Training.gov.au may not support
        # accredited courses well, so make
        # metadata optional.
        #
        try:

            metadata = self.get_metadata(
                qualification_code
            )

            title = metadata.get(
                "title"
            )

            national_code = metadata.get(
                "code"
            )

        except Exception:

            metadata = {}

            title = None

            national_code = (
                qualification_code
            )

        #
        # DTWD is now the primary source
        #
        dtwd = self.dtwd.get_metadata(
            qualification_code
        )

        return {
            "type":
                dtwd.get("type"),

            "qualificationCode":
                national_code
                or dtwd.get("nationalCode"),

            "qualificationTitle":
                title
                or dtwd.get("title"),

            "stateCode":
                dtwd.get("stateCode"),

            "dtwdStatus":
                dtwd.get("dtwdStatus"),

            "tgaStatus":
                dtwd.get("tgaStatus"),

            "approvedDate":
                dtwd.get("approvedDate"),

            "nominalHours":
                dtwd.get("nominalHours"),

            "fieldOfEducation":
                dtwd.get("fieldOfEducation"),

            "detailUrl":
                dtwd.get("detailUrl"),

            "pathways":
                dtwd.get("pathways", []),

            "coreUnits":
                dtwd.get("coreUnits", []),

            "electiveUnits":
                dtwd.get("electiveUnits", [])
        }

    #
    # API
    #
    def get_json(
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

            return json.loads(
                response.read().decode(
                    "utf-8"
                )
            )

    def get_metadata(
        self,
        qualification_code
    ):

        return self.get_json(
            f"{self.API_BASE}/training/"
            f"{qualification_code}"
            f"?include=all"
            f"&api-version=1.0"
        )