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

        metadata = self.get_metadata(
            qualification_code
        )

        dtwd = self.dtwd.get_qualification_metadata(
            qualification_code
        )

        return {
            "type": "Qualification",

            "qualificationCode":
                metadata["code"],

            "qualificationTitle":
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

            "pathways":
                dtwd.get(
                    "pathways",
                    []
                )
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

        with urlopen(request) as response:

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