import json
import re

import azure.functions as func

from urllib.request import Request
from urllib.request import urlopen


app = func.FunctionApp()


class UocParser:

    BASE_URL = "https://training.gov.au/Training/Details/{code}"

    def extract(self, unit_code: str):

        html = self._download_page(unit_code)

        return {
            "unitCode": unit_code,
            "unitTitle": self._extract_title(html),
            "items": self._extract_all_items(unit_code, html)
        }

    def _download_page(self, unit_code: str):

        request = Request(
            self.BASE_URL.format(code=unit_code),
            headers={
                "User-Agent": "Mozilla/5.0"
            }
        )

        with urlopen(request) as response:

            return response.read().decode(
                "utf-8",
                errors="ignore"
            )

    def _extract_title(self, html: str):

        match = re.search(
            r"<title>(.*?)</title>",
            html,
            re.IGNORECASE | re.DOTALL
        )

        if match:
            return re.sub(
                r"\s+",
                " ",
                match.group(1)
            ).strip()

        return ""

    def _extract_all_items(
        self,
        unit_code: str,
        html: str
    ):

        items = []

        items.extend(
            self._extract_elements(
                unit_code,
                html
            )
        )

        items.extend(
            self._extract_pcs(
                unit_code,
                html
            )
        )

        items.extend(
            self._extract_ke(
                unit_code,
                html
            )
        )

        items.extend(
            self._extract_pe(
                unit_code,
                html
            )
        )

        items.extend(
            self._extract_conditions(
                unit_code,
                html
            )
        )

        return items

    #
    # ELEMENTS
    #
    def _extract_elements(
        self,
        unit_code: str,
        html: str
    ):

        elements = []

        # TODO:
        # Find Elements section
        # Build:
        #
        # {
        #   "id": "ICTNWK536.1",
        #   "parentId": "ICTNWK536",
        #   "itemType": "Element"
        # }

        return elements

    #
    # PERFORMANCE CRITERIA
    #
    def _extract_pcs(
        self,
        unit_code: str,
        html: str
    ):

        pcs = []

        # TODO:
        # Create IDs:
        #
        # ICTNWK536.1.1
        # ICTNWK536.1.2

        return pcs

    #
    # KNOWLEDGE EVIDENCE
    #
    def _extract_ke(
        self,
        unit_code: str,
        html: str
    ):

        items = []

        # TODO:
        # Create IDs:
        #
        # ICTNWK536.KE1
        # ICTNWK536.KE1.1

        return items

    #
    # PERFORMANCE EVIDENCE
    #
    def _extract_pe(
        self,
        unit_code: str,
        html: str
    ):

        items = []

        # TODO:
        # Create IDs:
        #
        # ICTNWK536.PE1
        # ICTNWK536.PE1.1

        return items

    #
    # ASSESSMENT CONDITIONS
    #
    def _extract_conditions(
        self,
        unit_code: str,
        html: str
    ):

        items = []

        # TODO:
        # Create IDs:
        #
        # ICTNWK536.AC1
        # ICTNWK536.AC1.1

        return items


@app.route(
    route="get-uocinfo",
    methods=["POST"],
    auth_level=func.AuthLevel.FUNCTION
)
def get_uoc_info(
    req: func.HttpRequest
) -> func.HttpResponse:

    try:

        body = req.get_json()

        unit_codes = body.get(
            "unitCodes",
            []
        )

        parser = UocParser()

        output = []

        for code in unit_codes:

            #
            # Skip Victorian units
            #
            if code.upper().startswith("VU"):

                output.append(
                    {
                        "unitCode": code,
                        "status": "Skipped",
                        "reason": "Victorian accredited unit"
                    }
                )

                continue

            try:

                output.append(
                    parser.extract(code)
                )

            except Exception as ex:

                output.append(
                    {
                        "unitCode": code,
                        "status": "Error",
                        "message": str(ex)
                    }
                )

        return func.HttpResponse(
            json.dumps(
                {
                    "units": output
                }
            ),
            mimetype="application/json",
            status_code=200
        )

    except Exception as ex:

        return func.HttpResponse(
            json.dumps(
                {
                    "status": "Error",
                    "message": str(ex)
                }
            ),
            mimetype="application/json",
            status_code=500
        )