import re

from uoc_parser import UocParser
from qual_parser import QualParser


class TPParser:

    def __init__(self):

        self.uoc = UocParser()
        self.qual = QualParser()

    def extract(
        self,
        code
    ):

        code = code.strip().upper()

        #
        # Accredited Course
        # 22697VIC
        #
        if re.match(
            r"^\d+VIC$",
            code
        ):
            return self.qual.extract(
                code
            )

        #
        # Training Package Qualification
        # ICT50220
        #
        if re.match(
            r"^[A-Z]{3}\d{5}$",
            code
        ):
            return self.qual.extract(
                code
            )

        #
        # Victorian Unit
        # VU23884
        #
        if re.match(
            r"^VU\d+$",
            code
        ):
            return self.uoc.extract(
                code
            )

        #
        # Training Package Unit
        # ICTNWK536
        #
        if re.match(
            r"^[A-Z]{2,}\d{3,}$",
            code
        ):
            return self.uoc.extract(
                code
            )

        raise Exception(
            f"Unsupported code [{code}]"
        )