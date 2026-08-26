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
        # Qualification
        #
        if re.match(
            r"^[A-Z]{3}\d{5}$",
            code
        ):
            return self.qual.extract(
                code
            )

        #
        # Unit
        #
        if re.match(
            r"^[A-Z]{6}\d{3}$",
            code
        ):
            return self.uoc.extract(
                code
            )

        raise Exception(
            f"Unsupported code [{code}]"
        )