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

            return self.transform_qualification(
                self.qual.extract(code)
            )

        #
        # Qualification
        # ICT50220
        #
        if re.match(
            r"^[A-Z]{3}\d{5}$",
            code
        ):

            return self.transform_qualification(
                self.qual.extract(code)
            )

        #
        # Victorian Unit
        # VU23884
        #
        if re.match(
            r"^VU\d+$",
            code
        ):

            return self.transform_unit(
                self.uoc.extract(code)
            )

        #
        # Training Package Unit
        # ICTNWK536
        #
        if re.match(
            r"^[A-Z]{2,}\d{3,}$",
            code
        ):

            return self.transform_unit(
                self.uoc.extract(code)
            )

        raise Exception(
            f"Unsupported code [{code}]"
        )

    #
    # QUALIFICATIONS
    #
    def transform_qualification(
        self,
        source
    ):

        result = {
            "qualifications": [],
            "pathways": [],
            "units": [],
            "qualificationUnits": []
        }

        qualification_code = (
            source.get("stateCode")
        )

        result["qualifications"].append(
            {
                "qualificationCode":
                    qualification_code,

                "nationalCode":
                    source.get(
                        "qualificationCode"
                    ),

                "qualificationTitle":
                    source.get(
                        "qualificationTitle"
                    ),

                "qualificationType":
                    source.get(
                        "type"
                    ),

                "courseId":
                    source.get(
                        "courseId"
                    ),

                "qualificationId":
                    source.get(
                        "qualificationId"
                    ),

                "dtwdStatus":
                    source.get(
                        "dtwdStatus"
                    ),

                "tgaStatus":
                    source.get(
                        "tgaStatus"
                    ),

                "approvedDate":
                    source.get(
                        "approvedDate"
                    ),

                "nominalHours":
                    source.get(
                        "nominalHours"
                    ),

                "fieldOfEducation":
                    source.get(
                        "fieldOfEducation"
                    ),

                "detailUrl":
                    source.get(
                        "detailUrl"
                    ),

                "packagingRules":
                    source.get(
                        "packagingRules"
                    ),

                "coreUnitCount":
                    source.get("coreUnitCount"),

                "electiveUnitCount":
                   source.get("electiveUnitCount"),

                "totalUnitCount":
                    source.get("totalUnitCount"),

                "requiredSpecialisationUnitCount":
                    source.get("requiredSpecialisationUnitCount"),

            }
        )

        #
        # Pathways
        #
        for pathway in source.get(
            "pathways",
            []
        ):

            result["pathways"].append(
                {
                    "pathwayId":
                        pathway.get(
                            "pathwayId"
                        ),

                    "qualificationCode":
                        qualification_code,

                    "stateCode":
                        pathway.get(
                            "stateCode"
                        ),

                    "pathwayName":
                        pathway.get(
                            "streamName"
                        ),

                    "pathwayUrl":
                        pathway.get(
                            "pathwayUrl"
                        )
                }
            )

        #
        # Units
        #
        seen = set()

        all_units = (
            source.get("coreUnits", [])
            +
            source.get("electiveUnits", [])
        )

        for unit in all_units:

            state_code = (
                unit.get(
                    "stateCode"
                )
            )

            if state_code in seen:
                continue

            seen.add(
                state_code
            )

            result["units"].append(
                {
                    "stateCode":
                        state_code,

                    "nationalCode":
                        unit.get(
                            "nationalCode"
                        ),

                    "title":
                        unit.get(
                            "title"
                        ),

                    "unitHours":
                        unit.get(
                            "hours"
                        ),

                    "unitType":
                        unit.get(
                            "type"
                        )
                }
            )

        #
        # Qualification Units
        #
        for unit in source.get(
            "coreUnits",
            []
        ):

            result[
                "qualificationUnits"
            ].append(
                {
                    "qualificationCode":
                        qualification_code,

                    "unitCode":
                        unit.get(
                            "stateCode"
                        ),

                    "coreElective":
                        "C",

                    "electiveGroup":
                        None
                }
            )

        for unit in source.get(
            "electiveUnits",
            []
        ):

            stream = (
                unit.get(
                    "stream"
                ) or ""
            )

            #
            # Remove prefix
            #
            stream = re.sub(
                r"^Elective units:\s*",
                "",
                stream,
                flags=re.IGNORECASE
            ).strip()

            #
            # Optional cleanup
            #
            stream = re.sub(
                r"\s+Stream$",
                "",
                stream,
                flags=re.IGNORECASE
            ).strip()

            result[
                "qualificationUnits"
            ].append(
                {
                    "qualificationCode":
                        qualification_code,

                    "unitCode":
                        unit.get(
                            "stateCode"
                        ),

                    "coreElective":
                        "E",

                    "electiveGroup":
                        stream
                }
            )

        return result

    #
    # SINGLE UNIT
    #
    def transform_unit(
        self,
        source
    ):

        return {
            "qualifications": [],
            "pathways": [],
            "units": [
                {
                    "stateCode":
                        source.get(
                            "stateCode"
                        ),

                    "nationalCode":
                        source.get(
                            "nationalCode"
                        ),

                    "title":
                        source.get(
                            "title"
                        ),

                    "unitHours":
                        source.get(
                            "nominalHours"
                        ),

                    "unitType":
                        source.get(
                            "type"
                        )
                }
            ],
            "qualificationUnits": []
        }