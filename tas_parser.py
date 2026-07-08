from docx import Document
from io import BytesIO
from docx.oxml.ns import qn
import re
import xml
import html

class TASDoc:

    def __init__(self, source):

        self.source = source
        
        if isinstance(source, bytes):
            self.source = BytesIO(source)        
        self.doc = Document(self.source)

        self.tables = self._load_tables()
        self.template = self._detect_template()
        
        # Parsed data
        self.qualification = None
        self.delivery = None
        self.units = None
        
        self.trainers = None
        self.assessment_matrices = None
        self.assessment_calendars = None

        # Relationship data
        self.clusters = {}
        self.cluster_units = {}
        self.qualification_clusters = {}
        self.qualification_units = {}


    # =====================================================
    # CORE
    # =====================================================

    @staticmethod
    def clean(text):

        if text is None:
            return ""

        return " ".join(str(text).split())

    def _load_tables(self):

        tables = []

        for table_number, table in enumerate(self.doc.tables):

            rows = []

            for row in table.rows:

                rows.append([
                    self.clean(cell.text)
                    for cell in row.cells
                ])

            tables.append({
                "table_number": table_number,
                "table": table,
                "rows": rows
            })

        return tables

    def _detect_template(self):

        for table in self.tables:

            rows = table["rows"]

            for row in rows:

                text = " ".join(row)

                if (
                    "This TAS version is current as at:"
                    in text
                ):
                    return "new"

                if (
                    "Training package code and title"
                    in text
                    and len(self.tables) > 28
                ):
                    return "new"

        return "old"

    # =====================================================
    # QUALIFICATION
    # =====================================================

    def get_qualification(self):

        for table in self.tables:
            
            table_xml = str(table["table"]._tbl.xml)

            rows = table["rows"]

            for i, row in enumerate(rows):

                text = " ".join(row).lower()

                if (
                    "qualification national code and title"
                    in text
                ):

                    if i < 1 or i + 1 >= len(rows):
                        continue

                    training_package_row = rows[i - 1]
                    qualification_row = rows[i + 1]

                    # Training package

                    package_value = (
                        training_package_row[0]
                        .strip()
                    )

                    if " - " in package_value:

                        package_code, package_title = (
                            package_value.split(
                                " - ",
                                1
                            )
                        )

                    else:

                        package_code = None
                        package_title = package_value

                    # Qualification

                    qualification_value = (
                        qualification_row[0]
                        .strip()
                    )

                    qualification_parts = (
                        qualification_value.split(
                            " ",
                            1
                        )
                    )

                    if len(qualification_parts) == 2:

                        qualification_code = (
                            qualification_parts[0]
                        )

                        qualification_title = (
                            qualification_parts[1]
                        )

                    else:

                        qualification_code = None
                        qualification_title = (
                            qualification_value
                        )

                    return {

                        "training_package": {

                            "code":
                                package_code,

                            "title":
                                package_title,

                            "release_number":
                                training_package_row[1]
                                if len(training_package_row) > 1
                                else None,

                            "release_date":
                                training_package_row[2]
                                if len(training_package_row) > 2
                                else None
                        },

                        "qualification": {

                            "national_code":
                                qualification_code,

                            "title":
                                qualification_title,

                            "release_number":
                                qualification_row[1]
                                if len(qualification_row) > 1
                                else None,

                            "release_date":
                                qualification_row[2]
                                if len(qualification_row) > 2
                                else None,

                            "state_code":
                                qualification_row[3]
                                if len(qualification_row) > 3
                                else None,

                            "on_scope":
                                ( 'w14:checked w14:val="1"' in table_xml)
                        }
                    }

        return {}

    # =====================================================
    # DELIVERY DETAILS
    # =====================================================

    def get_delivery(self):

        result = {
            "delivery": None,
            "duration": None,
            "start_date": None,
            "finish_date": None
        }

        for table in self.tables:

            for row in table["rows"]:

                for i in range(len(row) - 1):

                    label = (
                        row[i]
                        .strip()
                        .lower()
                        .rstrip(":")
                    )

                    value = row[i + 1].strip()

                    if label == "duration":

                        result["duration"] = value

                        xml = table["table"]._tbl.xml

                        dates = re.findall(
                            r'w:fullDate="([^"]+)"',
                            xml
                        )

                        if len(dates) >= 2:

                            from datetime import datetime

                            result["start_date"] = (
                                datetime.fromisoformat(
                                    dates[0].replace("Z", "+00:00")
                                ).date()
                            )

                            result["finish_date"] = (
                                datetime.fromisoformat(
                                    dates[1].replace("Z", "+00:00")
                                ).date()
                            )

                    elif (
                            "campus" in label
                            or "delivery location" in label
                            or "delivery site" in label
                        ):
                            result["delivery"] = value

        return result

    # =====================================================
    # DELIVERY SCHEDULE
    # =====================================================

    def _is_delivery_schedule(self, rows):

        if len(rows) < 2:
            return False

        header_text = " ".join(
            " ".join(row)
            for row in rows[:2]
        ).lower()

        return (
            "cluster" in header_text
            and "unit of competency" in header_text
            and "nominal hours" in header_text
        )    

    def parse_core_elective(self, value):

        value = value.strip().upper()

        if value == "C":
            return "C", None

        if "IMPORT" in value:
            return "E", "IMPORT"

        match = re.search(
            r"E\s*\(([A-Z])\)",
            value
        )

        if match:
            return "E", match.group(1)

        return "E", None

    def get_units(self):

        units = []

        stage = 1

        for table in self.tables:

            rows = table["rows"]
            docx_table = table["table"]            
            cluster_map = self.build_cluster_map(
                    docx_table
                )


            if not self._is_delivery_schedule(rows):
                continue

            if units:
                stage += 1

            for row_index, row in enumerate(rows):

                if len(row) < 4:
                    continue

                code = row[1].strip()

                if not re.fullmatch(
                    r"[A-Z]{3,}\d{3,}",
                    code
                ):
                    continue

                # column mappings differ slightly
                if self.template == "new":

                    delivery_mode = (
                        row[13]
                        if len(row) > 13
                        else ""
                    )

                    core_elective_raw = (
                        row[14]
                        if len(row) > 14
                        else ""
                    )
                    
                    core_elective, elective_group = (
                        self.parse_core_elective(
                            core_elective_raw
                        )
)


                else:

                    delivery_mode = (
                        row[7]
                        if len(row) > 7
                        else ""
                    )

                    core_elective_raw = (
                        row[8]
                        if len(row) > 8
                        else ""
                    )

                    core_elective, elective_group = (
                        self.parse_core_elective(
                            core_elective_raw
                        )
                    )
                
                units.append({
                    "stage": stage,
                    "cluster":(
                                " ".join(cluster_map[row_index].split())
                                if cluster_map.get(row_index)
                                else None
                            ),
                    "national_code": code,
                    "state_code": row[2].strip(),
                    "unit_title": row[3].strip(),
                    "nominal_hours": (
                        row[4]
                        if len(row) > 4
                        else ""
                    ),
                    "delivery_mode": delivery_mode,
                    "core_elective": core_elective,
                    "elective_group": elective_group
                })

        return units

    # =====================================================
    # TRAINERS
    # =====================================================

    def get_trainers(self):

        trainers = []

        for table in self.tables:

            rows = table["rows"]

            if len(rows) < 2:
                continue

            header_text = " ".join(
                " ".join(r)
                for r in rows[:2]
            ).lower()

            #
            # OLD TEMPLATE
            #
            if "trainer / assessor" in header_text:

                for row in rows:

                    if len(row) < 4:
                        continue

                    code = row[0].strip()

                    if not re.fullmatch(
                        r"[A-Z]{3,}\d{3,}",
                        code
                    ):
                        continue

                    trainers.append(
                        {
                            "national_code": row[0],
                            "state_code": row[1],
                            "unit_title": row[2],
                            "trainer_assessor": row[3]
                        }
                    )

            #
            # NEW TEMPLATE
            #
            elif "trainer/assessors name" in header_text:

                for row in rows[2:]:

                    if len(row) < 3:
                        continue

                    code = row[0].strip()

                    if not re.fullmatch(
                        r"[A-Z]{3,}\d{3,}",
                        code
                    ):
                        continue

                    trainer = ""

                    if len(row) >= 3:
                        trainer = row[2]

                    trainers.append(
                        {
                            "national_code": code,
                            "unit_title": row[1],
                            "trainer_assessor": trainer
                        }
                    )

        return trainers

    # =====================================================
    # ASSESSMENT MATRIX
    # =====================================================

    def get_assessment_matrices(self):

        matrices = []

        for table in self.tables:

            rows = table["rows"]

            if len(rows) < 3:
                continue

            row1 = " ".join(rows[0]).lower()
            row2 = " ".join(rows[1]).lower()

            if (
                "national code" in row1
                and "unit of competency" in row1
            ):

                if (
                    "a" in row2
                    and "b" in row2
                    and "c" in row2
                    and "d" in row2
                ):
                    matrices.append(table)

        return matrices

    # =====================================================
    # ASSESSMENT CALENDAR
    # =====================================================

    def get_assessment_calendars(self):

        calendars = []

        for table in self.tables:

            rows = table["rows"]

            if len(rows) < 2:
                continue

            row1 = " ".join(rows[0]).lower()
            row2 = " ".join(rows[1]).lower()

            #
            # New template
            #
            if (
                (
                    "first semester" in row1
                    or "second semester" in row1
                )
                and "national code" in row2
            ):
                calendars.append(table)
                continue

            #
            # Old template
            #
            if (
                "national code" in row1
                and "wk1" in row1
            ):
                calendars.append(table)

        return calendars
        
    def get_cell_borders(self, cell):

        tc_pr = cell._tc.tcPr

        if tc_pr is None:
            return {}

        borders = tc_pr.find(qn("w:tcBorders"))

        if borders is None:
            return {}

        result = {}

        for side in ["top", "bottom", "left", "right"]:

            border = borders.find(qn(f"w:{side}"))

            if border is None:

                result[side] = None

            else:

                result[side] = border.get(
                    qn("w:val")
                )

        return result
        
    def get_cluster_state(self, cell):

        borders = self.get_cell_borders(cell)

        top = borders.get("top")
        bottom = borders.get("bottom")

        if top == "single" and bottom == "nil":
            return "start"

        if top == "nil" and bottom == "nil":
            return "middle"

        if top == "nil" and bottom == "single":
            return "end"

        if top == "single" and bottom == "single":
            return "standalone"

        return "unknown"
    
    def build_cluster_map(self, docx_table):

        cluster_map = {}

        current_cluster_parts = []

        current_rows = []

        for row_idx, row in enumerate(docx_table.rows):

            if len(row.cells) < 2:
                continue

            code = row.cells[1].text.strip()

            #
            # Only process actual unit rows
            #
            if not re.fullmatch(
                r"[A-Z]{3,}\d{3,}",
                code
            ):
                continue

            text = row.cells[0].text.strip()

            state = self.get_cluster_state(
                row.cells[0]
            )

            #
            # START OF CLUSTER
            #
            if state == "start":

                #
                # Flush previous unfinished cluster
                #
                if current_rows:

                    cluster_name = " ".join(
                        current_cluster_parts
                    ).strip()

                    for r in current_rows:
                        cluster_map[r] = cluster_name

                current_cluster_parts = []

                current_rows = [row_idx]

                if text:
                    current_cluster_parts.append(text)

            #
            # MIDDLE OF CLUSTER
            #
            elif state == "middle":

                current_rows.append(row_idx)

                if text:
                    current_cluster_parts.append(text)

            #
            # END OF CLUSTER
            #
            elif state == "end":

                #
                # End of multi-row cluster
                #
                if current_rows:

                    current_rows.append(row_idx)

                    if text:
                        current_cluster_parts.append(text)

                    cluster_name = " ".join(
                        current_cluster_parts
                    ).strip()

                    for r in current_rows:
                        cluster_map[r] = cluster_name

                    current_cluster_parts = []

                    current_rows = []

                #
                # Standalone row
                #
                else:

                    cluster_map[row_idx] = (
                        text if text else None
                    )

            #
            # STANDALONE CLUSTER
            #
            elif state == "standalone":

                cluster_map[row_idx] = (
                    text if text else None
                )

            #
            # UNKNOWN STATE
            #
            else:

                cluster_map[row_idx] = (
                    text if text else None
                )

        #
        # Flush unfinished cluster at end of table
        #
        if current_rows:

            cluster_name = " ".join(
                current_cluster_parts
            ).strip()

            for r in current_rows:
                cluster_map[r] = cluster_name

        return cluster_map      

    def make_cluster_key(self, cluster_name):

        cluster_name = html.unescape(cluster_name)

        cluster_name = re.sub(
            r'([a-zA-Z])(\d+)',
            r'\1-\2',
            cluster_name
        )

        cluster_name = cluster_name.lower()

        cluster_name = cluster_name.replace("&", "and")

        cluster_name = re.sub(
            r'[^a-z0-9]+',
            '-',
            cluster_name
        )

        cluster_name = re.sub(
            r'-+',
            '-',
            cluster_name
        )

        return cluster_name.strip("-")
    
    def make_folder_name(self, cluster_name):

        joining_words = {
            "and",
            "or",
            "of",
            "the",
            "for",
            "to",
            "in",
            "on",
            "at",
            "by",
            "with"
        }

        cluster_name = html.unescape(cluster_name)

        # Separate letters and numbers
        cluster_name = re.sub(
            r'([a-zA-Z])(\d+)',
            r'\1 \2',
            cluster_name
        )

        # Replace special chars with separator
        cluster_name = re.sub(
            r'[^A-Za-z0-9 ]+',
            ' - ',
            cluster_name
        )

        # Collapse multiples
        cluster_name = re.sub(
            r'(?:\s*-\s*)+',
            ' - ',
            cluster_name
        )

        cluster_name = ' '.join(cluster_name.split())

        words = []

        for i, word in enumerate(cluster_name.split()):

            if word == "-":
                words.append(word)
                continue

            lower = word.lower()
            
            # Preserve acronyms adn oddly cased words
            if any(c.isupper() for c in word[1:]):
                words.append(word)
                continue

            if (
                i > 0
                and lower in joining_words
            ):
                words.append(lower)
            else:
                words.append(lower.capitalize())

        return " ".join(words)
    
    def make_cluster_unit_id(self, cluster_key, state_code):

        parts = re.findall(
            r'(?:^|[-_])([a-z])',
            cluster_key
        )

        acronym = ''.join(parts).upper()

        return f"{acronym}-{state_code}"

    
    def build_relationships(self):

        self.clusters = {}
        self.cluster_units = {}
        self.qualification_clusters = {}

        qualification_code = self.qualification["qualification"]["state_code"]

        for unit in self.units:

            #
            # Qualification -> Unit
            #
            self.qualification_units[
                (qualification_code, unit["state_code"])
            ] = {
                "qualificationCode": qualification_code,
                "unitCode": unit["state_code"],
                "coreElective": unit["core_elective"],
                "electiveGroup": unit["elective_group"]
            }

            cluster_name = unit["cluster"]

            if not cluster_name:
                continue

            cluster_key = self.make_cluster_key(cluster_name)
            cluster_folder = self.make_folder_name(cluster_name)

            self.clusters[cluster_key] = {
                "clusterKey": cluster_key,
                "clusterName": cluster_name,
                "clusterFolder": cluster_folder
            }

            self.qualification_clusters[
                (qualification_code, cluster_key)
            ] = {
                "qualificationCode": qualification_code,
                "clusterKey": cluster_key,
                "stage": unit["stage"]
            }

            self.cluster_units[
                (cluster_key, unit["state_code"])
            ] = {
                "clusterUnitId": self.make_cluster_unit_id
                    (cluster_key, unit["state_code"]),
                "clusterKey": cluster_key,
                "stateCode": unit["state_code"]
            }            
            

    # =====================================================
    # MASTER OUTPUT
    # =====================================================

    def parse(self):
        
        self.qualification = self.get_qualification()
        self.delivery = self.get_delivery()
        self.units = self.get_units()        
        self.trainers = self.get_trainers()
        self.assessment_matrices = self.get_assessment_matrices()
        self.assessment_calendars = self.get_assessment_calendars()
        self.build_relationships()
        
        return self
    
    def to_lists(self):

        return {
            "template": self.template,
            "qualification": self.qualification,
            "delivery": self.delivery,
            "units": self.units,
            "clusters": list(self.clusters.values()),
            "qualification_clusters": list(
                self.qualification_clusters.values()
            ),
            "cluster_units": list(
                self.cluster_units.values()
            ),
            "qualification_units": list(
                self.qualification_units.values()
            ),
            "trainers":
                self.get_trainers()
         
        }