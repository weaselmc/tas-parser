from docx import Document
from io import BytesIO
import re


class LAPDoc:

    def __init__(self, source):

        self.source = source

        if isinstance(source, bytes):
            self.source = BytesIO(source)

        self.doc = Document(self.source)

        self.tables = self._load_tables()

        self.lap = None
        self.sessions = []
        self.assessments = []
        self.lecturers = []

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

    def find_table_containing(self, search_text):

        search_text = search_text.lower()

        for table in self.tables:

            table_text = "\n".join(
                " ".join(row)
                for row in table["rows"]
            ).lower()

            if search_text in table_text:
                return table

        return None

    # =====================================================
    # LAP HEADER
    # =====================================================

    def get_lap(self):

        result = {
            "lapCode": None,
            "semester": None,
            "year": None,
            "lecturerName": None
        }

        full_text = "\n".join(
            " ".join(row)
            for table in self.tables
            for row in table["rows"]
        )

        #
        # Old template:
        # 2026 Semester 1
        #
        match = re.search(
            r"(20\d{2})\s+Semester\s+(\d)",
            full_text,
            re.IGNORECASE
        )

        if match:

            result["year"] = int(
                match.group(1)
            )

            result["semester"] = int(
                match.group(2)
            )

            return result

        #
        # New template:
        # Semester: 1 Year: 2026
        #
        year_match = re.search(
            r"Year\s*:?\s*(20\d{2})",
            full_text,
            re.IGNORECASE
        )

        semester_match = re.search(
            r"Semester\s*:?\s*(\d)",
            full_text,
            re.IGNORECASE
        )

        if year_match:

            result["year"] = int(
                year_match.group(1)
            )

        if semester_match:

            result["semester"] = int(
                semester_match.group(1)
            )

        return result

    # =====================================================
    # LECTURERS
    # =====================================================

    def get_lecturers(self):

        lecturers = []

        table = self.find_table_containing(
            "lecturer name"
        )

        if table is None:
            return lecturers

        rows = table["rows"]

        for row in rows:

            if len(row) < 5:
                continue

            lecturer_name = (
                row[0].strip()
            )

            if (
                not lecturer_name
                or "lecturer name"
                in lecturer_name.lower()
            ):
                continue

            lecturers.append({

                "lecturerName":
                    lecturer_name,

                "email":
                    row[2].strip(),

                "phone":
                    row[1].strip(),

                "campus":
                    None,

                "room":
                    row[4].strip(),

                "contact":
                    row[3].strip()
            })

        return lecturers

    # =====================================================
    # ASSESSMENTS
    # =====================================================

    def get_assessment_type(self, title):

        value = title.lower()

        if "portfolio" in value:
            return "Portfolio"

        if "project" in value:
            return "Project"

        if "knowledge" in value:
            return "Knowledge Assessment"

        if "observation" in value:
            return "Observation Checklist"

        if "third party" in value:
            return "Third Party Report"

        if "workplace" in value:
            return "Workplace Assessment"

        if "rpl" in value:
            return "Recognition of Prior Learning"

        return "Other"

    def get_assessments(self):

        assessments = []

        for table in self.tables:

            rows = table["rows"]

            for row in rows:

                if len(row) < 3:
                    continue

                first_col = row[0].strip()

                if not first_col.startswith(
                    "Assessment"
                ):
                    continue

                number_match = re.search(
                    r"(\d+)",
                    first_col
                )

                if not number_match:
                    continue

                assessment_number = int(
                    number_match.group(1)
                )

                due_week = None

                week_match = re.search(
                    r"Week\s+(\d+)",
                    row[2],
                    re.IGNORECASE
                )

                if week_match:

                    due_week = int(
                        week_match.group(1)
                    )

                title = row[1].strip()

                assessments.append({

                    "assessmentNumber":
                        assessment_number,

                    "assessmentTitle":
                        title,

                    "assessmentType":
                        self.get_assessment_type(
                            title
                        ),

                    "dueWeek":
                        due_week
                })

        return assessments

    # =====================================================
    # SESSION TABLE DETECTION
    # =====================================================

    def is_session_table(self, table):

        table_text = "\n".join(
            " ".join(row)
            for row in table["rows"]
        ).lower()

        #
        # Works for both old and new templates
        #
        return (
            "session" in table_text
            and (
                "learning resources" in table_text
                or "learning resources*" in table_text
            )
            and "topic" in table_text
        )

    # =====================================================
    # SESSIONS
    # =====================================================

    def get_sessions(self):

        sessions = []

        for table in self.tables:

            if not self.is_session_table(table):
                continue

            rows = table["rows"]

            for row in rows:

                #
                # Session number must be numeric
                #
                try:
                    session_number = int(
                        str(row[0]).strip()
                    )
                except Exception:
                    continue

                #
                # Contact Hours
                #
                try:
                    face_to_face_hours = float(
                        row[1] or 0
                    )
                except Exception:
                    face_to_face_hours = 0

                topic = (
                    row[3].strip()
                    if len(row) > 3
                    else ""
                )

                learning_resources = (
                    row[4].strip()
                    if len(row) > 4
                    else ""
                )

                out_of_class_activity = (
                    row[5].strip()
                    if len(row) > 5
                    else ""
                )

                try:
                    out_of_class_hours = float(
                        row[6] or 0
                    )
                except Exception:
                    out_of_class_hours = 0

                #
                # Amount Of Training categories
                #
                assessment_hours = 0

                if (
                    "assessment" in topic.lower()
                    or "submission" in topic.lower()
                ):
                    assessment_hours = (
                        face_to_face_hours
                    )

                sessions.append({

                    "sessionNumber":
                        session_number,

                    "topic":
                        topic,

                    "faceToFaceHours":
                        face_to_face_hours,

                    "workshopHours":
                        0,

                    "workPlacementHours":
                        0,

                    "outOfClassStudyHours":
                        out_of_class_hours,

                    "tutorialSupportHours":
                        0,

                    "assessmentHours":
                        assessment_hours,

                    "learningResources":
                        learning_resources,

                    "outOfClassActivity":
                        out_of_class_activity
                })

        return sessions

    # =====================================================
    # PARSE
    # =====================================================

    def parse(self):

        self.lap = self.get_lap()

        self.lecturers = (
            self.get_lecturers()
        )

        if self.lecturers:

            self.lap["lecturerName"] = (
                self.lecturers[0][
                    "lecturerName"
                ]
            )

        self.assessments = (
            self.get_assessments()
        )

        self.sessions = (
            self.get_sessions()
        )

        return self

    # =====================================================
    # OUTPUT
    # =====================================================

    def to_lists(self):

        return {

            "lap":
                self.lap,

            "sessions":
                self.sessions,

            "assessments":
                self.assessments,

            "lecturers":
                self.lecturers
        }


# =====================================================
# EXAMPLE USAGE
# =====================================================

if __name__ == "__main__":

    lap = (
        LAPDoc(
            "Learning and Assessment Plan (F122A14).docx"
        )
        .parse()
    )

    data = lap.to_lists()

    print("\n=== LAP ===")
    print(data["lap"])

    print("\n=== LECTURERS ===")
    for lecturer in data["lecturers"]:
        print(lecturer)

    print("\n=== SESSIONS ===")
    for session in data["sessions"]:
        print(session)

    print("\n=== ASSESSMENTS ===")
    for assessment in data["assessments"]:
        print(assessment)