import json
from pathlib import Path

from lap_parser import LAPDoc

folder = Path(
    r"C:\Users\mbutt\OneDrive - TAFE\NMT Integrated Technologies and Automation (Teams) - KAD Docs\ICT60220 - AE08 - Advanced Diploma of Information Technology (Cyber Security)\C - Identity Management\2 KAD\Learning and Assessment Plan (F122A14).docx"
)

for docx_file in sorted(folder.glob("*.docx")):

    print(f"\nProcessing {docx_file.name}")

    result = (
        LAPDoc(str(docx_file))
        .parse()
        .to_lists()
    )

    print(
        json.dumps(
            result["lap"],
            indent=4,
            default=str
        )
    )

    print(
        f"Sessions: "
        f"{len(result['sessions'])}"
    )

    print(
        f"Assessments: "
        f"{len(result['assessments'])}"
    )

    print(
        f"Lecturers: "
        f"{len(result['lecturers'])}"
    )