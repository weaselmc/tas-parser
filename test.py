import json
from pathlib import Path

from lap_parser import LAPDoc

folder = Path(
    r"C:\Users\mbutt\OneDrive - TAFE\NMT Integrated Technologies and Automation (Teams) - KAD Docs\ICT60220 - AE08 - Advanced Diploma of Information Technology (Cyber Security)\C - Identity Management\2 KAD\\"
)

for docx_file in sorted(folder.glob("*.docx")):

    print(f"\nProcessing {docx_file.name}")

    lap = LAPDoc(str(docx_file))
    
    result = (lap       
        .parse()
        .to_lists()
    )

    lap.debug_tables()

    print(
        json.dumps(
            result["lap"],
            indent=4,
            default=str
        )
    )

    print(
            json.dumps(
                result["sessions"],
                indent=4,
                default=str
            )
        )
    
    print(
            json.dumps(
                result["assessments"],
                indent=4,
                default=str
            )
        )
    
    print(
            json.dumps(
                result["lecturers"],
                indent=4,
                default=str
            )
        )

    print(
                json.dumps(
                    result["sessionelementmappings"],
                    indent=4,
                    default=str
                )
            )