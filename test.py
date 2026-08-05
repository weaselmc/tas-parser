import json
from pathlib import Path

from tas_parser import TASDoc

folder = Path(
    r"C:\Users\mbutt\OneDrive - TAFE\NMT Integrated Technologies and Automation (Teams) - TAS Docs"
)

for docx_file in sorted(folder.glob("*.docx")):

    print("\n" + "=" * 100)
    print(docx_file.name)
    print("=" * 100)

    tas = TASDoc(
        str(docx_file)
    )

    tas.parse()

    print(
        f"\nTEMPLATE: {tas.template}"
    )

    print(
        f"\nTemplate: {tas.template}"
    )

    print(
    f"Classification: "
    f"{tas.delivery_content.get('qualification_classification')}"
    )

    print(
        f"Enrolment Type: "
        f"{tas.delivery.get('enrolment_type')}"
    )

    print(
        f"Stages: "
        f"{tas.delivery.get('stages')}"
    )