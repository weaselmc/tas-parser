import json

from tp_parser import TPParser
from tas_parser import TASDoc

#tas =TASDoc(r"C:\Users\mbutt\OneDrive - TAFE\NMT Integrated Technologies and Automation (Teams) - KAD Docs\22697VIC - BHE71 - Certificate IV in Integrated Technologies\01 TAS\22697VIC TAS.docx").parse()
#result = tas.to_lists()

parser = TPParser()
result = parser.extract("22697VIC")

print(
    json.dumps(
        result,
        indent=4,
        default=str
    )
)