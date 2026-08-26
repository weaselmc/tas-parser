import json

from tp_parser import TPParser

parser = TPParser()

result = parser.extract("22063VIC")

print(
    json.dumps(
        result,
        indent=4,
        default=str
    )
)