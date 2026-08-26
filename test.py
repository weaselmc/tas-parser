import json

from tp_parser import TPParser

parser = TPParser()

result = parser.extract("ICTNWK536")

print(
    json.dumps(
        result,
        indent=4,
        default=str
    )
)