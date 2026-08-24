import json

from get_uocinfo import UocParser


parser = UocParser()

result = parser.extract("ICTNWK536")

print(
    json.dumps(
        result,
        indent=4,
        default=str
    )
)