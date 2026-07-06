import json
import base64

import azure.functions as func

print("FUNCTION_APP_IMPORTING")

from tas_parser import TASDoc

app = func.FunctionApp()
print("FUNCTION_APP_CREATED")

@app.route(
    route="parse-tas",
    auth_level=func.AuthLevel.FUNCTION
)
def parse_tas(req: func.HttpRequest) -> func.HttpResponse:

    try:

        body = req.get_json()

        file_name = body.get("fileName")

        file_content = body.get("fileContent")

        file_bytes = base64.b64decode(
            file_content
        )

        tas = TASDoc(file_bytes)

        result = tas.parse()

        return func.HttpResponse(
            json.dumps(
                result,
                default=str
            ),
            mimetype="application/json",
            status_code=200
        )

    except Exception as ex:

        return func.HttpResponse(
            str(ex),
            status_code=500
        )