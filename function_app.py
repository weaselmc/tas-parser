

import traceback
import json
import base64

try:
    import tas_parser
    IMPORT_OK = True
except Exception as ex:
    IMPORT_OK = False
    IMPORT_ERROR = traceback.format_exc()

import azure.functions as func
from tas_parser import TASDoc

app = func.FunctionApp()
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

        result = TASDoc(file_bytes).parse().to_lists()         

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
    
@app.route(route="test")
def test(req: func.HttpRequest) -> func.HttpResponse:

    if IMPORT_OK:
        return func.HttpResponse(f"Import OK.")
    else:
        return func.HttpResponse(IMPORT_ERROR)