import json
import base64
import azure.functions as func

from tas_parser import TASDoc
from lap_parser import LAPDoc
from tp_parser import TPParser


app = func.FunctionApp()


# =====================================================
# RESPONSE HELPERS
# =====================================================

def success_response(
    data,
    status_code=200
):

    return func.HttpResponse(
        json.dumps(
            {
                "success": True,
                "data": data
            },
            default=str
        ),
        mimetype="application/json",
        status_code=status_code
    )


def error_response(
    code,
    message,
    status_code=500
):

    return func.HttpResponse(
        json.dumps(
            {
                "success": False,
                "error": {
                    "code": code,
                    "message": message
                }
            }
        ),
        mimetype="application/json",
        status_code=status_code
    )

# =====================================================
# TAS
# =====================================================

@app.route(
    route="parse-tas",
    auth_level=func.AuthLevel.FUNCTION
)
def parse_tas(
    req: func.HttpRequest
) -> func.HttpResponse:

    try:

        body = req.get_json()

        file_content = body.get(
            "fileContent"
        )

        if not file_content:

            return error_response(
                "INVALID_REQUEST",
                "fileContent is required",
                400
            )

        file_bytes = base64.b64decode(
            file_content
        )

        result = (
            TASDoc(file_bytes)
            .parse()
            .to_lists()
        )

        return success_response(
            result
        )

    except Exception as ex:

        return error_response(
            "TAS_PARSE_ERROR",
            str(ex)
        )


# =====================================================
# LAP
# =====================================================

@app.route(
    route="parse-lap",
    auth_level=func.AuthLevel.FUNCTION
)
def parse_lap(
    req: func.HttpRequest
) -> func.HttpResponse:

    try:

        body = req.get_json()

        file_content = body.get(
            "fileContent"
        )

        if not file_content:

            return error_response(
                "INVALID_REQUEST",
                "fileContent is required",
                400
            )

        file_bytes = base64.b64decode(
            file_content
        )

        result = (
            LAPDoc(file_bytes)
            .parse()
            .to_lists()
        )

        return success_response(
            result
        )

    except Exception as ex:

        return error_response(
            "LAP_PARSE_ERROR",
            str(ex)
        )


# =====================================================
# TP INFO
# =====================================================

@app.route(
    route="tpinfo",
    methods=["POST"],
    auth_level=func.AuthLevel.FUNCTION
)
def tp_info(
    req: func.HttpRequest
) -> func.HttpResponse:

    try:

        body = req.get_json()

        codes = body.get(
            "codes",
            []
        )

        if not codes:

            return error_response(
                "INVALID_REQUEST",
                "codes array is required",
                400
            )

        parser = TPParser()

        results = []
        errors = []

        for code in codes:

            try:

                results.append(
                    parser.extract(code)
                )

            except Exception as ex:

                errors.append(
                    {
                        "code": code,
                        "message": str(ex)
                    }
                )

        return success_response(
            {
                "results": results,
                "errors": errors
            }
        )

    except Exception as ex:

        return error_response(
            "TPINFO_ERROR",
            str(ex)
        )


# =====================================================
# HEALTH CHECK
# =====================================================

@app.route(
    route="health",
    auth_level=func.AuthLevel.FUNCTION
)
def health(
    req: func.HttpRequest
) -> func.HttpResponse:

    return success_response(
        {
            "service": "tas-parser",
            "status": "healthy"
        }
    )