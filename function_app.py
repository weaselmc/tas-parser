

import traceback

try:
    import tas_parser
    IMPORT_OK = True
except Exception as ex:
    IMPORT_OK = False
    IMPORT_ERROR = traceback.format_exc()

import azure.functions as func

try:
    from docx import Document
    STATUS = "docx imported OK"
except Exception as ex:
    STATUS = str(ex)

app = func.FunctionApp()

@app.route(route="hello")
def hello(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse("Hello World")