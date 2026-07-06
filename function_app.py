
print("FUNCTION_APP_IMPORTING")
import azure.functions as func
import tas_parser

app = func.FunctionApp()

@app.route(route="hello")
def hello(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse("Hello World")