from typing import List
from fastapi import FastAPI, HTTPException
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.listitems.caml.query import CamlQuery

app = FastAPI()

# Define your SharePoint site URL and list name
site_url = "https://viendaukhivn.sharepoint.com/sites/H2NH3DataSource"
list_name = "test"

# Client credentials
client_id = "15cf6582-65f2-4523-b3dd-d360f39bddf7"
client_secret = "HpzgMnjfeKeOF6OZI+vIsPTi0SOdeisfeSLgPsREvAQ="

def get_sharepoint_list_items():
    try:
        # Authenticate using client credentials
        ctx = ClientContext(site_url).with_credentials(ClientCredential(client_id, client_secret))

        # Load the web object
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()

        # Get the list
        sp_list = ctx.web.lists.get_by_title(list_name)
        ctx.load(sp_list)
        ctx.execute_query()

        # Query items in the list
        caml_query = CamlQuery()
        items = sp_list.get_items(caml_query)
        ctx.load(items)
        ctx.execute_query()

        item_list = []
        for item in items:
            item_list.append(item.properties)

        return item_list
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    return {"message": "Hello World"}

@app.get("/items")
def read_items():
    items = get_sharepoint_list_items()
    return {"items": items}
