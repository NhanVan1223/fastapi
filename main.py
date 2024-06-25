from typing import List
from fastapi import FastAPI, HTTPException
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.listitems.caml.query import CamlQuery
import os

app = FastAPI()

# Define your SharePoint site URL and list name
SITE_URL = "https://viendaukhivn.sharepoint.com/sites/H2NH3DataSource"
LIST_NAME = "test"

# Use environment variables for client credentials
CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID", "your_client_id_here")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET", "your_client_secret_here")

def get_sharepoint_list_items():
    try:
        # Authenticate using client credentials
        ctx = ClientContext(SITE_URL).with_credentials(ClientCredential(CLIENT_ID, CLIENT_SECRET))

        # Load the web object
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()

        # Get the list
        sp_list = ctx.web.lists.get_by_title(LIST_NAME)
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
        raise HTTPException(status_code=500, detail=f"Error accessing SharePoint list: {str(e)}")

@app.get("/")
def read_items():
    items = get_sharepoint_list_items()
    return {"items": items}
