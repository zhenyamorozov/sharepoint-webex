import os
import asyncio
from urllib.parse import urlparse, quote

from dotenv import load_dotenv

from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient
from msgraph.generated.sites.item.lists.item.items.items_request_builder import ItemsRequestBuilder

# Load environment variables from .env file
load_dotenv()

class SharePointItem(dict):
    def __init__(self, graph_item, sp_client, site_id, list_id, loop):
        # Initialize as dictionary with all field data
        super().__init__(graph_item.fields.additional_data)
        self._graph_item = graph_item
        self._sp_client = sp_client
        self._site_id = site_id
        self._list_id = list_id
        self._loop = loop
        self._original_data = dict(self)
    
    def save(self):
        """Save any changes made to this item"""
        # Find what changed
        changes = {k: v for k, v in self.items() if k not in self._original_data or self._original_data[k] != v}
        
        if changes:
            self._loop.run_until_complete(self._update_async(changes))
            self._original_data.update(changes)
    
    async def _update_async(self, changes):
        from msgraph.generated.models.field_value_set import FieldValueSet
        field_value_set = FieldValueSet()
        field_value_set.additional_data = changes
        await self._sp_client.sites.by_site_id(self._site_id).lists.by_list_id(self._list_id).items.by_list_item_id(self._graph_item.id).fields.patch(field_value_set)

class SP:

    def __init__(self, sp_site_url, sp_list_name, sp_folder_name):
        
        credentials = ClientSecretCredential(
            tenant_id=f"{os.getenv('SHAREPOINT_TENANT_ID', '')}.onmicrosoft.com",
            client_id=os.getenv("SHAREPOINT_CLIENT_ID", ""),
            client_secret=os.getenv("SHAREPOINT_CLIENT_SECRET", "")
        )

        self.client = GraphServiceClient(credentials=credentials)
        
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)
        self.loop.run_until_complete(self._initialize(sp_site_url, sp_list_name, sp_folder_name))

    def __del__(self):
        try:
            if hasattr(self, 'client') and self.client:
                # Check if event loop is still running
                try:
                    loop = asyncio.get_event_loop()
                    if loop.is_running():
                        # Schedule cleanup for later
                        loop.call_soon_threadsafe(self.client.close)
                    else:
                        # Run cleanup now
                        asyncio.run(self.client.close())
                except RuntimeError:
                    # Event loop is closed, can't cleanup
                    pass
        except:
            # Ignore all cleanup errors during shutdown
            pass
    
    async def _initialize(self, sp_site_url, sp_list_name, sp_folder_name):
        
        parsed_site_url = urlparse(sp_site_url)
        site_hostname = parsed_site_url.netloc
        site_path = parsed_site_url.path
        
        site_id = f"{site_hostname}:{site_path}"
        self.site = await self.client.sites.by_site_id(site_id).get()
        
        lists = await self.client.sites.by_site_id(self.site.id).lists.get()
        self.list = next(l for l in lists.value if l.display_name == sp_list_name)
        
        # Get items with fields expanded
        request_config = ItemsRequestBuilder.ItemsRequestBuilderGetRequestConfiguration()
        request_config.query_parameters = ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters()
        request_config.query_parameters.expand = ["fields"]
        request_config.query_parameters.filter = "startswith(contentType/id, '0x0120')"

        items = await self.client.sites.by_site_id(self.site.id).lists.by_list_id(self.list.id).items.get(request_configuration=request_config)
        
        # Filter for folders and find the specific folder
        folders = [item for item in items.value if item.content_type and item.content_type.id and item.content_type.id.startswith('0x0120')]
        self.folder = next(f for f in folders if f.fields and f.fields.additional_data.get('Title') == sp_folder_name)       

    async def _get_folder_items_async(self):
        # Microsoft Graph API limitation: Cannot directly filter SharePoint list items by folder
        # - No server-side folder filtering support for list items (unlike document libraries)
        # - webUrl field is not available for OData $filter operations
        # - FileDirRef and similar folder fields don't exist in SharePoint lists
        # Workaround: Get all items and filter client-side by URL path containing folder name
        
        # Get all items from the list with fields expanded
        request_config = ItemsRequestBuilder.ItemsRequestBuilderGetRequestConfiguration()
        request_config.query_parameters = ItemsRequestBuilder.ItemsRequestBuilderGetQueryParameters()
        request_config.query_parameters.expand = ["fields"]
        request_config.query_parameters.top = 5000  # Maximum allowed is usually 5000
        
        items = await self.client.sites.by_site_id(self.site.id).lists.by_list_id(self.list.id).items.get(request_configuration=request_config)
        
        # Return all non-folder items filtering by web_url
        folder_name = quote(self.folder.fields.additional_data.get('Title'))

        folder_items = [
            item for item in items.value
            if item.content_type 
            and not item.content_type.id.startswith('0x0120')  # Exclude folders
            and item.web_url
            and f"/{folder_name}/" in item.web_url  # URL-encoded folder name matching
        ]
        return folder_items
    
    def get_folder_items(self):
        items = self.loop.run_until_complete(self._get_folder_items_async())
        # Wrap each item in SharePointItem for easy dictionary access and saving
        return [SharePointItem(item, self.client, self.site.id, self.list.id, self.loop) for item in items]

    async def _get_list_columns_async(self):
        columns = await self.client.sites.by_site_id(self.site.id).lists.by_list_id(self.list.id).columns.get()
        return columns.value

    def get_list_columns(self):
        columns = self.loop.run_until_complete(self._get_list_columns_async())
        return {col.display_name: col.name for col in columns}



if __name__ == '__main__':
    sp = SP('https://cisco.sharepoint.com/sites/ASCandITCEcosystem', 'IPD Planning', 'FY26Q2 - IPD Week November')
    print(sp)
    folder_items = sp.get_folder_items()
    print([item['Title'] for item in folder_items])  # Now using dictionary access