import msal
import os
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'
def generate_access_token(client_id, scopes, cache_file='token_cache.bin'):
   # Setup token cache
   cache = msal.SerializableTokenCache()
   if os.path.exists(cache_file):
       cache.deserialize(open(cache_file, 'r').read())
   app = msal.PublicClientApplication(
       client_id,
       authority="https://login.microsoftonline.com/common",
       token_cache=cache
   )
   # Try to get token silently (from cache)
   accounts = app.get_accounts()
   if accounts:
       result = app.acquire_token_silent(scopes, account=accounts[0])
   else:
       # Use device code flow if no cached token
       flow = app.initiate_device_flow(scopes=scopes)
       if "user_code" not in flow:
           raise Exception("Failed to start device flow")
       print("Go to", flow["verification_uri"], "and enter code:", flow["user_code"])
       result = app.acquire_token_by_device_flow(flow)
   # Save updated cache
   if cache.has_state_changed:
       with open(cache_file, 'w') as f:
           f.write(cache.serialize())
   if "access_token" in result:
       return result
   else:
       raise Exception("Failed to acquire token: " + str(result))