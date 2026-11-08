# check_storage.py
import json
import googleapiclient.discovery
from google.oauth2.service_account import Credentials

creds = Credentials.from_service_account_file(
    "creds.json",
    scopes=["https://www.googleapis.com/auth/drive"]
)

service = googleapiclient.discovery.build("drive", "v3", credentials=creds, cache_discovery=False)

# Get storage quota info
about = service.about().get(fields="storageQuota").execute()
quota = about["storageQuota"]
used = int(quota.get("usage", 0))
limit = int(quota.get("limit", 0))
print(f"Used: {used / 1e9:.2f} GB / {limit / 1e9:.2f} GB\n")

# List ALL files sorted by size
results = service.files().list(
    pageSize=100,
    orderBy="quotaBytesUsed desc",
    fields="files(id, name, mimeType, size, quotaBytesUsed)"
).execute()

for f in results.get("files", []):
    size = int(f.get("quotaBytesUsed", 0))
    if size > 0:
        print(f"  {size / 1e6:.1f} MB  |  {f['name']}  |  {f['mimeType']}")
