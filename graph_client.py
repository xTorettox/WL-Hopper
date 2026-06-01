import requests

class GraphClient:
    def __init__(self, tenant_id, client_id, client_secret):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.token = self.get_token()

    def get_token(self):
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default"
        }

        r = requests.post(url, data=data)
        r.raise_for_status()
        return r.json()["access_token"]

    def search(self, query):
        url = "https://graph.microsoft.com/v1.0/search/query"

        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }

        body = {
            "requests": [
                {
                    "entityTypes": ["driveItem"],
                    "query": {
                        "queryString": query
                    }
                }
            ]
        }

        r = requests.post(url, headers=headers, json=body)
        r.raise_for_status()
        data = r.json()

        return data["value"][0]["hitsContainers"][0]["hits"]
