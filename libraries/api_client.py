import requests
import json
import logging

class WebexAPI:
    def __init__(self, token, org_id, api_logger):
        self.token = token
        self.org_id = org_id
        self.base_url = "https://webexapis.com/v1"
        self.api_logger = api_logger
    
    def call(self, method, endpoint, data=None, params=None):
        url = f"{self.base_url}/{endpoint}"
        headers = {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json"
        }
        
        self.api_logger.info(f"API Call: {method} {url}")
        if params:
            self.api_logger.info(f"Params: {json.dumps(params)}")
        if data:
            self.api_logger.info(f"Data: {json.dumps(data)}")
        
        try:
            response = requests.request(method, url, headers=headers, json=data, params=params)
            self.api_logger.info(f"Response Status: {response.status_code}")
            self.api_logger.info(f"Response: {response.text}")
            
            if response.status_code in [200, 201, 204]:
                return response.json() if response.text else {}
            else:
                self.api_logger.error(f"API Error: {response.status_code} - {response.text}")
                return {"error": response.text, "status_code": response.status_code}
        except Exception as e:
            self.api_logger.error(f"Exception during API call: {e}")
            return {"error": str(e)}
