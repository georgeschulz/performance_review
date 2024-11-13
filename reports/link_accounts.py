import os
from pyairtable import Api
import requests
import base64
from functools import wraps
from dotenv import load_dotenv
load_dotenv()

api = Api(os.getenv('SALES_BASE_KEY'))
airtable_table = api.table('appBsty6iukfNnuEK', 'tblenVnxR8q8iTGSk')

def ensure_valid_token(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except Exception as e:
            if hasattr(e, 'response') and e.response.status_code == 401:
                self.refresh()
                return func(self, *args, **kwargs)
            else:
                raise
    return wrapper

def link_accounts():
    class PestPacAPI:
        def __init__(self, username, password, client_id, client_secret, api_key, company_key):
            self.username = username
            self.password = password
            self.client_id = client_id
            self.client_secret = client_secret
            self.api_key = api_key
            self.company_key = company_key
            self.access_token = None
            self.refresh_token = None
            self.authenticate()
        
        def authenticate(self):
            auth_header = base64.b64encode(f"{self.client_id}:{self.client_secret}".encode('utf-8')).decode('utf-8')
            headers = {'Authorization': f'Bearer {auth_header}'}
            payload = {'grant_type': 'password', 'username': self.username, 'password': self.password}
            
            response = requests.post("https://is.workwave.com/oauth2/token?scope=openid", headers=headers, data=payload)
            if response.status_code == 200:
                response_json = response.json()
                self.access_token = response_json['access_token']
                self.refresh_token = response_json['refresh_token']
            else:
                raise Exception("Error getting token")

        def refresh(self):
            auth_header = base64.b64encode(f"{self.client_id}:{self.client_secret}".encode('utf-8')).decode('utf-8')
            headers = {'Authorization': f'Bearer {auth_header}'}
            payload = {'grant_type': 'refresh_token', 'refresh_token': self.refresh_token}
            
            response = requests.post("https://is.workwave.com/oauth2/token?scope=openid", headers=headers, data=payload)
            if response.status_code == 200:
                response_json = response.json()
                self.access_token = response_json['access_token']
            else:
                raise Exception("Error refreshing token")

        def _get_headers(self):
            return {
                'Authorization': f'Bearer {self.access_token}',
                'apikey': self.api_key,
                'tenant-id': self.company_key,
            }

        @ensure_valid_token
        def search_location(self, query):
            url = 'https://api.workwave.com/pestpac/v1/locations'
            params = {'q': query} if query else {}
            response = requests.get(url, headers=self._get_headers(), params=params)
            return response.json()

    pestpac = PestPacAPI(
        os.getenv('PESTPAC_USERNAME'),
        os.getenv('PESTPAC_PASSWORD'),
        os.getenv('PESTPAC_OUTH_CLIENT_ID'),
        os.getenv('PESTPAC_OUTH_SECRET'),
        os.getenv('PESTPAC_API_KEY'),
        os.getenv('PESTPAC_COMPANY_KEY')
    )

    airtable_records = airtable_table.all(formula="AND({PestPac Account} = '', NOT({Linkage Script Run} = 'Yes'), NOT({Close Status} = ''),  NOT({Close Status} = 'Open'), NOT({Close Status} = 'Disqualified: Exclude from pipeline'))")

    def check_phone_if_exists(phone):
        if phone:
            records = pestpac.search_location(query=phone)
            if len(records) == 0:
                return None
            return records[0]
        return None

    def check_email_if_exists(email):
        if email:
            records = pestpac.search_location(query=email)
            if len(records) == 0:
                return None
            return records[0]
        return None

    def check_name_if_exists(name):
        if name:
            records = pestpac.search_location(query=name)
            if len(records) == 0:
                return None
            return records[0]
        return None

    def find_account(phone, email, name):
        # Check phone first
        record = check_phone_if_exists(phone)
        
        # If phone check returns None, check email
        if record is None and email is not None:
            record = check_email_if_exists(email)
        
        # If email check returns None, check name
        if record is None and name is not None:
            record = check_name_if_exists(name)
        
        return record

    records = airtable_records
    for record in records:
        try:
            phone = record['fields'].get('Phone', None)
            email = record['fields'].get('Email', None)
            name = record['fields'].get('Customer Full Name', None)
            id = record['id']
            record = find_account(phone, email, name)
            location_id = None
            account_number = None
            if record is not None:
                location_id = record['LocationID']
                account_number = record['LocationCode']
            print(f"Linking: {phone}, {email}, {location_id}, {account_number}")
            update_fields = {'Linkage Script Run': 'Yes'}
            if location_id is not None:
                update_fields['LocationID'] = str(location_id)
            if account_number is not None:
                update_fields['PestPac Account'] = str(account_number)
            airtable_table.update(id, update_fields)
        except Exception as e:
            print(e)

