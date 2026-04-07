#!/usr/bin/env python3
"""
Test Microsoft Graph API Connection
"""

import requests
from msal import ConfidentialClientApplication
import configparser

# Load configuration
config = configparser.ConfigParser()
config.read('config.ini')

CLIENT_ID = config['Azure']['client_id']
CLIENT_SECRET = config['Azure']['client_secret']
TENANT_ID = config['Azure']['tenant_id']
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

def get_access_token():
    """Get access token"""
    app = ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    
    result = app.acquire_token_for_client(scopes=SCOPE)
    
    if "access_token" in result:
        return result['access_token']
    else:
        print(f"Failed to get token: {result}")
        return None

def test_api():
    """Test various API endpoints"""
    token = get_access_token()
    if not token:
        return
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    print("Testing Microsoft Graph API Connection...")
    print("=" * 50)
    
    # Test 1: Get users
    print("\n1. Testing user list access:")
    response = requests.get(
        f"{GRAPH_API_BASE}/users?$top=5",
        headers=headers
    )
    if response.status_code == 200:
        users = response.json().get('value', [])
        print(f"✓ Found {len(users)} users")
        for user in users[:3]:
            print(f"  - {user.get('userPrincipalName', 'Unknown')}")
    else:
        print(f"✗ Error: {response.status_code} - {response.text}")
    
    # Test 2: Get messages for a specific user
    if response.status_code == 200 and users:
        test_user = users[0]['userPrincipalName']
        print(f"\n2. Testing email access for {test_user}:")
        
        # Try simple query first
        response = requests.get(
            f"{GRAPH_API_BASE}/users/{test_user}/messages?$top=5",
            headers=headers
        )
        if response.status_code == 200:
            messages = response.json().get('value', [])
            print(f"✓ Found {len(messages)} messages")
            for msg in messages[:3]:
                print(f"  - {msg.get('subject', 'No subject')}")
        else:
            print(f"✗ Error: {response.status_code} - {response.text}")
    
    # Test 3: Search across all mailboxes
    print("\n3. Testing search across mailboxes:")
    search_query = "search=emails"
    response = requests.get(
        f"{GRAPH_API_BASE}/search/query",
        headers=headers,
        json={
            "requests": [
                {
                    "entityTypes": ["message"],
                    "query": {
                        "queryString": "test"
                    },
                    "from": 0,
                    "size": 5
                }
            ]
        }
    )
    if response.status_code == 200:
        print("✓ Search API accessible")
    else:
        print(f"✗ Error: {response.status_code} - {response.text}")

if __name__ == "__main__":
    test_api()