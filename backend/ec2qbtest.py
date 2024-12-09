from flask import Flask, redirect, request, session, jsonify
from requests_oauthlib import OAuth2Session
import os
import json

app = Flask(__name__)
app.secret_key = os.urandom(24)

CLIENT_ID = "ABrYy3xkpvI2SHZcuPEJiqAIYByYpp3gGZuz8cZ0ZFnjzkSRDL"
CLIENT_SECRET = "pgtarhRAi2X60vppF7GdvjEe0hJVYwOihM7Tdw9Q"
REDIRECT_URI = 'https://3.142.73.158:8000/callback'
AUTHORIZATION_BASE_URL = 'https://appcenter.intuit.com/connect/oauth2'
TOKEN_URL = 'https://oauth.platform.intuit.com/oauth2/v1/tokens/bearer'
API_BASE_URL = 'https://sandbox-quickbooks.api.intuit.com/v3/company'
SCOPES = ['com.intuit.quickbooks.accounting']

def make_oauth_session(token=None, state=None, redirect_uri=None):
    return OAuth2Session(
        client_id=CLIENT_ID,
        token=token,
        state=state,
        redirect_uri=redirect_uri,
        scope=SCOPES
    )

@app.route('/login')
def login():
    oauth = make_oauth_session(redirect_uri=REDIRECT_URI)
    authorization_url, state = oauth.authorization_url(AUTHORIZATION_BASE_URL)
    session['oauth_state'] = state
    return redirect(authorization_url)

@app.route('/callback')
def callback():
    oauth = make_oauth_session(state=session['oauth_state'], redirect_uri=REDIRECT_URI)
    token = oauth.fetch_token(
        TOKEN_URL,
        authorization_response=request.url,
        client_secret=CLIENT_SECRET
    )
    session['oauth_token'] = token
    print(f"Token is: {session['oauth_token']}")
    realm_id = request.args.get('realmId')
    session['realm_id'] = realm_id
    print(realm_id)
    return redirect('/get-reports')


@app.route('/view-reports')
def open_reports():
    url = 'https://sandbox.qbo.intuit.com/app/reports'
    return redirect(url)

@app.route('/get-reports')
def get_reports():
    print("test1")
    oauth = make_oauth_session(token=session['oauth_token'])
    realm_id = session.get('realm_id')  # Store and retrieve realm_id appropriately
    url = f"{API_BASE_URL}/{realm_id}/reports/BalanceSheet"
    print("test2")
    # Example of querying for specific accounts if the API supports such queries
    params = {
    'start_date': '2024-01-01',
    'end_date': '2024-08-06',
    'item_ids': '1150040000',  # Example item ID
    'accounting_method': 'Accrual'
    }

    response = oauth.get(url, params=params)
    print(response.status_code)
    print("test3")
    if (response.status_code == 200):
        print("test4")
        data = response.json()
        print(data)

        with open ('def-rev-report.json', 'w') as file:
            json.dump(data, file)
        return data
    else:
        return f"Error: {response.status_code} - {response.text}"


@app.route('/sales-transactions')
def get_sales_transactions():
    print("Test")
    oauth = make_oauth_session(token=session['oauth_token'])
    realm_id = session.get('realm_id')  # Store and retrieve realm_id appropriately
    print("Test2")

    query = "SELECT * FROM Invoice"
    url = f"{API_BASE_URL}/{realm_id}/query"
    response = oauth.get(url, params={'query': query})
    print(response.status_code)
    print("Test3")
    content_type = response.headers.get('Content-Type')
    print(f'The content type is: {content_type}')

    if response.status_code == 200:
        print("test4")
        sales_transactions = response.text#response.json().get('QueryResponse', {}).get('Invoice', [])
        print("Test5")
        print(sales_transactions)
        return jsonify(sales_transactions)
    else:
        return f"Error: {response.status_code} - {response.text}"

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8000, debug=True, ssl_context=('cert.pem', 'key.pem'))