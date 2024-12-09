from flask import Flask, redirect, request, session, url_for
from requests_oauthlib import OAuth2Session
import os
import json

app = Flask(__name__) #creates Flask web app instace
app.secret_key = os.urandom(24) #sets 24-byte secrete key for session

# Quickbooks endpoints
AUTHORIZATION_BASE_URL = 'https://appcenter.intuit.com/connect/oauth2' #redirects to quickbooks authorization page
TOKEN_URL = 'https://oauth.platform.intuit.ocm/oauth2/v1/tokens/bearer' #exchange code for access token
API_BASE_URL = 'https://quickbooks.api.intuit.com/v3/company' #base api url for api requrests (PRODUCTION)
#API_BASE_URL = 'sandbox-quickbooks.api.intuit.com/v3/company'

# Credentials from Intuit app
CLIENT_ID = "ABrYy3xkpvI2SHZcuPEJiqAIYByYpp3gGZuz8cZ0ZFnjzkSRDL" #sandbox
CLIENT_SECRET = "pgtarhRAi2X60vppF7GdvjEe0hJVYwOihM7Tdw9Q" #sandbox
REDIRECT_URI = 'https://localhost:8000/callback'



#make OAuth2Session (from requests-oauth) lets us make a session with a token, state, and redirect_uri as parameters
def make_oauth_session(token=None, state=None, redirect_uri=None):
    return OAuth2Session(
        client_id=CLIENT_ID,
        token=token,
        state=state,
        redirect_uri=redirect_uri,
        scope=['come.intuit.quickbooks.accounting']
    )

@app.route('/')
def index():
    return 'Welcome to QuickBooks OAuth 2.0 integration! <a href="/login">Login with QuickBooks</a>'

#root url '/' displays welcome message with path to '/login'

@app.route('/login')
def login():
    oauth = make_oauth_session(redirect_uri=REDIRECT_URI)
    authorization_url, state = oauth.authorization_url(AUTHORIZATION_BASE_URL)
    session['oauth_state'] = state #prevents csrf somehow
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
    return redirect(url_for('.profile'))

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

if __name__ == '__main__':
    app.run(port=8000, debug=True)









