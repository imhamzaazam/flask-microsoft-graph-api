"""Flask-OAuthlib sample for Microsoft Graph"""
# Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
# See LICENSE in the project root for license information.
import uuid
from bs4 import BeautifulSoup
import flask
from flask_oauthlib.client import OAuth
from azure.cosmos import cosmos_client
import urllib3
import config

# CFG = {
#     "ENDPOINT": "https://kraken-db.documents.azure.com:443",
#     "PRIMARYKEY": "iBLlh7E6VWW0nx2F9xXI82jXHsDcJFWQT3QU78Hq9gQXca1qgOSohSiv1U1JghMntK5J84SslorsVMaKjtt4yw==",
#     "DATABASE": "kraken",  # Prolly looks more like a name to you
#     "CONTAINER": "kraken-db"  # Prolly looks more like a name to you
# }
CFG = {
    "ENDPOINT": "https://localhost:8081",
    "PRIMARYKEY": "C2y6yDjf5/R+ob0N8A7Cgv30VRDJIWEHLM+4QDU5DE2nQ9nDuVTqobD4b8mGGyPMbIZnqyMsEcaGQy67XIw/Jw==",
    "DATABASE": "CosmosDatabase",  # Prolly looks more like a name to you
    "CONTAINER": "CosmosContainer"  # Prolly looks more like a name to you
}
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


APP = flask.Flask(__name__, template_folder='static/templates')

CONTAINER_LINK = f"dbs/{CFG['DATABASE']}/colls/{CFG['CONTAINER']}"
FEEDOPTIONS = {}
FEEDOPTIONS["enableCrossPartitionQuery"] = True

# Initialize the Cosmos client
client = cosmos_client.CosmosClient(
    url_connection=CFG["ENDPOINT"], auth={"masterKey": CFG["PRIMARYKEY"]}
)

APP.debug = True
APP.secret_key = 'development'
OAUTH = OAuth(APP)
MSGRAPH = OAUTH.remote_app(
    'microsoft', consumer_key=config.CLIENT_ID, consumer_secret=config.CLIENT_SECRET,
    request_token_params={'scope': config.SCOPES},
    base_url=config.RESOURCE + config.API_VERSION + '/',
    request_token_url=None, access_token_method='POST',
    access_token_url=config.AUTHORITY_URL + config.TOKEN_ENDPOINT,
    authorize_url=config.AUTHORITY_URL + config.AUTH_ENDPOINT)

@APP.route('/')
def homepage():
    """Render the home page."""
    return flask.render_template('homepage.html', sample='Flask-OAuthlib')

@APP.route('/login')
def login():
    """Prompt user to authenticate."""
    flask.session['state'] = str(uuid.uuid4())
    return MSGRAPH.authorize(callback=config.REDIRECT_URI, state=flask.session['state'])

@APP.route('/login/authorized')
def authorized():
    """Handler for the application's Redirect Uri."""
    if str(flask.session['state']) != str(flask.request.args['state']):
        raise Exception('state returned to redirect URL does not match!')
    response = MSGRAPH.authorized_response()
    flask.session['access_token'] = response['access_token']
    return flask.redirect('/graphcall')

@APP.route('/graphcall')
def graphcall():
    """Confirm user authentication by calling Graph and displaying some data."""
    first_pass = True
    ctr = 0
    count = -1
    while first_pass: 
        endpoint_l = 'me/messages?$select=subject,body,sender&$count=true&skip={}'.format(ctr)
        ep_2 = 'me/'
        headers = {'SdkVersion': 'sample-python-flask',
                    'x-client-SKU': 'sample-python-flask',
                    'client-request-id': str(uuid.uuid4()),
                    'return-client-request-id': 'true'}
        graphdata = MSGRAPH.get(endpoint_l, headers=headers).data
        emaildata = MSGRAPH.get(ep_2, headers=headers).data
        #count = graphdata[key][2]
        for key in graphdata:
            if count == -1 :
                if(key=="@odata.count"):
                    count = graphdata["@odata.count"]
                    
                    print("COUNTER: ",count)
                #print(ctr)
            if(key == "value"):
                length = len(graphdata[key])
                for i in range(length):
                    #print("Subject: ", graphdata[key][i]["subject"])
                    #print("Email: ", graphdata[key][i]["sender"])
                    #soup = BeautifulSoup(graphdata[key][i]["subject"])
                    #if soup is not None:
                    subject = graphdata[key][i]["subject"]
                    sender = graphdata[key][i]["sender"]["emailAddress"]["name"]
                    data = graphdata[key][i]["body"]["content"]
                    soup = BeautifulSoup(data)
                    #print("Body: ",  soup.get_text(" "))
                    #print("Sender: ", sender)
                    #temp = graphdata[key][i]["subject"]
                    temp = {
                        "Email": emaildata, 
                        "Subject": subject,
                        "Body": soup.get_text(" "), 
                        "sender": sender,

                    }
                    #client.UpsertItem(CONTAINER_LINK, temp)
                    client.CreateItem(CONTAINER_LINK, temp)
                    #print("Body: ", type(graphdata[key][i]["body"]["content"]))
        #print(count)
        if ctr > count:
            first_pass = False 
        
        ctr += 10
    return flask.render_template('graphcall.html',
                                 graphdata=graphdata,
                                 endpoint_l=config.RESOURCE + config.API_VERSION + '/' + endpoint_l,
                                 sample='Flask-OAuthlib')

@MSGRAPH.tokengetter
def get_token():
    """Called by flask_oauthlib.client to retrieve current access token."""
    return (flask.session.get('access_token'), '')

if __name__ == '__main__':
    APP.run()
# if __name__ == '__main__':
    # app.run(host='0.0.0.0', port=80)
