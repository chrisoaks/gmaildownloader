
from __future__ import print_function
import httplib2
import os
import base64
import time

from apiclient import errors
from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/gmail.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'


def GetAttachments(service, user_id, msg_id, prefix=""):
  """Get and store attachment from Message with given id.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    msg_id: ID of Message containing attachment.
    store_dir: The directory used to store attachments.
  """
  
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()
    if (message["labelIds"] != [u'CHAT']):
      print (message["labelIds"])
      try:
         for part in message['payload']['parts']:
           if part['filename']:
             print (part['filename'])
             print (part['body']['attachmentId'])
             print (msg_id)
             if 'data' in part['body']:
               data=part['body']['data']
             else:
               att_id=part['body']['attachmentId']
               att=service.users().messages().attachments().get(userId=user_id, messageId=msg_id,id=att_id).execute()
               data=att['data']
             file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
             path = "C:/Users/Chris/Downloads/" + part['filename']

             with open(path, 'wb') as f:
               f.write(file_data)
           
      except KeyError:
        print ("keyerror with " + msg_id)
  except errors.HttpError:
    print('An error occurred: %s' % error)
def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'gmail-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def ListMessagesMatchingQuery(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  print("1")
  try:
    response = service.users().messages().list(userId=user_id,
                                               maxResults=10,
                                               q='has:attachment').execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])
    i = 0
    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q='has:attachment',
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])
      print("2")
      

    return messages[:10]
  except errors.HttpError, error:
    print('An error occurred:')

def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)

    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    while True:
        
        messages = ListMessagesMatchingQuery(service,'me')

        #print(messages)
        for dictionaries in messages:
            
            GetAttachments(service, 'me', dictionaries[u'id'],r'C:\Users\Steve\Desktop\Incoming')
        time.sleep(120)
        if not labels:
            print('No labels found.')
        else:
          print('Labels:')
          for label in labels:
            print(label['name'])

if __name__ == '__main__':
    main()

