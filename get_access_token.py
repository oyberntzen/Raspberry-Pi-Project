from pyo365 import Account
from credentials import credentails

scopes = ['https://graph.microsoft.com/Mail.ReadWrite','offline_access']

account = Account(credentials, scopes=scopes)

url = account.connection.get_authorization_url()
print(url)
result_url = input('Paste the result url here...')
account.connection.request_token(result_url)
