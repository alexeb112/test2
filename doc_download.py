from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import sys
import codecs
import urllib
sys.stdout=codecs.getwriter('utf-8')(sys.stdout)

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/documents.readonly']

# The ID of a sample document.
DOCUMENT_ID = '1jjgEQQPwLQAtD3hmFxhKaAW3BbebpyoLTUvGDuoW_6U'
heading_map = {'HEADING_1':'#','HEADING_2':'##','HEADING_3':'###','HEADING_4':'####','HEADING_5':'#####','HEADING_6':'######'}
def handle_paragraph(paragraph, is_toc = False):
  ret = ''
  p_style = paragraph['paragraphStyle']
  needs_newline = False
  if 'namedStyleType' in p_style and p_style['namedStyleType'] in heading_map:
    ret+= heading_map[p_style['namedStyleType']]
    needs_newline = True
  elif 'bullet' in paragraph:
    bullet = paragraph['bullet']
    if 'nestingLevel' not in bullet:
      ret+='1. '
    else:
      ret+= ('  '*bullet['nestingLevel'] + '1. ')
  for element in paragraph['elements']:
    if 'textRun' in element:
      text = element['textRun']['content']
      striped_text = text.strip()
      formatted_text = striped_text
      if text.isspace():
        ret+= text
        return ret
      text_style = element['textRun']['textStyle']
      if 'bold' in text_style and text_style['bold']:
        formatted_text = '**' + striped_text + '**'
      if 'underline' in text_style and text_style['underline']:
        formatted_text = '<ins>'+ striped_text +'</ins>'

      ret+=text.replace(striped_text, formatted_text)

      # TODO: add support for images
    elif 'pageBreak' in element:
      pass
    elif 'inlineObjectElement' in element:
      ret+='![]('+element['inlineObjectElement']['inlineObjectId']+')'
    else:
      print(element)
  if needs_newline:
    ret+='\n'
  return ret

def handle_table(table):
  ret = ''
  num_cols = len(table['tableStyle']['tableColumnProperties'])
  head = True
  for row in table['tableRows']:
    ret += '| '
    for cell in row['tableCells']:
      ret += handle_contents(cell['content']).strip()
      ret+= '| '
    ret += '\n'
    if head:
      ret += ('| '+ '--- |'*num_cols) + '\n'
      head = False

  return ret

def handle_contents(contents, is_toc = False):
  ret = ''
  for content in contents:
    if 'tableOfContents' in content:
      continue
    elif 'paragraph' in content:
      ret += handle_paragraph(content['paragraph'], is_toc)
    elif 'sectionBreak' in content:
      continue
      pass
    elif 'table' in content:
      ret += handle_table(content['table'])
    else:
      print('=========')
      print(content)
      break
  return ret

def main():
  """Shows basic usage of the Docs API.
  Prints the title of a sample document.
  """
  creds = None
  # The file token.pickle stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
      creds = pickle.load(token)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
       creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          'credentials.json', SCOPES)
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
      pickle.dump(creds, token)

  service = build('docs', 'v1', credentials=creds)

  # Retrieve the documents contents from the Docs service.
  document = service.documents().get(documentId=DOCUMENT_ID).execute()

  print(handle_contents(document.get('body')['content']))
  #print('done\n\n\n\n\n\n-----')
  #print(document['inlineObjects'])
  return
  for key, object in document['inlineObjects'].items():
    url = object['inlineObjectProperties']['embeddedObject']['imageProperties']['contentUri']
    urllib.urlretrieve(url, key)

if __name__ == '__main__':
    main()
