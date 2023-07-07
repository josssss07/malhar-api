from io import BytesIO
from fastapi import FastAPI
import json
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload
from pydantic import BaseModel
import qrcode
import uvicorn
import hypercorn

app = FastAPI()

json_file_path = 'xaveirs.json'

# with open(json_file_path, 'r') as j:
#     contents = json.loads(j.read())
    # print(contents)

cred_dict = {
  "type": "service_account",
  "project_id": "ecc-app-388505",
  "private_key_id": "6e31801f545801c79d8ae2bff5dc7ca75f3f060b",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCuANlLQdrAIMb7\noCtHIKeYRgoxPP+xjkW5IRO3YQekCEf6qzLZ7CsDKxofXlrGmJLdh7q/jRIZRuVn\nKcXv+RGhLgIk2VFZKy6mnSRDjUdeHLi5a9qG1YpOWsRyHL1RARELtZB/xh3A5nHg\nqeX2DcjmjkDq24qscwFh04fni60nIDD5gtIS8/xD5+CQMOZrF8hofkCTlW3eGWCU\nbJXG9nNfnQQNzRfcZQ7W7cJ/v1ijWUoFB1t6kKRo58EXMZmWEpwXEsaofr4TkiIw\nZoaM1QkbX65CutzFNifdiWumE3q8uOVSQv5hOxrwFc51SL/RgosvyZU+i2xx8rNR\nUyGadc7zAgMBAAECggEAFL5rKUqnjoIZ1sOohSlCcaff7TDNdth0PXbRB+qcY5TN\nJ/gi7tH16sHHsneoMMbds2VMASwLPVIzZRIY09wImwOGI+h4hz3bqOdQ/HCqUnDw\nIMLX4X0xqNevVb9RXofXBWNW37S5tVaDuvnmPWg1TC8nVBhqNtbbElOKfzMqqGA+\nPM66CYoDepVNMc033Nd8qy38sgFNue6nGYIvk/7HvsIOPtujFRY87Pg7iUV+5eXu\nMRSCE4c/OI0b2BNE+aERw4YJBU6FYtH0Bbku0SY+BgmXF1IKKtOi6/m986uOKwFx\ntq8N1DHiRIwxFQ1yXy0tq+j9HUqjwxXKCsmeRwh+KQKBgQDz5SnTTz+9CJ7mg/NA\nlq731unKpo0a3rc+rFtCMqosDE2AAFYe3uIxB7DfnzCZpwsKz1Rf4/LWRNIB2Smm\nHLeLm4dsPOV2Lnl7aKJi1P09eC+wJPM1iy5L2wJ2DlUhpyg2S5pinbk6u949RwZG\nLqluNtag4g77QSK7jlFYmiNBDwKBgQC2o6q36nbonJtj6Tu5v9qPq5t6BZti6gHu\nQx1hqaZLiNtLiMRE82BD5rnOWWWvT4D/k/n7OgrK+CKUDQJx2v8kpZ9Sk/s6ary/\nj9xeW7QKzQrMk8QADzZMBDWSO6edHuImIDlAzCJZGJSRoaTTEQ34NMpsHIj9X0oE\nQml0UBUL3QKBgB+h9U7G1Il7+MAFFSOnZ4IVibUS7PIzOKkUSbHISHH2FocnhAP0\n/HlHehVG3FLRa4k3YsYdFn3b5zD+LXyx9MxSm7naHBi75l2vMICJB19VmznJURH8\nv0BvY15UdY4r0/dWzutDcabAfw6Li7DGlIsK6cNsDm1gAVU6HCrVswTZAoGAC8s1\n0vqJAyxZvHHrMWt6KZzjRVXnWtPRnBkpZI0X9/i2cII8alds9/WGOhT7w/5WTiq4\nEckFuWWclgLhDYkewLcROrvjlTofRh98E3vIfIoREHTaS0awMuhyrSa9BCBiaiPa\njpyf+zDjJyRRCHApfsWp7KFLF1F37h57wM8LDOECgYEAlHTO12K31WusrooOoPBe\n4owI3720EEo330ytNYp1L20rlUH4D1wI/UG+qcZVXg9ehu/OjXkCBQHcpGAuKzHx\nWIzvG7Y8nvK3Sv5IHXDjjY9B3BMOjTYxG1LdskdW56mdiyICUoitsDqNKnvmMBvX\nQsWDtvRNaeDVgeRI5XfvLrs=\n-----END PRIVATE KEY-----\n",
  "client_email": "app-helper@ecc-app-388505.iam.gserviceaccount.com",
  "client_id": "109978774170613118113",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/app-helper%40ecc-app-388505.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

credentials = ServiceAccountCredentials.from_json_keyfile_dict(cred_dict, scopes=["https://www.googleapis.com/auth/drive"])
# forms_service = build('forms', 'v1', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)
sheets_service = build('sheets', 'v4', credentials=credentials)

spreadsheet_id = '1A-HPU-VjE0fTXPZqw2n1BbVp1vumFyX8TG1mREZCsKw'

@app.get("/")
def hello():
    return {"message": "API is running", "status": 200}


@app.get("/test")
def hello():
    return {"message": "API is running", "status": 200}


@app.get("/generate-qr")
def new_qr(response_id: str):
    try:
        qr_string = response_id
        img = qrcode.make(qr_string)
        stream = BytesIO()
        img.save(stream)
#         16Gjz0uD3LhVcv5gW_0NeGhjf30ZWFJnn
        file_metadata = {'name': f"{response_id}.png", 'parents': ['14wI-x6x146_WZjUrmLu15mFVTgX9g_aN']}
        media = MediaIoBaseUpload(stream, mimetype='image/png', )
        file = drive_service.files().create(body=file_metadata, media_body=media,
                                            fields='id').execute()
        # print(F'File ID: {file.get("id")}')
        return {"message": "QR Generated successfully! ", "status": 200}
    except:
        return {"message": "Unexpected error occurred", "status": 503}


@app.get("/responses")
def getresponses(event_id: str):
    try:
        batch_get_values_by_data_filter_request_body = {
            'value_render_option': 'FORMATTED_VALUE',
            'data_filters': [
                {
                    'gridRange': {
                        # 'sheetId': '1497832988',
                        'sheetId': event_id,
                    },
                },
            ],
        }

        data = sheets_service.spreadsheets().values().batchGetByDataFilter(spreadsheetId=spreadsheet_id,
                                                                           body=batch_get_values_by_data_filter_request_body).execute()
        # print(data['valueRanges'][0]['valueRange'])

        # ['Timestamp', 'Name', 'UID', 'Email_Address', 'Response ID', 'Attendance']
        headers = data['valueRanges'][0]['valueRange']['values'][0]

        values = data['valueRanges'][0]['valueRange']["values"]

        res = {"values": []}

        for value in values[1:]:
            value_dict = {}
            for i, value_key in enumerate(value):
                try:
                    value_dict[headers[i]] = value_key
                except:
                    value_dict[headers[i]] = ""
            res['values'].append(value_dict)

        return {"message": "Data Fetch Successful", "status": 200, "result": res}
    except:
        return {"message": "Invalid ID", "status": 403}


@app.get("/events/{dept}")
def get_events(dept: str):
    all_sheets = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    print(all_sheets['sheets'])

    events = []
    misc = []

    for sheet in all_sheets['sheets']:
        sheet_id = sheet['properties']['sheetId']
        sheet_title = sheet['properties']['title']
        event_json = {
            "name":sheet_title,
            "id":sheet_id
        }
        try:
            if str(sheet_title).split("-")[1].strip().casefold() == dept.casefold():
                events.append(event_json)
        except IndexError:
            misc.append(event_json)

    print(events)

    if dept.casefold() == 'misc'.casefold():
        result = misc
    else:
        result = events

    return {"message": "Events fetched", "result": result, "status": 200}


@app.get("/events")
def get_events():
    all_sheets = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    print(all_sheets['sheets'])

    events = []
    for sheet in all_sheets['sheets']:
        sheet_id = sheet['properties']['sheetId']
        sheet_title = sheet['properties']['title']
        event_json = {
            "name":sheet_title,
            "id":sheet_id
        }
        events.append(event_json)

    print(events)
    return {"message": "Events fetched", "result": events, "status": 200}


class ColumnData(BaseModel):
    col_name: str
    data: list


@app.put("/responses/by/col")  # TODO: get responses by col
def update_sheet_by_col(event_id: str, values: ColumnData):
    # Find desired column
    try:
        batch_get_values_by_data_filter_request_body = {
            'value_render_option': 'FORMATTED_VALUE',
            'data_filters': [
                {
                    'gridRange': {
                        # 'sheetId': '1497832988',
                        'sheetId': event_id,
                    },
                },
            ],
        }

        data = sheets_service.spreadsheets().values().batchGetByDataFilter(spreadsheetId=spreadsheet_id,
                                                                           body=batch_get_values_by_data_filter_request_body).execute()

        headers = data['valueRanges'][0]['valueRange']['values'][0]
        col_to_update_index = 0

        for x in headers:
            if x == values.col_name:
                col_to_update_index = headers.index(x)

        # col_to_update_index_letter = get_column_letter(col_to_update_index+1)
        print(col_to_update_index)

        # First column is timestamp which would never be changed, hence raise exception if index of col to update is 0
        if col_to_update_index == 0:
            raise Exception()
    except:
        return {"message": f"Error: Column not found in sheet {event_id}", "status": 403}

    # Update attendance column
    batch_update_values_by_data_filter_request_body = {
        'value_input_option': 'USER_ENTERED',
        'data': [
            {
                "dataFilter": {
                    "gridRange": {
                        "sheetId": event_id,
                        "startRowIndex": 1,
                        "startColumnIndex": col_to_update_index,
                        "endColumnIndex": col_to_update_index + 1,
                    }
                },
                "majorDimension": 'COLUMNS',
                "values": [
                    values.data
                ]
            }
        ],
    }

    request = sheets_service.spreadsheets().values().batchUpdateByDataFilter(spreadsheetId=spreadsheet_id,
                                                                             body=batch_update_values_by_data_filter_request_body).execute()

    return {"message": f"Updated sheet {event_id}", "Data received": values}
