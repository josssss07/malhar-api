from io import BytesIO
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from PIL import Image
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload
from pydantic import BaseModel
import qrcode
import uvicorn
import os
import hypercorn

app = FastAPI()

app.mount("/", StaticFiles(directory="static"), name="static")

cred_dict = {
  "type": "service_account",
  "project_id": "qr-system-392315",
  "private_key_id": "5bb704e5446d7c8a781f1b49c99835d473f3857c",
  "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvAIBADANBgkqhkiG9w0BAQEFAASCBKYwggSiAgEAAoIBAQDG93Vb7NqOGm9y\nRiB+v9mBMGMvT9azjiXo5qBSeOMyCNzukIWOJ3sqdh/MEjYKQVPLZApmWZ71t7Kv\nt86R9tmJWQmjPM/rn/RR6WA1XqpNszTzoldEfuJA8iYPZCKHY3n9DEz9WsMvqptL\n3Y9Gcq1E4hURTTVA3a7iXm3D+Bs8tQ5P49NOb8FEdqLBmxGWjlXB1Wm3rgBP68e0\neb/KAMASsNwuZTKgAWkcc5zUKrpszRwJ88D9nThT3yyrzYjqqdWpQ5B7I8Ihp3v7\nQW4aCNu14JOZm9nTSc7D9oTToVCB7Z+u1+hhqyqfIWTH2PGhOmKXs7wvGkPYAcrO\n2dZNpfdxAgMBAAECggEAHJrLBpJAwZlPWHB4j4BpdsdMGysxRNM05Az0E6dJkpHv\noWSprsxCStJ8s8wSdNr7fww86oYL8muchuK4EggZMkPYvN0rT0bJy1Tv/fxEI7OV\n80jtnu1W/dtSWXgd9rT4fsbb8rivwxSWCwwpYwltm6+dM8WT2GWGlQjaYwgyckGc\nimIDwqvn5cL/CKPSUXw9xNJo83b8u7HaDQTpsLuwihce+muI+Nec8ru06Oiy5nRk\ne0aebUOFp3eq020lmRk3WFB9/bEuIVTqY3CLYEFF+DPfocWBPYJiXIxrTYqq6BIk\n+ucMU//zbbTSBNnLEpMcjYo4Rt8AynnHFEhcpMsGWQKBgQDmqK/+e8PV14W77kEb\nUp/3l3xfiO0I1v/ZHWczs2+WEjFl0IHs0srrzT3DN7nsdwI+WmnoriNEguyQ9D6C\nV7pWtyl0ZizdcCuCwSdo4AVXAp85z+TmEvQ3JTHzbtQrXIDEwN4woq1db4qBlcji\nwWR7ZpdAcpjc3JRpk30Eh5OIGwKBgQDc02u+hkb+PI2nT5sXNvuXys/H49+p9rR0\n85a9RrrdAcHImNVup91gSC69t60S4yw4LBy8i6IV3G6bMsARU6kcw7MERou85t96\ntjDgMuCgGi5j564RLTB7HsYX8WHPRHrK0PTm/oIjG7Ghc7YuKis99OFrdXJ3u7lq\nMk2GQ0hPYwKBgA2dDyYZ7+kDG2WIHleafM6EJlcNIdBPwDH0Rk8K3B8jW78Cm2h6\n2HcqWebXtyV5sKw1ceLFxFca0xOLFtVikCDmFpBEJ4m6QRcqh0TtU+NayEMFPiFG\nJRvMGjKf6+3TO9Zg+7wrHchR+P7E9GJNv7x4xZyMJvGCI1BH4C0QQ2jZAoGAQho/\nrH7BjaVOugDIakCZO60IWcvKzjP9fOzV+L7NfQ7jlUq6yS8Sw5mX9E6hToAWYpJB\n3+bklCqyKV8dw5IJk4S5d9QuTFPIHhGfu90/BG4R6PIXVyjx1Ek3/z7QipzzLPcz\n+FnDVwMZPz1GEbepzhgZynMT2uek7zczobdOVAUCgYAVWzWG/62Dq6Na0OW4F9q+\n3GrL2lt8rlkhHzK0PhKqoDXw/sa/L31Qz/5EUkNrFPbdtZdqUld2Ka7kbkUFz0FQ\nepaBeRpLPwt/exmMKB1YLXutEr6KE5CkTR0qfrXPnIyQNVLiUNYfBT/Yjwk/mXlV\nK7ZDLZoEM9sCSu4j2/jCXQ==\n-----END PRIVATE KEY-----\n",
  "client_email": "api-handler@qr-system-392315.iam.gserviceaccount.com",
  "client_id": "117656494409799298413",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/api-handler%40qr-system-392315.iam.gserviceaccount.com",
  "universe_domain": "googleapis.com"
}

credentials = ServiceAccountCredentials.from_json_keyfile_dict(cred_dict, scopes=["https://www.googleapis.com/auth/drive"])
# forms_service = build('forms', 'v1', credentials=credentials)
drive_service = build('drive', 'v3', credentials=credentials)
sheets_service = build('sheets', 'v4', credentials=credentials)

spreadsheet_id = '1SQ0XUe5gC26CVMm7Lmrk-uGZTZlfs6g-DcclqfB-FUI'

@app.get("/")
def hello():
    return {"message": "API is running", "status": 200}


@app.get("/test")
def hello():
    return {"message": "API is running - test", "status": 200}


@app.get("/generate-qr/normal")
def new_qr(response_id: str, ticket_type:str):
    print(os.getcwd())
    try:
        qr = qrcode.QRCode(box_size=14)
        qr_string = response_id
        img = Image.open('late-stag.jpg')
        if ticket_type == "Stag":
            print("Stag")
        elif ticket_type == "Couple":
            img = Image.open('late-couple.jpg')
        qr.add_data(qr_string)
        qr.make()
        img_qr = qr.make_image(fill_color="black", back_color="#E6E6FA")
        pos = (1440,90)
        img.paste(img_qr, pos)
        stream = BytesIO()
        img.save(stream, format='JPEG')
        file_metadata = {'name': f"{response_id}.jpeg", 'parents': ['1E-MPKuk-RNKYiBRuUQAOt5vXtRLFT4Si']}
        media = MediaIoBaseUpload(stream, mimetype='image/jpeg', )
        file = drive_service.files().create(body=file_metadata, media_body=media,
                                            fields='id').execute()
        # print(F'File ID: {file.get("id")}')
        return {"message": "QR Generated successfully! ", "status": 200}
    except Exception as e:
        print(e)
        return {"message": "Unexpected error occurred", "status": 503}


@app.get("/generate-qr/early")
def new_qr(response_id: str, ticket_type:str):
    try:
        qr = qrcode.QRCode(box_size=14)
        qr_string = response_id
        img = Image.open(f'early-stag.jpg')
        if ticket_type == "Stag":
            print("Stag")
        elif ticket_type == "Couple":
            img = Image.open(f'early-couple.jpg')
        qr.add_data(qr_string)
        qr.make()
        img_qr = qr.make_image(fill_color="black", back_color="#E6E6FA")
        pos = (1440, 90)
        img.paste(img_qr, pos)
        stream = BytesIO()
        img.save(stream, format='JPEG')
        file_metadata = {'name': f"{response_id}.jpeg", 'parents': ['1E-MPKuk-RNKYiBRuUQAOt5vXtRLFT4Si']}
        media = MediaIoBaseUpload(stream, mimetype='image/jpeg', )
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
