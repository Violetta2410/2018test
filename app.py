import os.path
import pickle
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import json
import requests
import logging
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from googleapiclient.discovery import build
import datetime
import logging

# 로깅 설정
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Google OAuth 설정
SCOPES = [
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "openid",
    "https://www.googleapis.com/auth/calendar",
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/calendar.events.readonly",
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/contacts",
    "https://www.googleapis.com/auth/contacts.readonly",
    "https://www.googleapis.com/auth/tasks.readonly",
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# 포트 설정 (Google Cloud Console에 등록된 포트와 일치)
# 포트 설정
PORT = 8501

# Flow 생성 시 리디렉션 URI
flow = Flow.from_client_secrets_file(
    "credentials.json",
    scopes=SCOPES,
    redirect_uri=f"http://localhost:{PORT}/oauth2callback",
)

# 스코프 검증 비활성화 추가
flow.oauth2session.verify_scope = False


class OAuthHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        logger.info(
            "%s - - [%s] %s"
            % (self.address_string(), self.log_date_time_string(), format % args)
        )

    def do_GET(self):
        logger.debug(f"Received GET request: {self.path}")
        try:
            parsed_url = urlparse(self.path)

            # favicon 요청 처리
            if parsed_url.path == "/favicon.ico":
                self.send_response(204)
                self.end_headers()
                return

            query_params = parse_qs(parsed_url.query)

            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()

            if "code" in query_params:
                self.server.oauth_code = query_params["code"][0]
                logger.info("Successfully received OAuth code")
                response = """
                <html>
                <head>
                    <title>인증 완료</title>
                    <style>
                        body { font-family: Arial, sans-serif; text-align: center; padding-top: 50px; }
                        h1 { color: #4285f4; }
                    </style>
                </head>
                <body>
                    <h1>인증이 완료되었습니다</h1>
                    <p>이 창을 닫으셔도 됩니다.</p>
                    <script>window.close();</script>
                </body>
                </html>
                """
                self.wfile.write(response.encode("utf-8"))
            else:
                logger.warning(
                    f"No code in query parameters. Query params: {query_params}"
                )
                self.wfile.write("인증 코드를 받지 못했습니다.".encode("utf-8"))

        except Exception as e:
            logger.error(f"Error handling request: {str(e)}", exc_info=True)
            self.wfile.write("처리 중 오류가 발생했습니다.".encode("utf-8"))


def get_all_sheets_data(creds):
    """
    사용자의 모든 Google Sheets 데이터를 가져옵니다.
    """
    try:
        # Sheets API 서비스 생성
        sheets_service = build("sheets", "v4", credentials=creds)

        # Drive API를 사용하여 사용자의 모든 스프레드시트 파일 목록 가져오기
        drive_service = build("drive", "v3", credentials=creds)
        files = (
            drive_service.files()
            .list(
                q="mimeType='application/vnd.google-apps.spreadsheet'",
                spaces="drive",
                fields="files(id, name)",
                pageSize=1000,
            )
            .execute()
        )

        all_sheets_data = []

        for file in files.get("files", []):
            try:
                # 각 스프레드시트의 모든 시트 정보 가져오기
                spreadsheet = (
                    sheets_service.spreadsheets()
                    .get(spreadsheetId=file["id"])
                    .execute()
                )

                for sheet in spreadsheet.get("sheets", []):
                    sheet_name = sheet["properties"]["title"]
                    range_name = f"{sheet_name}!A1:Z1000"  # 적절한 범위 설정

                    # 시트 데이터 가져오기
                    result = (
                        sheets_service.spreadsheets()
                        .values()
                        .get(spreadsheetId=file["id"], range=range_name)
                        .execute()
                    )

                    values = result.get("values", [])

                    if values:
                        sheet_data = {
                            "spreadsheet_name": file["name"],
                            "spreadsheet_id": file["id"],
                            "sheet_name": sheet_name,
                            "data": values,
                        }
                        all_sheets_data.append(sheet_data)

                        logger.info(
                            f"Retrieved data from sheet '{sheet_name}' in '{file['name']}'"
                        )

            except Exception as sheet_error:
                logger.error(
                    f"Error processing spreadsheet {file['name']}: {sheet_error}"
                )
                continue

        # 시트 데이터 JSON으로 저장
        if all_sheets_data:
            with open("sheets_data.json", "w", encoding="utf-8") as f:
                json.dump(all_sheets_data, f, ensure_ascii=False, indent=4)
                logger.info("Saved sheets data to sheets_data.json")

        return all_sheets_data

    except Exception as e:
        logger.error(f"Error fetching sheets data: {e}")
        return None


def get_credentials():
    creds = None

    if os.path.exists("token.pickle"):
        try:
            with open("token.pickle", "rb") as token:
                creds = pickle.load(token)
            logger.info("Loaded existing credentials from token.pickle")
        except Exception as e:
            logger.error(f"Error loading token file: {str(e)}")
            os.remove("token.pickle")
            logger.info("Removed corrupted token.pickle file")
            return None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                logger.info("Refreshed expired credentials")
            except Exception as e:
                logger.error(f"Error refreshing token: {str(e)}")
                return None
        else:
            try:
                flow = Flow.from_client_secrets_file(
                    "credentials.json",
                    scopes=SCOPES,
                    redirect_uri=f"http://localhost:{PORT}/oauth2callback",
                )

                # OAuth 세션 설정
                flow.oauth2session._client.verify_scopes = False
                flow.oauth2session.scope = sorted(SCOPES)  # 스코프 정렬

                auth_url, _ = flow.authorization_url(
                    access_type="offline",
                    include_granted_scopes="true",
                    prompt="consent",
                )

                server = HTTPServer(("localhost", PORT), OAuthHandler)
                server.oauth_code = None
                logger.info(f"Started OAuth callback server on port {PORT}")

                webbrowser.open(auth_url)

                while server.oauth_code is None:
                    server.handle_request()

                server.server_close()
                logger.info("OAuth callback server stopped")

                try:
                    flow.fetch_token(code=server.oauth_code)
                    creds = flow.credentials
                    logger.info("Successfully obtained new credentials")
                except Warning as w:
                    # 스코프 변경 경고를 무시하고 계속 진행
                    logger.warning(f"Scope warning (ignored): {str(w)}")
                    creds = flow.credentials
                except Exception as e:
                    logger.error(f"Error fetching token: {str(e)}")
                    raise

                with open("token.pickle", "wb") as token:
                    pickle.dump(creds, token)
                logger.info("Saved credentials to token.pickle")

            except Exception as e:
                logger.error(
                    f"Error in authentication process: {str(e)}", exc_info=True
                )
                raise

    return creds


def get_user_info(creds):
    try:
        headers = {"Authorization": f"Bearer {creds.token}"}
        response = requests.get(
            "https://www.googleapis.com/oauth2/v2/userinfo", headers=headers
        )
        response.raise_for_status()

        user_info = response.json()
        logger.debug(f"Received user info: {user_info}")
        return {
            "email": user_info.get("email", ""),
            "name": user_info.get("name", ""),
            "picture": user_info.get("picture", ""),
        }
    except Exception as e:
        logger.error(f"Error getting user info: {str(e)}")
        return None


def get_all_calendar_events(creds):
    try:
        service = build("calendar", "v3", credentials=creds)
        calendar_list = service.calendarList().list().execute()

        all_events = []

        for calendar_list_entry in calendar_list["items"]:
            calendar_id = calendar_list_entry["id"]

            try:
                # 시작 시간 기준으로 최근 1년치와 향후 1년치 이벤트 가져오기
                now = datetime.datetime.utcnow()
                one_year_ago = (now - datetime.timedelta(days=365)).isoformat() + "Z"
                one_year_later = (now + datetime.timedelta(days=365)).isoformat() + "Z"

                events_result = (
                    service.events()
                    .list(
                        calendarId=calendar_id,
                        timeMin=one_year_ago,  # 1년 전부터
                        timeMax=one_year_later,  # 1년 후까지
                        maxResults=1000,  # 캘린더당 최대 1000개 이벤트
                        singleEvents=True,
                        orderBy="startTime",
                    )
                    .execute()
                )

                events = events_result.get("items", [])

                # 각 이벤트에 캘린더 이름 추가
                for event in events:
                    event["calendar_name"] = calendar_list_entry.get(
                        "summary", "Unnamed Calendar"
                    )

                    # 이벤트 시작, 종료 시간 필드 안전하게 처리
                    if "start" in event and "dateTime" not in event["start"]:
                        event["start"]["dateTime"] = (
                            event["start"].get("date") + "T00:00:00"
                        )
                    if "end" in event and "dateTime" not in event["end"]:
                        event["end"]["dateTime"] = (
                            event["end"].get("date") + "T23:59:59"
                        )

                all_events.extend(events)

            except Exception as calendar_error:
                logger.error(
                    f"Error fetching events for calendar {calendar_id}: {calendar_error}"
                )

        # 날짜순으로 정렬
        all_events.sort(
            key=lambda x: x["start"].get(
                "dateTime", x["start"].get("date", "9999-99-99")
            )
        )

        # 캘린더 이벤트 JSON으로 저장
        if all_events:
            with open("calendar_events.json", "w", encoding="utf-8") as f:
                json.dump(all_events, f, ensure_ascii=False, indent=4)
            logger.info(f"총 {len(all_events)}개의 캘린더 이벤트 저장 완료")

        return all_events

    except Exception as e:
        logger.error(f"모든 캘린더 이벤트를 가져오는 중 오류 발생: {e}")
        return None


def get_calendar_events(creds):
    try:
        # Google Calendar API 서비스 생성
        service = build("calendar", "v3", credentials=creds)

        # 현재 날짜부터 1년 후까지의 이벤트 가져오기
        now = datetime.datetime.utcnow().isoformat() + "Z"  # 'Z' indicates UTC time
        one_year_later = (
            datetime.datetime.utcnow() + datetime.timedelta(days=365)
        ).isoformat() + "Z"

        # 캘린더 이벤트 조회
        events_result = (
            service.events()
            .list(
                calendarId="primary",  # 기본 캘린더
                timeMin=now,
                timeMax=one_year_later,
                maxResults=50,  # 최대 50개 이벤트
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )

        events = events_result.get("items", [])

        # 이벤트 상세 정보 출력
        if not events:
            logger.info("No upcoming events found.")
        else:
            logger.info("Upcoming events:")
            for event in events:
                start = event["start"].get("dateTime", event["start"].get("date"))
                summary = event.get("summary", "No title")
                logger.info(f"{summary} - {start}")

        return events

    except Exception as e:
        logger.error(f"An error occurred while fetching calendar events: {e}")
        return None


def get_all_tasks(creds):
    try:
        tasks_service = build("tasks", "v1", credentials=creds)
        task_lists = tasks_service.tasklists().list().execute()

        all_tasks = []

        for task_list in task_lists.get("items", []):
            try:
                # 모든 태스크 가져오기 (날짜 제한 없음)
                tasks = (
                    tasks_service.tasks()
                    .list(
                        tasklist=task_list["id"],
                        maxResults=1000,  # 각 리스트당 최대 1000개 태스크
                        showHidden=True,
                    )
                    .execute()
                )

                for task in tasks.get("items", []):
                    task["task_list_name"] = task_list.get(
                        "title", "이름 없는 태스크 리스트"
                    )

                all_tasks.extend(tasks.get("items", []))

            except Exception as task_list_error:
                logger.error(
                    f"태스크 리스트 {task_list['id']} 가져오기 중 오류: {task_list_error}"
                )

        # 상태와 제목으로 정렬
        all_tasks.sort(key=lambda x: (x.get("status", ""), x.get("title", "")))

        # 태스크 정보 로깅
        if all_tasks:
            logger.info(f"총 태스크 수: {len(all_tasks)}")
            for task in all_tasks:
                logger.info(
                    f"제목: {task.get('title', '제목 없음')}, "
                    f"상태: {task.get('status', '상태 없음')}, "
                    f"리스트: {task.get('task_list_name', '알 수 없는 리스트')}"
                )
        else:
            logger.info("태스크를 찾을 수 없습니다.")

        return all_tasks

    except Exception as e:
        logger.error(f"모든 태스크를 가져오는 중 오류 발생: {e}")
        return None


def get_people_contacts(creds):
    try:
        # Google People API 서비스 생성
        people_service = build("people", "v1", credentials=creds)

        # 연락처 가져오기 (최대 1000개)
        contacts = (
            people_service.people()
            .connections()
            .list(
                resourceName="people/me",
                pageSize=1000,
                personFields="names,emailAddresses,phoneNumbers,organizations,addresses,birthdays",
            )
            .execute()
        )

        all_contacts = contacts.get("connections", [])

        # 연락처 정보 로깅
        if all_contacts:
            logger.info(f"총 연락처 수: {len(all_contacts)}")
            for contact in all_contacts:
                name = contact.get("names", [{}])[0].get("displayName", "이름 없음")
                emails = [
                    email.get("value", "")
                    for email in contact.get("emailAddresses", [])
                ]
                phones = [
                    phone.get("value", "") for phone in contact.get("phoneNumbers", [])
                ]

                logger.info(f"이름: {name}")
                if emails:
                    logger.info(f"이메일: {', '.join(emails)}")
                if phones:
                    logger.info(f"전화번호: {', '.join(phones)}")
                logger.info("---")
        else:
            logger.info("연락처를 찾을 수 없습니다.")

        return all_contacts

    except Exception as e:
        logger.error(f"연락처를 가져오는 중 오류 발생: {e}")
        return None


def main():
    try:
        # 인증 처리
        creds = get_credentials()
        if not creds:
            logger.error("Failed to get credentials")
            return

        # 사용자 정보 가져오기
        user_info = get_user_info(creds)
        if user_info:
            print("\n사용자 정보:")
            print(f"이름: {user_info['name']}")
            print(f"이메일: {user_info['email']}")
            print(f"프로필 사진 URL: {user_info['picture']}")
        else:
            print("사용자 정보를 가져오는데 실패했습니다.")

        # 모든 캘린더 이벤트 가져오기
        all_calendar_events = get_all_calendar_events(creds)
        if all_calendar_events:
            print(f"\n총 {len(all_calendar_events)}개의 일정:")
            for event in all_calendar_events:
                start = event["start"].get("dateTime", event["start"].get("date"))
                calendar_name = event.get("calendar_name", "알 수 없는 캘린더")
                print(f"제목: {event.get('summary', '제목 없음')}")
                print(f"날짜: {start}")
                print(f"캘린더: {calendar_name}")
                print("---")

            # 캘린더 이벤트 JSON으로 저장
            with open("calendar_events.json", "w", encoding="utf-8") as f:
                json.dump(all_calendar_events, f, ensure_ascii=False, indent=4)
                print("캘린더 이벤트를 calendar_events.json에 저장했습니다.")

        # 모든 태스크 가져오기
        all_tasks = get_all_tasks(creds)
        if all_tasks:
            print(f"\n총 {len(all_tasks)}개의 태스크:")
            for task in all_tasks:
                print(f"제목: {task.get('title', '제목 없음')}")
                print(f"상태: {task.get('status', '상태 없음')}")
                print(
                    f"태스크 리스트: {task.get('task_list_name', '알 수 없는 리스트')}"
                )
                print("---")

            # 태스크 JSON으로 저장
            with open("tasks.json", "w", encoding="utf-8") as f:
                json.dump(all_tasks, f, ensure_ascii=False, indent=4)
                print("태스크를 tasks.json에 저장했습니다.")

        # 모든 연락처 가져오기
        all_contacts = get_people_contacts(creds)
        if all_contacts:
            print(f"\n총 {len(all_contacts)}개의 연락처:")
            for contact in all_contacts:
                name = contact.get("names", [{}])[0].get("displayName", "이름 없음")
                emails = [
                    email.get("value", "")
                    for email in contact.get("emailAddresses", [])
                ]
                phones = [
                    phone.get("value", "") for phone in contact.get("phoneNumbers", [])
                ]

                print(f"이름: {name}")
                if emails:
                    print(f"이메일: {', '.join(emails)}")
                if phones:
                    print(f"전화번호: {', '.join(phones)}")
                print("---")

            # 연락처 JSON으로 저장
            with open("contacts.json", "w", encoding="utf-8") as f:
                json.dump(all_contacts, f, ensure_ascii=False, indent=4)
                print("연락처를 contacts.json에 저장했습니다.")

        # 모든 시트 데이터 가져오기
        sheets_data = get_all_sheets_data(creds)
        if sheets_data:
            print(f"\n총 {len(sheets_data)}개의 시트 데이터를 가져왔습니다:")
            for sheet in sheets_data:
                print(f"스프레드시트: {sheet['spreadsheet_name']}")
                print(f"시트: {sheet['sheet_name']}")
                print(f"데이터 행 수: {len(sheet['data'])}")
                print("---")

    except Exception as e:
        logger.error(f"프로그램 오류: {str(e)}", exc_info=True)


# 원래 있던 스크립트 실행 코드 유지
if __name__ == "__main__":
    main()
