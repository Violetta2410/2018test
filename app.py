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
    "https://www.googleapis.com/auth/calendar.readonly",
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/calendar.events.readonly",
    "https://www.googleapis.com/auth/contacts.readonly",
    "https://www.googleapis.com/auth/tasks.readonly",
    # 새로 추가된 API 범위들
    "https://www.googleapis.com/auth/contacts.readonly",
    "https://www.googleapis.com/auth/tasks.readonly",
    "https://www.googleapis.com/auth/contacts",
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

# 안전하지 않은 로컬 연결 허용
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"


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
                # 리디렉션 URI를 정확히 지정
                flow = Flow.from_client_secrets_file(
                    "credentials.json",
                    scopes=SCOPES,
                    redirect_uri=f"http://localhost:{PORT}/oauth2callback",
                )

                auth_url, _ = flow.authorization_url(
                    access_type="offline",
                    include_granted_scopes="true",
                    prompt="consent",  # 새로운 토큰 요청
                )

                # 서버 시작
                server = HTTPServer(("localhost", PORT), OAuthHandler)
                server.oauth_code = None
                logger.info(f"Started OAuth callback server on port {PORT}")

                # 브라우저에서 인증 URL 열기
                logger.info(f"Opening auth URL: {auth_url}")
                webbrowser.open(auth_url)

                # 코드를 받을 때까지 서버 실행
                while server.oauth_code is None:
                    server.handle_request()

                # 서버 종료
                server.server_close()
                logger.info("OAuth callback server stopped")

                # 토큰 얻기
                flow.fetch_token(code=server.oauth_code)
                creds = flow.credentials
                logger.info("Successfully obtained new credentials")

                # 토큰 저장
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
                # 날짜 제한 제거
                events_result = (
                    service.events()
                    .list(
                        calendarId=calendar_id,
                        maxResults=1000,  # 캘린더당 최대 1000개 이벤트
                        singleEvents=True,
                        orderBy="startTime",
                    )
                    .execute()
                )

                events = events_result.get("items", [])

                for event in events:
                    event["calendar_name"] = calendar_list_entry.get(
                        "summary", "unnamed calendar"
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

        # 이벤트 정보 로깅
        if all_events:
            logger.info(f"Total events found: {len(all_events)}")
            for event in all_events:
                start = event["start"].get(
                    "dateTime", event["start"].get("date", "날짜 없음")
                )
                summary = event.get("summary", "No title")
                calendar_name = event.get("calendar_name", "Unknown calendar")
                logger.info(f"{summary} - {start} (from {calendar_name})")
        else:
            logger.info("No events found across all calendars.")

        return all_events

    except Exception as e:
        logger.error(f"An error occurred while fetching all calendar events: {e}")
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

    except Exception as e:
        logger.error(f"프로그램 오류: {str(e)}", exc_info=True)


# 원래 있던 스크립트 실행 코드 유지
if __name__ == "__main__":
    main()
