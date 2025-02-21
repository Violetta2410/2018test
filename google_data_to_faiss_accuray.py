import os
from dotenv import load_dotenv
import json
import logging
from typing import List, Dict
from datetime import datetime

from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain.docstore.document import Document

# Load environment variables
load_dotenv()
# 로깅 설정
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


class GoogleDataProcessor:
    def __init__(self, faiss_save_path="./google_data_faiss"):
        # Get API key from environment variable
        self.openai_api_key = os.getenv("OPENAI_API_KEY")
        if not self.openai_api_key:
            raise ValueError("OPENAI_API_KEY not found in environment variables")

        self.embeddings = OpenAIEmbeddings(openai_api_key=self.openai_api_key)
        self.faiss_save_path = faiss_save_path
        self.text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=500, chunk_overlap=50
        )

    def _process_calendar_events(self, events: List[Dict]) -> List[Document]:
        """캘린더 이벤트를 문서로 변환 - 날짜 처리 개선"""
        docs = []
        for event in events:
            try:
                # 시작 시간 처리
                start_key = "dateTime" if "dateTime" in event["start"] else "date"
                end_key = "dateTime" if "dateTime" in event["end"] else "date"

                # 시작/종료 시간 파싱
                if start_key == "dateTime":
                    start_time = datetime.fromisoformat(
                        event["start"][start_key].replace("Z", "").replace("+09:00", "")
                    )
                    end_time = datetime.fromisoformat(
                        event["end"][end_key].replace("Z", "").replace("+09:00", "")
                    )
                else:
                    start_time = datetime.strptime(
                        event["start"][start_key], "%Y-%m-%d"
                    )
                    end_time = datetime.strptime(event["end"][end_key], "%Y-%m-%d")

                # 표준화된 날짜 문자열 생성
                date_str = start_time.strftime("%Y-%m-%d")
                start_time_str = start_time.strftime("%H:%M")
                end_time_str = end_time.strftime("%H:%M")

                content = f"""
               날짜: {date_str}
               시간: {start_time_str} ~ {end_time_str}
               Event_Type: calendar
               Title: {event.get('summary', '제목 없음')}
               Start: {start_time.isoformat()}
               End: {end_time.isoformat()}
               Calendar: {event.get('calendar_name', '알 수 없는 캘린더')}
               Description: {event.get('description', '설명 없음')}
               """

                # 메타데이터에 날짜 정보 추가
                metadata = {
                    "type": "calendar_event",
                    "date": date_str,
                    "start_time": start_time.isoformat(),
                    "end_time": end_time.isoformat(),
                    "title": event.get("summary", "제목 없음"),
                    "calendar": event.get("calendar_name", "알 수 없는 캘린더"),
                }

                docs.append(Document(page_content=content, metadata=metadata))

            except Exception as e:
                logger.error(f"이벤트 처리 중 오류: {e}")
                continue

        return docs

    def _process_tasks(self, tasks: List[Dict]) -> List[Document]:
        """태스크를 문서로 변환"""
        docs = []
        for task in tasks:
            content = f"""
           제목: {task.get('title', '제목 없음')}
           상태: {task.get('status', '상태 없음')}
           태스크 리스트: {task.get('task_list_name', '알 수 없는 리스트')}
           """
            docs.append(
                Document(
                    page_content=content,
                    metadata={
                        "type": "task",
                        "title": task.get("title", "제목 없음"),
                        "list_name": task.get("task_list_name", "알 수 없는 리스트"),
                    },
                )
            )
        return docs

    def _process_contacts(self, contacts: List[Dict]) -> List[Document]:
        """연락처를 문서로 변환"""
        docs = []
        for contact in contacts:
            name = contact.get("names", [{}])[0].get("displayName", "이름 없음")
            emails = [
                email.get("value", "") for email in contact.get("emailAddresses", [])
            ]
            phones = [
                phone.get("value", "") for phone in contact.get("phoneNumbers", [])
            ]

            content = f"""
           이름: {name}
           이메일: {', '.join(emails) if emails else '이메일 없음'}
           전화번호: {', '.join(phones) if phones else '전화번호 없음'}
           """
            docs.append(
                Document(
                    page_content=content, metadata={"type": "contact", "name": name}
                )
            )
        return docs

    def _process_sheets_data(self, sheets_data: List[Dict]) -> List[Document]:
        """Google Sheets 데이터를 문서로 변환"""
        docs = []
        for sheet in sheets_data:
            if len(sheet["data"]) < 2:
                continue

            headers = sheet["data"][0]
            rows = sheet["data"][1:]

            for row_idx, row in enumerate(rows, start=2):
                row_dict = {}
                for i, value in enumerate(row):
                    if i < len(headers):
                        row_dict[headers[i]] = value

                content = f"""
               스프레드시트: {sheet['spreadsheet_name']}
               시트: {sheet['sheet_name']}
               행 번호: {row_idx}
               데이터:
               """
                for header, value in row_dict.items():
                    content += f"{header}: {value}\n"

                docs.append(
                    Document(
                        page_content=content,
                        metadata={
                            "type": "sheets_data",
                            "spreadsheet_name": sheet["spreadsheet_name"],
                            "sheet_name": sheet["sheet_name"],
                            "row_number": row_idx,
                        },
                    )
                )
        return docs

    def create_faiss_index(self, events, tasks, contacts, sheets_data):
        """모든 데이터를 통합하여 FAISS 인덱스 생성"""
        documents = []

        # 캘린더 이벤트 처리
        for event in events:
            try:
                # 시작/종료 시간 처리
                start_key = "dateTime" if "dateTime" in event["start"] else "date"
                end_key = "dateTime" if "dateTime" in event["end"] else "date"

                # 시작/종료 시간 파싱
                if start_key == "dateTime":
                    start_time = datetime.fromisoformat(
                        event["start"][start_key].replace("Z", "").replace("+09:00", "")
                    )
                    end_time = datetime.fromisoformat(
                        event["end"][end_key].replace("Z", "").replace("+09:00", "")
                    )
                else:
                    start_time = datetime.strptime(
                        event["start"][start_key], "%Y-%m-%d"
                    )
                    end_time = datetime.strptime(event["end"][end_key], "%Y-%m-%d")

                # 검색을 위한 표준화된 날짜 형식
                date_str = start_time.strftime("%Y-%m-%d")
                start_time_str = start_time.strftime("%H:%M")
                end_time_str = end_time.strftime("%H:%M")

                content = f"""
               날짜: {date_str}
               시간: {start_time_str} ~ {end_time_str}
               Event_Type: calendar
               Title: {event.get('summary', '제목 없음')}
               Calendar: {event.get('calendar_name', '알 수 없는 캘린더')}
               Description: {event.get('description', '설명 없음')}
               """

                metadata = {
                    "date": date_str,
                    "start_time": start_time.isoformat(),
                    "end_time": end_time.isoformat(),
                    "title": event.get("summary", "제목 없음"),
                    "calendar": event.get("calendar_name"),
                }

                documents.append(Document(page_content=content, metadata=metadata))

            except Exception as e:
                logger.error(f"이벤트 처리 오류: {e}")
                logger.error(f"문제의 이벤트: {event}")
                continue

        # FAISS 인덱스 생성
        db = FAISS.from_documents(documents, self.embeddings)
        db.save_local(self.faiss_save_path)

        logger.info(f"Created FAISS index with {len(documents)} documents")
        return db


def find_events_by_date(date_str):
    """특정 날짜의 이벤트를 찾아 요약하는 함수"""
    try:
        # calendar_events.json 파일 로드
        with open("calendar_events.json", "r", encoding="utf-8") as f:
            events = json.load(f)

        # 입력된 날짜 포맷팅
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date()

        # 해당 날짜의 이벤트 필터링
        matching_events = []
        for event in events:
            # 시간대 정보를 포함한 날짜 처리
            if "dateTime" in event["start"]:
                start_time = datetime.fromisoformat(
                    event["start"]["dateTime"].replace("+09:00", "")
                ).date()
            else:
                start_time = datetime.strptime(
                    event["start"]["date"], "%Y-%m-%d"
                ).date()

            if start_time == target_date:
                matching_events.append(
                    {
                        "summary": event.get("summary", "제목 없음"),
                        "start_time": (
                            datetime.fromisoformat(
                                event["start"]["dateTime"].replace("+09:00", "")
                            ).strftime("%H:%M")
                            if "dateTime" in event["start"]
                            else "00:00"
                        ),
                        "end_time": (
                            datetime.fromisoformat(
                                event["end"]["dateTime"].replace("+09:00", "")
                            ).strftime("%H:%M")
                            if "dateTime" in event["end"]
                            else "23:59"
                        ),
                    }
                )

        # 결과 출력
        if matching_events:
            print(f"{date_str}의 일정:")
            for event in matching_events:
                print(
                    f"- {event['summary']} ({event['start_time']} ~ {event['end_time']})"
                )
        else:
            print(f"{date_str}에는 일정이 없습니다.")

        return matching_events

    except FileNotFoundError:
        print("calendar_events.json 파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")


def main():
    try:
        # calendar_events.json 파일 로드
        with open("calendar_events.json", "r", encoding="utf-8") as f:
            events = json.load(f)
            print(f"Loaded {len(events)} events from calendar_events.json")

        # 프로세서 초기화 및 인덱스 생성
        processor = (
            GoogleDataProcessor()
        )  # API 키는 이제 클래스 내부에서 환경 변수로부터 가져옵니다
        faiss_db = processor.create_faiss_index(
            events=events, tasks=[], contacts=[], sheets_data=[]
        )

        print("FAISS index created successfully!")

    except FileNotFoundError:
        print("calendar_events.json 파일을 찾을 수 없습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")


if __name__ == "__main__":
    main()
