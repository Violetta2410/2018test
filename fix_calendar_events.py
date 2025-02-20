import json
from datetime import datetime, timedelta


def fix_calendar_events(input_file, output_file):
    with open(input_file, "r", encoding="utf-8") as f:
        events = json.load(f)

    fixed_events = []
    for event in events:
        # 이벤트 생성 날짜를 기반으로 임시 날짜 설정
        created_date = datetime.fromisoformat(event["created"].replace("Z", "+00:00"))

        # 시작, 종료 시간 복구
        event["start"] = {
            "dateTime": created_date.isoformat(),
            "timeZone": "Asia/Seoul",
        }
        event["end"] = {
            "dateTime": (created_date + timedelta(hours=1)).isoformat(),
            "timeZone": "Asia/Seoul",
        }

        fixed_events.append(event)

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(fixed_events, f, ensure_ascii=False, indent=2)

    print(f"Fixed {len(fixed_events)} events")


# 사용 예시
fix_calendar_events("calendar_events.json", "fixed_calendar_events.json")
