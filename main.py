import os
import datetime
from fastapi import FastAPI, Request
from supabase import create_client, Client
from dotenv import load_dotenv

load_dotenv()
app = FastAPI()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


# 응답 텍스트 포장 도우미
def make_response(text: str):
    return {"version": "2.0", "template": {"outputs": [{"simpleText": {"text": text}}]}}


# 한국 시간(KST) 가져오기 도우미
def get_kst_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=9)


@app.post("/kakao")
def kakao_response(request: Request):
    import asyncio
    body = asyncio.run(request.json())

    user_id = body.get("userRequest", {}).get("user", {}).get("id", "알수없음")
    intent_name = body.get("intent", {}).get("name", "")
    params = body.get("action", {}).get("params", {})

    try:
        # ==========================================
        # 🟢 [기능 1] 회원가입 (기존과 동일)
        # ==========================================
        if intent_name == "회원가입":
            name = params.get("name")
            student_id = params.get("student_id")
            department = params.get("department")
            gender = params.get("gender")

            supabase.table("users").upsert({
                "kakao_id": user_id, "name": name, "student_id": student_id,
                "department": department, "gender": gender, "role": "member"
            }).execute()
            return make_response(f"환영합니다, {name}님! 🎉\n성공적으로 동아리원에 등록되었습니다.")

        # ==========================================
        # 👑 [기능 2] 관리자: 세미나 생성 (자동 시간 계산)
        # ==========================================
        elif intent_name == "세미나생성":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin":
                return make_response("⛔ 관리자 권한이 없습니다.")

            week_name = params.get("week_name", "새로운 주차")

            # 다가오는 금요일 18:00 ~ 일요일 23:59 계산기
            now_kst = get_kst_now()
            days_to_friday = (4 - now_kst.weekday()) % 7

            # 만약 오늘이 금요일인데 이미 오후 6시가 지났다면, 다음 주 금요일로 세팅
            if days_to_friday == 0 and now_kst.hour >= 18:
                days_to_friday = 7

            open_time = (now_kst + datetime.timedelta(days=days_to_friday)).replace(hour=18, minute=0, second=0)
            close_time = (open_time + datetime.timedelta(days=2)).replace(hour=23, minute=59, second=59)

            open_str = open_time.isoformat()
            close_str = close_time.isoformat()

            # 이전 세미나들 비활성화
            supabase.table("seminars").update({"is_active": False}).neq("is_active", None).execute()

            # 월요일/목요일 두 개의 세미나 방 동시 생성 (정원 30명 고정)
            supabase.table("seminars").insert([
                {"week_name": week_name, "day": "월요일", "capacity": 30, "open_time": open_str, "close_time": close_str,
                 "is_active": True},
                {"week_name": week_name, "day": "목요일", "capacity": 30, "open_time": open_str, "close_time": close_str,
                 "is_active": True}
            ]).execute()

            msg = (f"✅ [{week_name}] 예약 세팅 완료!\n\n"
                   f"📌 월/목 세션 동시 오픈\n"
                   f"📌 정원: 각 30명\n"
                   f"⏰ 오픈: {open_time.strftime('%m월 %d일(금) 18:00')}\n"
                   f"⏰ 마감: {close_time.strftime('%m월 %d일(일) 23:59')}")
            return make_response(msg)

        # ==========================================
        # 🔵 [기능 3] 부원: 선착순 요일 투표
        # ==========================================
        elif intent_name == "참석투표":
            day_choice = params.get("day_choice")
            if not day_choice:
                return make_response("요일을 선택해 주세요.")

            # 1. 활성화된 세미나 가져오기
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data:
                return make_response("현재 진행 중인 세미나 투표가 없습니다. 😅")

            # 2. 오픈/마감 시간 체크
            now_str = get_kst_now().isoformat()
            open_str = seminars_res.data[0]["open_time"]
            close_str = seminars_res.data[0]["close_time"]

            if now_str < open_str:
                open_dt = datetime.datetime.fromisoformat(open_str)
                return make_response(f"아직 투표 기간이 아닙니다!\n(오픈: {open_dt.strftime('%m월 %d일 18:00')})")
            if now_str > close_str:
                return make_response("이번 주 세미나 투표가 마감되었습니다. 🥲")

            # 3. 이번 주에 이미 투표했는지 체크 (월/목 중복 방지)
            active_ids = [s["id"] for s in seminars_res.data]
            existing_vote = supabase.table("attendances").select("*").in_("seminar_id", active_ids).eq("kakao_id",
                                                                                                       user_id).execute()

            if existing_vote.data:
                voted_id = existing_vote.data[0]["seminar_id"]
                voted_day = next(s["day"] for s in seminars_res.data if s["id"] == voted_id)
                status_kr = "참석 확정" if existing_vote.data[0]["status"] == "attending" else "대기자"
                return make_response(f"이미 {voted_day} 세미나에 투표하셨습니다!\n📌 현재 상태: {status_kr}")

            # 4. 선택한 요일의 정원 체크 및 등록
            target_seminar = next((s for s in seminars_res.data if s["day"] == day_choice), None)
            if not target_seminar:
                return make_response("월요일 또는 목요일 버튼을 눌러주세요.")

            attendees = supabase.table("attendances").select("id", count="exact").eq("seminar_id",
                                                                                     target_seminar["id"]).eq("status",
                                                                                                              "attending").execute()
            current_count = attendees.count if attendees.count else 0

            if current_count < target_seminar["capacity"]:
                status = "attending"
                msg = f"{day_choice} 참석이 확정되었습니다! 🎉\n(현재: {current_count + 1}/30명)"
            else:
                status = "waitlisted"
                msg = f"{day_choice} 정원이 초과되어 대기자로 등록되었습니다. 🥲\n(취소자 발생 시 자동 전환)"

            supabase.table("attendances").insert({
                "seminar_id": target_seminar["id"], "kakao_id": user_id, "status": status
            }).execute()

            return make_response(f"[{target_seminar['week_name']}] 투표 결과:\n{msg}")

        # 기타 (연결 테스트 등)
        elif intent_name == "연결 테스트":
            return make_response("안녕! 파이썬 서버와 연결되었어! 🚀")
        else:
            return make_response("준비되지 않은 기능입니다.")

    except Exception as e:
        print(f"Error: {e}")
        return make_response("서버 처리 중 오류가 발생했습니다.")