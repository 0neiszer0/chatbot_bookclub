import os
import datetime
from fastapi import FastAPI
from supabase import create_client, Client
from dotenv import load_dotenv

load_dotenv()
app = FastAPI()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


# 텍스트 응답 래퍼
def make_response(text: str):
    return {"version": "2.0", "template": {"outputs": [{"simpleText": {"text": text}}]}}


# KST 시간 가져오기
def get_kst_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=9)


@app.post("/kakao")
def kakao_response(body: dict):
    user_id = body.get("userRequest", {}).get("user", {}).get("id", "알수없음")
    intent_name = body.get("intent", {}).get("name", "")
    params = body.get("action", {}).get("params", {})

    try:
        # ==========================================
        # 🟢 1. 회원가입
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
        # 👑 2. 세미나 생성 (관리자)
        # ==========================================
        elif intent_name == "세미나생성":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin":
                return make_response("⛔ 관리자 권한이 없습니다.")

            week_name = params.get("week_name", "새로운 주차")

            now_kst = get_kst_now()
            days_to_friday = (4 - now_kst.weekday()) % 7
            if days_to_friday == 0 and now_kst.hour >= 18:
                days_to_friday = 7

            open_time = (now_kst + datetime.timedelta(days=days_to_friday)).replace(hour=18, minute=0, second=0)
            close_time = (open_time + datetime.timedelta(days=2)).replace(hour=23, minute=59, second=59)

            supabase.table("seminars").update({"is_active": False}).eq("is_active", True).execute()
            supabase.table("seminars").insert([
                {"week_name": week_name, "day": "월요일", "capacity": 30, "open_time": open_time.isoformat(),
                 "close_time": close_time.isoformat(), "is_active": True},
                {"week_name": week_name, "day": "목요일", "capacity": 30, "open_time": open_time.isoformat(),
                 "close_time": close_time.isoformat(), "is_active": True}
            ]).execute()

            return make_response(
                f"✅ [{week_name}] 예약 세팅 완료!\n⏰ 오픈: {open_time.strftime('%m/%d(금) 18:00')}\n⏰ 마감: {close_time.strftime('%m/%d(일) 23:59')}")

        # ==========================================
        # 🔵 3. 참석 투표
        # ==========================================
        elif intent_name == "참석투표":
            day_choice = params.get("day_choice")
            if not day_choice: return make_response("요일을 선택해 주세요.")

            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("진행 중인 세미나가 없습니다. 😅")

            now_str = get_kst_now().isoformat()
            if now_str < seminars_res.data[0]["open_time"]: return make_response("아직 투표 기간이 아닙니다! (금요일 18시 오픈)")
            if now_str > seminars_res.data[0]["close_time"]: return make_response("투표가 마감되었습니다. 🥲")

            active_ids = [s["id"] for s in seminars_res.data]
            existing_vote = supabase.table("attendances").select("*").in_("seminar_id", active_ids).eq("kakao_id",
                                                                                                       user_id).execute()
            if existing_vote.data: return make_response("이미 투표하셨습니다!\n(상태 확인: '내 상태' 입력)")

            target_seminar = next((s for s in seminars_res.data if s["day"] == day_choice), None)
            if not target_seminar: return make_response("월요일 또는 목요일을 선택해주세요.")

            attendees = supabase.table("attendances").select("id", count="exact").eq("seminar_id",
                                                                                     target_seminar["id"]).eq("status",
                                                                                                              "attending").execute()
            current_count = attendees.count if attendees.count else 0

            status = "attending" if current_count < target_seminar["capacity"] else "waitlisted"
            msg_status = "참석 확정 🎉" if status == "attending" else "대기자 등록 🥲 (취소자 발생 시 자동 승급)"

            supabase.table("attendances").insert(
                {"seminar_id": target_seminar["id"], "kakao_id": user_id, "status": status}).execute()
            return make_response(f"[{target_seminar['week_name']} - {day_choice}]\n{msg_status}")

        # ==========================================
        # 🟡 4. 투표 취소 및 자동 승급
        # ==========================================
        elif intent_name == "투표취소":
            seminars_res = supabase.table("seminars").select("id").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("진행 중인 세미나가 없습니다.")
            active_ids = [s["id"] for s in seminars_res.data]

            my_vote = supabase.table("attendances").select("*").in_("seminar_id", active_ids).eq("kakao_id",
                                                                                                 user_id).execute()
            if not my_vote.data: return make_response("취소할 투표 내역이 없습니다.")

            vote_data = my_vote.data[0]
            supabase.table("attendances").delete().eq("id", vote_data["id"]).execute()

            # 내가 '참석 확정'이었으면 대기자 1번 승급시키기
            if vote_data["status"] == "attending":
                next_in_line = supabase.table("attendances").select("id").eq("seminar_id", vote_data["seminar_id"]).eq(
                    "status", "waitlisted").order("created_at").limit(1).execute()
                if next_in_line.data:
                    supabase.table("attendances").update({"status": "attending"}).eq("id", next_in_line.data[0][
                        "id"]).execute()

            return make_response("✅ 투표가 정상적으로 취소되었습니다.\n(대기자가 있었다면 자동으로 승급 처리되었습니다.)")

        # ==========================================
        # 🟢 5. 내 상태 확인
        # ==========================================
        elif intent_name == "내상태":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("현재 진행 중인 세미나가 없습니다.")

            active_ids = [s["id"] for s in seminars_res.data]
            my_vote = supabase.table("attendances").select("*").in_("seminar_id", active_ids).eq("kakao_id",
                                                                                                 user_id).execute()

            if not my_vote.data: return make_response("이번 주 세미나에 투표하지 않으셨습니다.")

            target_seminar = next(s for s in seminars_res.data if s["id"] == my_vote.data[0]["seminar_id"])
            status_kr = "✅ 참석 확정" if my_vote.data[0]["status"] == "attending" else "⏳ 대기 중"
            return make_response(
                f"[{target_seminar['week_name']}]\n📌 요일: {target_seminar['day']}\n📌 상태: {status_kr}\n\n사정이 생겨 못 오신다면 '투표취소'를 입력해주세요!")

        # ==========================================
        # 👑 6. 명단 확인 (관리자)
        # ==========================================
        elif intent_name == "명단확인":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin":
                return make_response("⛔ 관리자만 열람 가능합니다.")

            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("진행 중인 세미나가 없습니다.")

            attendances = supabase.table("attendances").select("status, seminar_id, users(name)").in_("seminar_id",
                                                                                                      [s["id"] for s in
                                                                                                       seminars_res.data]).eq(
                "status", "attending").execute()

            mon_list, thu_list = [], []
            for a in attendances.data:
                seminar_day = next(s["day"] for s in seminars_res.data if s["id"] == a["seminar_id"])
                # users 테이블 조인 결과에서 이름 추출
                user_name = a.get("users", {}).get("name", "알수없음") if a.get("users") else "알수없음"

                if seminar_day == "월요일":
                    mon_list.append(user_name)
                elif seminar_day == "목요일":
                    thu_list.append(user_name)

            return make_response(
                f"📋 [참석 확정 명단]\n\n[월요일] ({len(mon_list)}/30)\n{', '.join(mon_list) if mon_list else '없음'}\n\n[목요일] ({len(thu_list)}/30)\n{', '.join(thu_list) if thu_list else '없음'}")

        # 기본 응답
        elif intent_name == "연결 테스트":
            return make_response("연결 성공! 🚀")
        else:
            return make_response("준비되지 않은 기능입니다.")

    except Exception as e:
        print(f"Error: {e}")
        return make_response("서버 처리 중 오류가 발생했습니다.")