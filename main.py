import os
import datetime
import random
from collections import defaultdict
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from supabase import create_client, Client
from dotenv import load_dotenv

load_dotenv()
app = FastAPI()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


def make_response(text: str):
    return {"version": "2.0", "template": {"outputs": [{"simpleText": {"text": text}}]}}


def get_kst_now():
    return datetime.datetime.utcnow() + datetime.timedelta(hours=9)


# ==========================================
# 🤖 카카오톡 챗봇 엔드포인트
# ==========================================
@app.post("/kakao")
def kakao_response(body: dict):
    user_id = body.get("userRequest", {}).get("user", {}).get("id", "알수없음")
    intent_name = body.get("intent", {}).get("name", "")
    params = body.get("action", {}).get("params", {})

    try:
        # [1. 회원가입, 2. 세미나생성, 3. 참석투표, 4. 투표취소, 5. 내상태, 6. 명단확인 코드는 이전과 동일하게 유지]
        if intent_name == "회원가입":
            name, student_id, dept, gender = params.get("name"), params.get("student_id"), params.get(
                "department"), params.get("gender")
            supabase.table("users").upsert(
                {"kakao_id": user_id, "name": name, "student_id": student_id, "department": dept, "gender": gender,
                 "role": "member"}).execute()
            return make_response(f"환영합니다, {name}님! 🎉\n성공적으로 동아리원에 등록되었습니다.")

        elif intent_name == "세미나생성":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin": return make_response("⛔ 관리자 권한이 없습니다.")
            week_name = params.get("week_name", "새로운 주차")
            now_kst = get_kst_now()
            days_to_friday = (4 - now_kst.weekday()) % 7
            if days_to_friday == 0 and now_kst.hour >= 18: days_to_friday = 7
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
            if existing_vote.data: return make_response("이미 투표하셨습니다! (상태 확인: '내 상태')")
            target_seminar = next((s for s in seminars_res.data if s["day"] == day_choice), None)
            attendees = supabase.table("attendances").select("id", count="exact").eq("seminar_id",
                                                                                     target_seminar["id"]).eq("status",
                                                                                                              "attending").execute()
            current_count = attendees.count if attendees.count else 0
            status = "attending" if current_count < target_seminar["capacity"] else "waitlisted"
            msg_status = "참석 확정 🎉" if status == "attending" else "대기자 등록 🥲 (취소자 발생 시 자동 승급)"
            supabase.table("attendances").insert(
                {"seminar_id": target_seminar["id"], "kakao_id": user_id, "status": status}).execute()
            return make_response(f"[{target_seminar['week_name']} - {day_choice}]\n{msg_status}")

        elif intent_name == "투표취소":
            seminars_res = supabase.table("seminars").select("id").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("진행 중인 세미나가 없습니다.")
            my_vote = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("kakao_id",
                                                                                                             user_id).execute()
            if not my_vote.data: return make_response("취소할 투표 내역이 없습니다.")
            vote_data = my_vote.data[0]
            supabase.table("attendances").delete().eq("id", vote_data["id"]).execute()
            if vote_data["status"] == "attending":
                next_in_line = supabase.table("attendances").select("id").eq("seminar_id", vote_data["seminar_id"]).eq(
                    "status", "waitlisted").order("created_at").limit(1).execute()
                if next_in_line.data: supabase.table("attendances").update({"status": "attending"}).eq("id",
                                                                                                       next_in_line.data[
                                                                                                           0][
                                                                                                           "id"]).execute()
            return make_response("✅ 취소 완료. (대기자가 있었다면 자동 승급되었습니다.)")

        elif intent_name == "내상태":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("진행 중인 세미나가 없습니다.")
            my_vote = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("kakao_id",
                                                                                                             user_id).execute()
            if not my_vote.data: return make_response("이번 주 세미나에 투표하지 않으셨습니다.")
            target_seminar = next(s for s in seminars_res.data if s["id"] == my_vote.data[0]["seminar_id"])
            status_kr = "✅ 참석 확정" if my_vote.data[0]["status"] == "attending" else "⏳ 대기 중"
            team_info = f"\n📌 소속 조: {my_vote.data[0].get('team_name', '아직 편성 안 됨')}" if my_vote.data[0][
                                                                                            "status"] == "attending" else ""
            return make_response(
                f"[{target_seminar['week_name']}]\n📌 요일: {target_seminar['day']}\n📌 상태: {status_kr}{team_info}\n\n사정이 생기면 '투표취소'를 입력해주세요.")

        elif intent_name == "명단확인":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin": return make_response("⛔ 관리자 전용입니다.")
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("진행 중인 세미나가 없습니다.")
            att_res = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("status",
                                                                                                             "attending").execute()
            users_res = supabase.table("users").select("kakao_id, name").in_("kakao_id", [a["kakao_id"] for a in
                                                                                          att_res.data]).execute()
            user_dict = {u["kakao_id"]: u["name"] for u in users_res.data}
            mon_list, thu_list = [], []
            for a in att_res.data:
                day = next(s["day"] for s in seminars_res.data if s["id"] == a["seminar_id"])
                name = user_dict.get(a["kakao_id"], "알수없음")
                team = f"({a.get('team_name', '-')})" if a.get('team_name') else ""
                if day == "월요일":
                    mon_list.append(f"{name}{team}")
                else:
                    thu_list.append(f"{name}{team}")
            return make_response(
                f"📋 [참석 확정 명단]\n\n[월요일]\n{', '.join(mon_list) if mon_list else '없음'}\n\n[목요일]\n{', '.join(thu_list) if thu_list else '없음'}")

        # ==========================================
        # 📝 [신규] 발제문 제출
        # ==========================================
        elif intent_name == "발제문제출":
            topic_content = params.get("topic_content")
            if not topic_content: return make_response("제출할 내용을 적어주세요.")

            seminars_res = supabase.table("seminars").select("id").eq("is_active", True).execute()
            if not seminars_res.data: return make_response("현재 활성화된 세미나가 없습니다.")

            # 내 투표 내역 찾기
            my_vote = supabase.table("attendances").select("id").in_("seminar_id",
                                                                     [s["id"] for s in seminars_res.data]).eq(
                "kakao_id", user_id).eq("status", "attending").execute()
            if not my_vote.data: return make_response("이번 주 '참석 확정' 상태인 부원만 발제문을 제출할 수 있습니다.")

            # 발제자로 마킹 및 내용 저장
            supabase.table("attendances").update({"is_facilitator": True, "topic_content": topic_content}).eq("id",
                                                                                                              my_vote.data[
                                                                                                                  0][
                                                                                                                  "id"]).execute()
            return make_response("✅ 발제문이 성공적으로 제출되었습니다!\n(자동으로 이번 주 조장으로 우선 배정됩니다.)")

        # ==========================================
        # 🧠 [신규] 랜덤 시뮬레이션 조 편성 + 리포트
        # ==========================================
        elif intent_name == "조편성":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin": return make_response("⛔ 관리자 전용입니다.")

            day_choice = params.get("day_choice")
            if not day_choice: return make_response("어느 요일 조를 편성할까요? (월요일/목요일)")

            seminar_res = supabase.table("seminars").select("*").eq("is_active", True).eq("day", day_choice).execute()
            if not seminar_res.data: return make_response(f"활성화된 {day_choice} 세미나가 없습니다.")
            target_seminar = seminar_res.data[0]

            att_res = supabase.table("attendances").select("*").eq("seminar_id", target_seminar["id"]).eq("status",
                                                                                                          "attending").execute()
            if not att_res.data: return make_response("참석 확정자가 없습니다.")

            users_res = supabase.table("users").select("kakao_id, name, gender").in_("kakao_id", [a["kakao_id"] for a in
                                                                                                  att_res.data]).execute()
            user_dict = {u["kakao_id"]: u for u in users_res.data}

            attendees = [{"att_id": a["id"], "kakao_id": a["kakao_id"], "name": user_dict[a["kakao_id"]]["name"],
                          "gender": user_dict[a["kakao_id"]]["gender"], "is_fac": a.get("is_facilitator", False)} for a
                         in att_res.data if a["kakao_id"] in user_dict]

            # 과거 만남 기록 가져오기
            past_res = supabase.table("attendances").select("seminar_id, team_name, kakao_id").neq("team_name",
                                                                                                   "None").execute()
            past_groups = defaultdict(list)
            for p in past_res.data:
                if p["team_name"]: past_groups[f"{p['seminar_id']}_{p['team_name']}"].append(p["kakao_id"])

            past_encounters = defaultdict(int)
            for members in past_groups.values():
                for i in range(len(members)):
                    for j in range(i + 1, len(members)):
                        pair = tuple(sorted([members[i], members[j]]))
                        past_encounters[pair] += 1

            # 1만 회 랜덤 시뮬레이션
            num_people = len(attendees)
            num_teams = max(1, num_people // 4)
            team_sizes = [4] * num_teams
            for i in range(num_people % 4): team_sizes[i] += 1

            best_score = float('inf')
            best_teams = []

            facs = [a for a in attendees if a["is_fac"]]
            norms = [a for a in attendees if not a["is_fac"]]

            for _ in range(10000):
                random.shuffle(facs)
                random.shuffle(norms)

                teams = [[] for _ in range(num_teams)]
                f_idx, n_idx = 0, 0

                for f in facs:
                    teams[f_idx % num_teams].append(f)
                    f_idx += 1

                for i, size in enumerate(team_sizes):
                    while len(teams[i]) < size and n_idx < len(norms):
                        teams[i].append(norms[n_idx])
                        n_idx += 1

                score = 0
                for t in teams:
                    males = sum(1 for x in t if x['gender'] == '남')
                    females = sum(1 for x in t if x['gender'] == '여')
                    score += abs(males - females) * 100  # 성비 불균형 페널티 (최우선)

                    for i in range(len(t)):
                        for j in range(i + 1, len(t)):
                            pair = tuple(sorted([t[i]['kakao_id'], t[j]['kakao_id']]))
                            if pair in past_encounters:
                                score += past_encounters[pair] * 10  # 과거 만남 페널티

                if score < best_score:
                    best_score, best_teams = score, teams
                    if score == 0: break

            # DB에 최적 결과 저장 & 리포트 작성
            report = f"✨ [{target_seminar['week_name']} - {day_choice}] 조 편성 초안\n\n"
            for i, team in enumerate(best_teams):
                team_name = f"{i + 1}조"
                fac_names = [m['name'] for m in team if m['is_fac']]
                norm_names = [m['name'] for m in team if not m['is_fac']]
                m_cnt = sum(1 for m in team if m['gender'] == '남')
                f_cnt = len(team) - m_cnt

                report += f"📍 {team_name} (여{f_cnt}, 남{m_cnt})\n"
                report += f"👑발제자: {', '.join(fac_names) if fac_names else '없음'}\n"
                report += f"부원: {', '.join(norm_names)}\n"

                # 특이사항 (과거 만남) 체크
                warnings = []
                for idx1 in range(len(team)):
                    for idx2 in range(idx1 + 1, len(team)):
                        pair = tuple(sorted([team[idx1]['kakao_id'], team[idx2]['kakao_id']]))
                        if pair in past_encounters:
                            warnings.append(f"{team[idx1]['name']}&{team[idx2]['name']}({past_encounters[pair]}회)")

                if warnings:
                    report += f"⚠️ 특이사항: {', '.join(warnings)}\n\n"
                else:
                    report += f"⚠️ 특이사항: 중복 인원 없음!\n\n"

                for m in team:
                    supabase.table("attendances").update({"team_name": team_name}).eq("id", m["att_id"]).execute()

            # 관리자용 웹 링크 (Render 서버 주소로 자동 연결)
            server_host = os.environ.get("RENDER_EXTERNAL_URL", "http://localhost:8000")
            edit_url = f"{server_host}/admin/edit_teams?seminar_id={target_seminar['id']}&day={day_choice}"

            response = {
                "version": "2.0",
                "template": {
                    "outputs": [{"simpleText": {"text": report}}],
                    "quickReplies": [
                        {"label": "🛠️ 조원 수정하기", "action": "webLink", "webLinkUrl": edit_url}
                    ]
                }
            }
            return response

        elif intent_name == "연결 테스트":
            return make_response("연결 성공! 🚀")
        else:
            return make_response("준비되지 않은 기능입니다.")

    except Exception as e:
        print(f"Error: {e}")
        return make_response("서버 처리 중 오류가 발생했습니다.")


# ==========================================
# 🌐 [신규] 관리자 조원 교체용 비밀 웹페이지
# ==========================================
@app.get("/admin/edit_teams", response_class=HTMLResponse)
def edit_teams_page(seminar_id: int, day: str):
    # 1. 참석자 명단 가져오기
    att_res = supabase.table("attendances").select("*").eq("seminar_id", seminar_id).eq("status", "attending").execute()
    users_res = supabase.table("users").select("kakao_id, name").in_("kakao_id",
                                                                     [a["kakao_id"] for a in att_res.data]).execute()
    user_dict = {u["kakao_id"]: u["name"] for u in users_res.data}

    attendees = [
        {"att_id": a["id"], "name": user_dict.get(a["kakao_id"], "알수없음"), "team_name": a.get("team_name", "1조")} for a
        in att_res.data]
    attendees.sort(key=lambda x: x["team_name"])

    # 2. HTML 화면 생성 (깔끔한 모바일 UI)
    html_content = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{day} 조원 수정</title>
        <style>
            body {{ font-family: 'Pretendard', sans-serif; padding: 20px; background-color: #f4f5f7; }}
            .container {{ background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
            h2 {{ text-align: center; color: #333; }}
            .member-row {{ display: flex; justify-content: space-between; align-items: center; padding: 10px 0; border-bottom: 1px solid #eee; }}
            select {{ padding: 8px; border-radius: 6px; border: 1px solid #ccc; font-size: 16px; }}
            .save-btn {{ display: block; width: 100%; padding: 15px; margin-top: 20px; background: #fee500; border: none; border-radius: 8px; font-size: 18px; font-weight: bold; cursor: pointer; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>🛠️ {day} 조원 핀셋 수정</h2>
            <p style="text-align: center; color: #666; font-size: 14px;">사람 이름 옆의 조를 누르고 <strong>[저장]</strong>을 누르세요.</p>
            <div id="members-list">
    """

    team_options = "".join([f"<option value='{i}조'>{i}조</option>" for i in range(1, 11)])

    for att in attendees:
        html_content += f"""
        <div class="member-row">
            <span>👤 {att['name']}</span>
            <select id="user-{att['att_id']}" data-id="{att['att_id']}">
                <option value="{att['team_name']}" selected hidden>{att['team_name']}</option>
                {team_options}
            </select>
        </div>
        """

    html_content += """
            </div>
            <button class="save-btn" onclick="saveTeams()">💾 변경사항 저장</button>
            <p id="result-msg" style="text-align: center; margin-top: 15px; color: green; font-weight: bold;"></p>
        </div>
        <script>
            async function saveTeams() {
                const selects = document.querySelectorAll("select");
                const payload = Array.from(selects).map(s => ({ "att_id": s.getAttribute("data-id"), "team_name": s.value }));

                const response = await fetch("/api/admin/save_teams", {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ "updates": payload })
                });

                if (response.ok) {
                    document.getElementById("result-msg").innerText = "저장 성공! 카톡 창으로 돌아가셔도 됩니다.";
                } else {
                    document.getElementById("result-msg").innerText = "저장에 실패했습니다.";
                    document.getElementById("result-msg").style.color = "red";
                }
            }
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content, status_code=200)


# ==========================================
# 🌐 [신규] 관리자 조원 교체 저장 API
# ==========================================
@app.post("/api/admin/save_teams")
def save_teams_api(body: dict):
    updates = body.get("updates", [])
    try:
        # DB에 바뀐 팀 이름 저장
        for u in updates: supabase.table("attendances").update({"team_name": u["team_name"]}).eq("id",
                                                                                                 u["att_id"]).execute()
        return {"status": "success"}
    except Exception as e:
        return {"status": "error", "message": str(e)}