import os
import datetime
import random
import json
from collections import defaultdict
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from supabase import create_client, Client
from dotenv import load_dotenv

load_dotenv()
app = FastAPI()

# 템플릿 엔진 설정 (templates 폴더 연결)
templates = Jinja2Templates(directory="templates")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


def make_kakao_response(text: str, quick_replies: list = None):
    response = {"version": "2.0", "template": {"outputs": [{"simpleText": {"text": text}}]}}
    if quick_replies: response["template"]["quickReplies"] = quick_replies
    return response


def get_kst_now(): return datetime.datetime.utcnow() + datetime.timedelta(hours=9)


def get_server_host(): return os.environ.get("RENDER_EXTERNAL_URL", "http://localhost:8000")


@app.post("/kakao")
def kakao_bot_main(body: dict):
    user_id = body.get("userRequest", {}).get("user", {}).get("id", "알수없음")
    intent_name = body.get("intent", {}).get("name", "")
    params = body.get("action", {}).get("params", {})

    try:
        if intent_name == "연결 테스트":
            return make_kakao_response("연결 성공! 🚀")

        elif intent_name == "회원가입":
            name, student_id, dept, gender = params.get("name"), params.get("student_id"), params.get(
                "department"), params.get("gender")
            supabase.table("users").upsert(
                {"kakao_id": user_id, "name": name, "student_id": student_id, "department": dept, "gender": gender,
                 "role": "member"}).execute()
            return make_kakao_response(f"환영합니다, {name}님! 🎉\n회원 등록 완료.")

        elif intent_name == "세미나생성":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin": return make_kakao_response("⛔ 관리자 권한이 없습니다.")
            week_name = params.get("week_name", "새로운 주차")
            now_kst = get_kst_now();
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
            return make_kakao_response(
                f"✅ [{week_name}] 예약 세팅 완료!\n⏰ 오픈: {open_time.strftime('%m/%d(금) 18:00')}\n⏰ 마감: {close_time.strftime('%m/%d(일) 23:59')}")

        elif intent_name == "참석투표":
            day_choice = params.get("day_choice")
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
            now_str = get_kst_now().isoformat()
            if now_str < seminars_res.data[0]["open_time"]: return make_kakao_response("아직 투표 기간이 아닙니다!")
            if now_str > seminars_res.data[0]["close_time"]: return make_kakao_response("투표가 마감되었습니다.")
            active_ids = [s["id"] for s in seminars_res.data]
            existing_vote = supabase.table("attendances").select("*").in_("seminar_id", active_ids).eq("kakao_id",
                                                                                                       user_id).execute()
            if existing_vote.data: return make_kakao_response("이미 투표하셨습니다!\n('내 상태' 입력)")
            target_seminar = next((s for s in seminars_res.data if s["day"] == day_choice), None)
            attendees = supabase.table("attendances").select("id", count="exact").eq("seminar_id",
                                                                                     target_seminar["id"]).eq("status",
                                                                                                              "attending").execute()
            current_count = attendees.count if attendees.count else 0
            status = "attending" if current_count < target_seminar["capacity"] else "waitlisted"
            supabase.table("attendances").insert(
                {"seminar_id": target_seminar["id"], "kakao_id": user_id, "status": status}).execute()
            msg = "참석 확정 🎉" if status == "attending" else "대기자 등록 🥲"
            return make_kakao_response(f"[{day_choice}]\n{msg}")

        elif intent_name == "투표취소":
            seminars_res = supabase.table("seminars").select("id").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
            my_vote = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("kakao_id",
                                                                                                             user_id).execute()
            if not my_vote.data: return make_kakao_response("취소할 투표 내역이 없습니다.")
            vote_data = my_vote.data[0]
            supabase.table("attendances").delete().eq("id", vote_data["id"]).execute()
            if vote_data["status"] == "attending":
                next_in_line = supabase.table("attendances").select("id").eq("seminar_id", vote_data["seminar_id"]).eq(
                    "status", "waitlisted").order("created_at").limit(1).execute()
                if next_in_line.data: supabase.table("attendances").update({"status": "attending"}).eq("id",
                                                                                                       next_in_line.data[
                                                                                                           0][
                                                                                                           "id"]).execute()
            return make_kakao_response("✅ 취소 완료. (대기자가 있었다면 자동 승급되었습니다.)")

        elif intent_name == "내상태":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
            my_vote = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("kakao_id",
                                                                                                             user_id).execute()
            if not my_vote.data: return make_kakao_response("이번 주 세미나에 투표하지 않으셨습니다.")
            target_seminar = next(s for s in seminars_res.data if s["id"] == my_vote.data[0]["seminar_id"])
            status_kr = "✅ 참석 확정" if my_vote.data[0]["status"] == "attending" else "⏳ 대기 중"
            team_info = f"\n📌 소속 조: {my_vote.data[0].get('team_name', '아직 편성 안 됨')}" if my_vote.data[0][
                                                                                            "status"] == "attending" else ""
            return make_kakao_response(
                f"[{target_seminar['week_name']}]\n📌 요일: {target_seminar['day']}\n📌 상태: {status_kr}{team_info}\n\n사정이 생기면 '투표취소'를 입력해주세요.")

        elif intent_name == "명단확인":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin": return make_kakao_response("⛔ 관리자 전용입니다.")
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
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
            return make_kakao_response(
                f"📋 [참석 확정 명단]\n\n[월요일]\n{', '.join(mon_list) if mon_list else '없음'}\n\n[목요일]\n{', '.join(thu_list) if thu_list else '없음'}")

        # ----------------------------------------------------
        # 웹페이지 연결 인텐트들
        # ----------------------------------------------------
        elif intent_name == "기록열람":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin": return make_kakao_response("⛔ 관리자 전용입니다.")
            history_url = f"{get_server_host()}/admin/history?admin_key={SUPABASE_KEY[:10]}"
            replies = [{"label": "📋 기록 열람 페이지 오픈", "action": "webLink", "webLinkUrl": history_url}]
            return make_kakao_response("아래 버튼을 눌러 과거 기록을 확인하세요!", replies)

        elif intent_name == "발제문제출":
            seminars_res = supabase.table("seminars").select("id").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("활성화된 세미나가 없습니다.")
            active_ids = [s["id"] for s in seminars_res.data]
            my_vote = supabase.table("attendances").select("id, seminar_id").in_("seminar_id", active_ids).eq(
                "kakao_id", user_id).eq("status", "attending").execute()
            if not my_vote.data: return make_kakao_response("이번 주 '참석 확정' 부원만 발제 제출이 가능합니다.")

            submit_url = f"{get_server_host()}/submit_topic?att_id={my_vote.data[0]['id']}"
            replies = [{"label": "📝 발제문 작성하기", "action": "webLink", "webLinkUrl": submit_url}]
            return make_kakao_response("아래 버튼을 눌러 웹페이지에서 책 정보와 발제 내용을 작성해 주세요!", replies)

        elif intent_name == "조편성":
            user_db = supabase.table("users").select("role").eq("kakao_id", user_id).execute()
            if not user_db.data or user_db.data[0].get("role") != "admin": return make_kakao_response("⛔ 관리자 전용입니다.")
            day_choice = params.get("day_choice")
            seminar_res = supabase.table("seminars").select("*").eq("is_active", True).eq("day", day_choice).execute()
            if not seminar_res.data: return make_kakao_response(f"활성화된 {day_choice} 세미나가 없습니다.")
            target_id = seminar_res.data[0]["id"]

            att_res = supabase.table("attendances").select("*").eq("seminar_id", target_id).eq("status",
                                                                                               "attending").execute()
            if len(att_res.data) < 2: return make_kakao_response("참석자가 너무 적어 조 편성이 불가능합니다.")

            users_res = supabase.table("users").select("kakao_id, name, gender").in_("kakao_id", [a["kakao_id"] for a in
                                                                                                  att_res.data]).execute()
            user_dict = {u["kakao_id"]: u for u in users_res.data}
            attendees = [{"att_id": a["id"], "kakao_id": a["kakao_id"], "name": user_dict[a["kakao_id"]]["name"],
                          "gender": user_dict[a["kakao_id"]]["gender"], "is_fac": a.get("is_facilitator", False)} for a
                         in att_res.data if a["kakao_id"] in user_dict]

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

            best_teams, best_score = [], float('inf')
            facs = [a for a in attendees if a["is_fac"]];
            norms = [a for a in attendees if not a["is_fac"]]
            num_teams = max(1, len(attendees) // 4);
            team_sizes = [4] * num_teams
            for i in range(len(attendees) % 4): team_sizes[i] += 1

            for _ in range(10000):
                random.shuffle(facs);
                random.shuffle(norms);
                teams = [[] for _ in range(num_teams)]
                f_idx = 0
                for f in facs: teams[f_idx % num_teams].append(f); f_idx += 1
                n_idx = 0
                for i, size in enumerate(team_sizes):
                    while len(teams[i]) < size and n_idx < len(norms): teams[i].append(norms[n_idx]); n_idx += 1

                score = 0
                for t in teams:
                    m = sum(1 for x in t if x['gender'] == '남');
                    f = len(t) - m
                    score += abs(m - f) * 100
                    for i in range(len(t)):
                        for j in range(i + 1, len(t)):
                            pair = tuple(sorted([t[i]['kakao_id'], t[j]['kakao_id']]))
                            if pair in past_encounters: score += past_encounters[pair] * 10
                if score < best_score: best_score, best_teams = score, teams
                if score == 0: break

            report = f"✨ [{day_choice}] 조 편성 초안\n\n"
            for i, team in enumerate(best_teams):
                t_name = f"{i + 1}조";
                m = sum(1 for x in t if x['gender'] == '남');
                f = len(t) - m
                facs_str = ", ".join([x['name'] for x in team if x['is_fac']])
                norms_str = ", ".join([x['name'] for x in team if not x['is_fac']])
                report += f"📍 {t_name} (여{f}, 남{m})\n👑발제: {facs_str if facs_str else '없음'}\n부원: {norms_str}\n\n"

                warnings = []
                for idx1 in range(len(team)):
                    for idx2 in range(idx1 + 1, len(team)):
                        pair = tuple(sorted([team[idx1]['kakao_id'], team[idx2]['kakao_id']]))
                        if pair in past_encounters: warnings.append(
                            f"{team[idx1]['name']}&{team[idx2]['name']}({past_encounters[pair]}회)")
                if warnings:
                    report += f"⚠️ 특이사항: {', '.join(warnings)}\n\n"
                else:
                    report += f"⚠️ 특이사항: 깨끗함!\n\n"

                for m_att in team: supabase.table("attendances").update({"team_name": t_name}).eq("id", m_att[
                    "att_id"]).execute()

            edit_url = f"{get_server_host()}/admin/edit_teams?seminar_id={target_id}&day={day_choice}&admin_key={SUPABASE_KEY[:10]}"
            replies = [{"label": "🛠️ 조원 핀셋 수정하기", "action": "webLink", "webLinkUrl": edit_url}]
            return make_kakao_response(report, replies)

    except Exception as e:
        print(f"Error: {e}")
        return make_kakao_response("서버 오류 발생.")


# ==========================================
# 🌐 웹페이지 라우팅 (Jinja2 Templates 연동)
# ==========================================

@app.get("/submit_topic", response_class=HTMLResponse)
def submit_topic_page(request: Request, att_id: int):
    # HTML 파일에 변수(att_id)를 넘겨서 렌더링
    return templates.TemplateResponse("topic_submit.html", {"request": request, "att_id": att_id})


@app.post("/api/submit_topic")
def api_submit_topic(data: dict):
    try:
        att_id = data.pop('att_id')
        topic_json = {
            "book_title": data.get('book_title'), "author": data.get('author'),
            "range": data.get('range'), "questions": [data.get('q1'), data.get('q2'), data.get('q3')]
        }
        supabase.table("attendances").update({"is_facilitator": True, "topic_content": topic_json}).eq("id",
                                                                                                       att_id).execute()
        return {"status": "success"}
    except Exception as e:
        return {"status": "error", "message": str(e)}


@app.get("/admin/edit_teams", response_class=HTMLResponse)
def edit_teams_page(request: Request, seminar_id: int, day: str, admin_key: str):
    if admin_key != SUPABASE_KEY[:10]: return HTMLResponse("⛔ 권한 없음")

    att_res = supabase.table("attendances").select("*").eq("seminar_id", seminar_id).eq("status", "attending").execute()
    users_res = supabase.table("users").select("kakao_id, name").in_("kakao_id",
                                                                     [a["kakao_id"] for a in att_res.data]).execute()
    user_dict = {u["kakao_id"]: u["name"] for u in users_res.data}
    attendees = [
        {"att_id": a["id"], "name": user_dict.get(a["kakao_id"], "알수없음"), "team_name": a.get("team_name", "1조")} for a
        in att_res.data]
    attendees.sort(key=lambda x: x["team_name"])

    return templates.TemplateResponse("admin_edit_teams.html", {
        "request": request, "day": day, "attendees": attendees
    })


@app.post("/api/admin/save_teams")
def api_save_teams(body: dict):
    try:
        for u in body.get("updates", []):
            supabase.table("attendances").update({"team_name": u["team_name"]}).eq("id", u["att_id"]).execute()
        return {"status": "success"}
    except Exception as e:
        return {"status": "error", "message": str(e)}


@app.get("/admin/history", response_class=HTMLResponse)
def admin_history_page(request: Request, admin_key: str):
    if admin_key != SUPABASE_KEY[:10]: return HTMLResponse("⛔ 권한 없음")

    seminars_res = supabase.table("seminars").select("*").eq("is_active", False).order("created_at", desc=True).limit(
        10).execute()
    seminar_ids = [s["id"] for s in seminars_res.data]

    if not seminar_ids:
        return templates.TemplateResponse("admin_history.html", {"request": request, "history_data": []})

    att_res = supabase.table("attendances").select("*").in_("seminar_id", seminar_ids).eq("status", "attending").neq(
        "team_name", "None").execute()
    kakao_ids = list(set([a["kakao_id"] for a in att_res.data]))
    users_res = supabase.table("users").select("kakao_id, name").in_("kakao_id", kakao_ids).execute()
    user_dict = {u["kakao_id"]: u["name"] for u in users_res.data}

    history_data = []
    for sem in seminars_res.data:
        sem_id = sem["id"]
        sem_att = [a for a in att_res.data if a["seminar_id"] == sem_id]
        teams = defaultdict(list)
        for att in sem_att:
            name = user_dict.get(att["kakao_id"], "알수없음")
            teams[att["team_name"]].append({
                "name": name, "is_fac": att["is_facilitator"], "topic": att["topic_content"]
            })

        history_data.append({
            "week": sem["week_name"], "day": sem["day"],
            "date": datetime.datetime.fromisoformat(sem["created_at"]).strftime('%Y-%m-%d'),
            "teams": dict(sorted(teams.items()))
        })

    return templates.TemplateResponse("admin_history.html", {"request": request, "history_data": history_data})