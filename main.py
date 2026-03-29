import os
import datetime
import random
import json
from io import BytesIO
from collections import defaultdict
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from docxtpl import DocxTemplate
from supabase import create_client, Client
from dotenv import load_dotenv

load_dotenv()
app = FastAPI()
templates = Jinja2Templates(directory="templates")

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


# 💡 일반 텍스트 + 메시지 버튼 응답용
def make_kakao_response(text: str, quick_replies: list = None):
    response = {"version": "2.0", "template": {"outputs": [{"simpleText": {"text": text}}]}}
    if quick_replies: response["template"]["quickReplies"] = quick_replies
    return response


# 💡 [핵심] 텍스트 + 예쁜 웹 링크 버튼 카드 응답용! (이게 있어야 링크가 보입니다)
def make_kakao_link_response(text: str, button_label: str, url: str):
    return {
        "version": "2.0",
        "template": {
            "outputs": [
                {"simpleText": {"text": text}},
                {
                    "textCard": {
                        "title": "🔗 웹페이지 이동",
                        "description": "아래 버튼을 눌러주세요.",
                        "buttons": [{"action": "webLink", "label": button_label, "webLinkUrl": url}]
                    }
                }
            ]
        }
    }


def get_kst_now(): return datetime.datetime.utcnow() + datetime.timedelta(hours=9)


def get_server_host(): return os.environ.get("RENDER_EXTERNAL_URL", "http://localhost:8000")


@app.post("/kakao")
def kakao_bot_main(body: dict):
    user_id = body.get("userRequest", {}).get("user", {}).get("id", "알수없음")
    intent_name = body.get("intent", {}).get("name", "")
    params = body.get("action", {}).get("params", {})

    try:
        if intent_name == "연결 테스트": return make_kakao_response("연결 성공! 🚀\n\n✅ 서버가 정상 작동 중입니다.")

        # ==========================================
        # 🛡️ 유저 DB 사전 조회
        # ==========================================
        user_db = supabase.table("users").select("*").eq("kakao_id", user_id).execute()
        user_info = user_db.data[0] if user_db.data else None

        if intent_name == "회원가입":
            if user_info:
                return make_kakao_response(f"✅ 이미 가입된 회원입니다! ({user_info['name']}님)\n\n(정보 수정이 필요하다면 관리자에게 문의해 주세요.)")

            name, student_id, dept, gender = params.get("name"), params.get("student_id"), params.get(
                "department"), params.get("gender")
            supabase.table("users").upsert(
                {"kakao_id": user_id, "name": name, "student_id": student_id, "department": dept, "gender": gender,
                 "role": "member"}).execute()
            return make_kakao_response(
                f"환영합니다, {name}님! 🎉\n\n▪️ 이름: {name}\n▪️ 학번: {student_id}\n▪️ 소속: {dept}\n\n이제부터 챗봇을 이용하실 수 있습니다! 😊")

        if not user_info and intent_name in ["참석투표", "투표취소", "내상태", "발제문제출"]:
            return make_kakao_response("⛔ 아직 등록되지 않았습니다.\n먼저 [회원가입]을 진행해 주세요!")

        if intent_name in ["세미나생성", "조편성", "명단확인", "기록열람", "관리자", "수동관리", "발제문보기"]:
            if not user_info or user_info.get("role") != "admin":
                return make_kakao_response("⛔ 운영진(관리자) 전용 기능입니다.")

        # ==========================================
        # 📖 일반 부원 가이드
        # ==========================================
        if intent_name == "도움말":
            guide_msg = (
                "📖 [독서모임 챗봇 전체 사용 가이드] 📖\n\n"
                "동아리 활동을 편하게 관리할 수 있는 명령어 모음입니다!\n\n"
                "1️⃣ [참석투표]\n"
                "▪️ 매주 열리는 세미나(월/목) 중 원하는 요일을 선택해 선착순으로 자리를 찜합니다.\n"
                "▪️ 정원이 초과되면 자동으로 대기자로 등록됩니다.\n\n"
                "2️⃣ [투표취소]\n"
                "▪️ 사정이 생겼을 때 클릭하면 내 투표가 취소되며, 대기자가 있을 경우 자리가 자동으로 넘어갑니다.\n\n"
                "3️⃣ [내상태]\n"
                "▪️ 이번 주 내가 신청한 요일과, '참석 확정'인지 '대기'인지, 그리고 편성된 조를 확인합니다.\n\n"
                "4️⃣ [발제문제출]\n"
                "▪️ 참석 확정 부원 전용! 조장이 되기 위한 발제문을 웹페이지 폼에서 예쁘게 제출할 수 있습니다.\n\n"
                "👇 아래 버튼을 클릭하면 바로 실행됩니다!"
            )
            replies = [
                {"label": "참석투표", "action": "message", "messageText": "참석투표"},
                {"label": "투표취소", "action": "message", "messageText": "투표취소"},
                {"label": "내상태", "action": "message", "messageText": "내상태"},
                {"label": "발제문제출", "action": "message", "messageText": "발제문제출"}
            ]
            return make_kakao_response(guide_msg, replies)

        # ==========================================
        # 👑 관리자 전용 기능들
        # ==========================================
        if intent_name == "관리자":
            admin_guide = (
                "👑 [운영진 전체 명령어 가이드] 👑\n\n"
                "1️⃣ [세미나생성]\n"
                "▪️ 이번 주 세미나(월/목) 예약을 오픈합니다. (금 18:00 ~ 일 23:59 자동설정)\n\n"
                "2️⃣ [수동관리]\n"
                "▪️ 결석자 발생 시 관리자가 웹에서 직접 부원을 '추가/삭제'할 수 있습니다.\n\n"
                "3️⃣ [조편성]\n"
                "▪️ 1만 회 시뮬레이션으로 최적 조를 짜고 웹에서 핀셋 수정(이동)합니다.\n\n"
                "4️⃣ [발제문보기]\n"
                "▪️ 이번 주 제출된 발제문들을 웹에서 모아보고 워드 파일로 다운받습니다.\n\n"
                "5️⃣ [명단확인]\n"
                "▪️ 이번 주 확정자 명단을 카톡 텍스트로 뽑습니다.\n\n"
                "6️⃣ [기록열람]\n"
                "▪️ 종료된 과거 세미나들의 편성 기록과 발제문을 봅니다."
            )
            replies = [
                {"label": "세미나생성", "action": "message", "messageText": "세미나생성"},
                {"label": "조편성", "action": "message", "messageText": "조편성"},
                {"label": "수동관리", "action": "message", "messageText": "수동관리"},
                {"label": "발제문보기", "action": "message", "messageText": "발제문보기"}
            ]
            return make_kakao_response(admin_guide, replies)

        elif intent_name == "수동관리":
            url = f"{get_server_host()}/admin/manual_manage?admin_key={SUPABASE_KEY[:10]}"
            return make_kakao_link_response("✅ 이번 주 참석자를 수동으로 추가하거나 삭제할 수 있는 페이지입니다.", "🛠️ 수동 관리 웹 열기", url)

        elif intent_name == "발제문보기":
            url = f"{get_server_host()}/admin/current_topics?admin_key={SUPABASE_KEY[:10]}"
            return make_kakao_link_response("✅ 이번 주에 제출된 발제문들을 실시간으로 확인하고 워드로 다운받으세요!", "📚 현재 발제문 확인", url)

        elif intent_name == "기록열람":
            url = f"{get_server_host()}/admin/history?admin_key={SUPABASE_KEY[:10]}"
            return make_kakao_link_response("✅ 과거 기록 확인 및 워드 다운로드 페이지입니다.", "📋 기록 열람 페이지 오픈", url)

        elif intent_name == "세미나생성":
            week_name = params.get("week_name", "새로운 주차")

            # 🚨 [핵심] 동일한 주차 중복 생성 방지 로직!
            existing = supabase.table("seminars").select("id").eq("week_name", week_name).execute()
            if existing.data:
                return make_kakao_response(f"⛔ 이미 '{week_name}'(으)로 생성된 세미나가 존재합니다.\n다른 주차 이름을 입력해 주세요!")

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
            return make_kakao_response(f"✅ [{week_name}] 예약 오픈 세팅 완료!")

        elif intent_name == "명단확인":
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
                f"📋 [이번 주 확정 명단]\n\n[월요일]\n{', '.join(mon_list) if mon_list else '없음'}\n\n[목요일]\n{', '.join(thu_list) if thu_list else '없음'}")

        elif intent_name == "조편성":
            utterance = body.get("userRequest", {}).get("utterance", "")
            if "월" in utterance:
                day_choice = "월요일"
            elif "목" in utterance:
                day_choice = "목요일"
            else:
                replies = [{"label": "월요일 조편성", "action": "message", "messageText": "월요일 조편성"},
                           {"label": "목요일 조편성", "action": "message", "messageText": "목요일 조편성"}]
                return make_kakao_response("어느 요일 조를 편성할까요?\n아래 버튼을 선택해주세요.", replies)

            seminar_res = supabase.table("seminars").select("*").eq("is_active", True).eq("day", day_choice).execute()
            if not seminar_res.data: return make_kakao_response(f"활성화된 {day_choice} 세미나가 없습니다.")
            target_seminar = seminar_res.data[0];
            target_id = target_seminar["id"]

            att_res = supabase.table("attendances").select("*").eq("seminar_id", target_id).eq("status",
                                                                                               "attending").execute()
            if len(att_res.data) < 2: return make_kakao_response("참석자가 너무 적어 조 편성이 불가능합니다.")

            users_res = supabase.table("users").select("kakao_id, name, gender").in_("kakao_id", [a["kakao_id"] for a in
                                                                                                  att_res.data]).execute()
            user_dict = {u["kakao_id"]: u for u in users_res.data}
            attendees = [{"att_id": a["id"], "kakao_id": a["kakao_id"], "name": user_dict[a["kakao_id"]]["name"],
                          "gender": user_dict[a["kakao_id"]]["gender"], "is_fac": a.get("is_facilitator", False)} for a
                         in att_res.data if a["kakao_id"] in user_dict]

            forty_days_ago = (get_kst_now() - datetime.timedelta(days=40)).isoformat()
            past_res = supabase.table("attendances").select("seminar_id, team_name, kakao_id, created_at").neq(
                "team_name", "None").gte("created_at", forty_days_ago).execute()
            now = get_kst_now()

            past_groups = defaultdict(list);
            seminar_dates = {}
            for p in past_res.data:
                if p["team_name"]:
                    key = f"{p['seminar_id']}_{p['team_name']}"
                    past_groups[key].append(p["kakao_id"])
                    if key not in seminar_dates:
                        try:
                            seminar_dates[key] = datetime.datetime.fromisoformat(p["created_at"].replace('Z', '+00:00'))
                        except:
                            seminar_dates[key] = now

            pair_max_penalty = defaultdict(int)
            for key, members in past_groups.items():
                days_ago = (now.replace(tzinfo=None) - seminar_dates[key].replace(tzinfo=None)).days
                if days_ago <= 7:
                    penalty = 50
                elif days_ago <= 14:
                    penalty = 30
                elif days_ago <= 21:
                    penalty = 15
                elif days_ago <= 28:
                    penalty = 5
                else:
                    penalty = 1
                for i in range(len(members)):
                    for j in range(i + 1, len(members)):
                        pair = tuple(sorted([members[i], members[j]]));
                        pair_max_penalty[pair] = max(pair_max_penalty[pair], penalty)

            best_teams, best_score = [], float('inf')
            facs = [a for a in attendees if a["is_fac"]];
            norms = [a for a in attendees if not a["is_fac"]]
            num_teams = max(1, len(attendees) // 4);
            team_sizes = [4] * num_teams
            for i in range(len(attendees) % 4): team_sizes[i] += 1

            for _ in range(5000):  # 5000번 최적화로 타임아웃 방지
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
                            score += pair_max_penalty[pair]
                if score < best_score: best_score, best_teams = score, teams
                if score == 0: break

            report = f"✅ [{target_seminar['week_name']} - {day_choice}] 조 편성 초안\n\n"
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
                        if pair in pair_max_penalty and pair_max_penalty[pair] > 1: warnings.append(
                            f"{team[idx1]['name']}&{team[idx2]['name']}")
                if warnings: report += f"⚠️ 특이사항(최근만남): {', '.join(warnings)}\n\n"

                for m_att in team: supabase.table("attendances").update({"team_name": t_name}).eq("id", m_att[
                    "att_id"]).execute()

            url = f"{get_server_host()}/admin/edit_teams?seminar_id={target_id}&day={day_choice}&admin_key={SUPABASE_KEY[:10]}"
            return make_kakao_link_response(report, "🛠️ 조원 핀셋 수정하기", url)

        # ==========================================
        # 🔵 일반 부원 전용 기능들
        # ==========================================
        elif intent_name == "참석투표":
            utterance = body.get("userRequest", {}).get("utterance", "")
            if "월" in utterance:
                day_choice = "월요일"
            elif "목" in utterance:
                day_choice = "목요일"
            else:
                replies = [{"label": "월요일 신청", "action": "message", "messageText": "월요일 참석투표"},
                           {"label": "목요일 신청", "action": "message", "messageText": "목요일 참석투표"}]
                return make_kakao_response("어느 요일 세미나에 참석하시겠습니까?\n아래 버튼을 선택해주세요.", replies)

            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
            now_str = get_kst_now().isoformat()
            if now_str < seminars_res.data[0]["open_time"]: return make_kakao_response("아직 투표 기간이 아닙니다!")
            if now_str > seminars_res.data[0]["close_time"]: return make_kakao_response("투표가 마감되었습니다.")

            active_ids = [s["id"] for s in seminars_res.data]
            existing_vote = supabase.table("attendances").select("*").in_("seminar_id", active_ids).eq("kakao_id",
                                                                                                       user_id).execute()
            if existing_vote.data: return make_kakao_response("이미 투표하셨습니다!\n('내 상태' 버튼을 눌러 확인해 보세요.)")

            target_seminar = next((s for s in seminars_res.data if s["day"] == day_choice), None)
            if not target_seminar: return make_kakao_response(f"활성화된 {day_choice} 세미나가 없습니다.")

            attendees = supabase.table("attendances").select("id", count="exact").eq("seminar_id",
                                                                                     target_seminar["id"]).eq("status",
                                                                                                              "attending").execute()
            current_count = attendees.count if attendees.count else 0

            status = "attending" if current_count < target_seminar["capacity"] else "waitlisted"
            supabase.table("attendances").insert(
                {"seminar_id": target_seminar["id"], "kakao_id": user_id, "status": status}).execute()

            status_text = "참석 확정 🎉" if status == "attending" else "대기자 등록 ⏳ (취소자 발생 시 승급)"
            return make_kakao_response(
                f"✅ 투표 완료!\n▪️ 주차: {target_seminar['week_name']}\n▪️ 요일: {day_choice}\n▪️ 상태: {status_text}")

        elif intent_name == "투표취소":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")

            my_vote = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("kakao_id",
                                                                                                             user_id).execute()
            if not my_vote.data: return make_kakao_response("이번 주에 취소할 투표 내역이 없습니다.")
            vote_data = my_vote.data[0]

            supabase.table("attendances").delete().eq("id", vote_data["id"]).execute()

            if vote_data["status"] == "attending":
                next_in_line = supabase.table("attendances").select("id").eq("seminar_id", vote_data["seminar_id"]).eq(
                    "status", "waitlisted").order("created_at").limit(1).execute()
                if next_in_line.data: supabase.table("attendances").update({"status": "attending"}).eq("id",
                                                                                                       next_in_line.data[
                                                                                                           0][
                                                                                                           "id"]).execute()

            return make_kakao_response("✅ 참석 취소가 완료되었습니다.\n대기자가 있었다면 자리가 양보되었습니다.")

        elif intent_name == "내상태":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
            my_vote = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("kakao_id",
                                                                                                             user_id).execute()
            if not my_vote.data: return make_kakao_response("이번 주 세미나에 아직 투표하지 않으셨습니다!")

            target_seminar = next(s for s in seminars_res.data if s["id"] == my_vote.data[0]["seminar_id"])
            status_kr = "✅ 참석 확정" if my_vote.data[0]["status"] == "attending" else "⏳ 대기 중"
            team_info = f"\n▪️ 소속 조: {my_vote.data[0].get('team_name', '아직 편성 안 됨')}" if my_vote.data[0][
                                                                                             "status"] == "attending" else ""
            return make_kakao_response(
                f"[{target_seminar['week_name']}]\n▪️ 예약 요일: {target_seminar['day']}\n▪️ 현재 상태: {status_kr}{team_info}")

        elif intent_name == "발제문제출":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("활성화된 세미나가 없습니다.")
            my_vote = supabase.table("attendances").select("id, seminar_id").in_("seminar_id", [s["id"] for s in
                                                                                                seminars_res.data]).eq(
                "kakao_id", user_id).eq("status", "attending").execute()
            if not my_vote.data: return make_kakao_response("이번 주 '참석 확정' 부원만 제출 가능합니다.")

            target_seminar = next(s for s in seminars_res.data if s["id"] == my_vote.data[0]["seminar_id"])
            url = f"{get_server_host()}/submit_topic?att_id={my_vote.data[0]['id']}"
            return make_kakao_link_response(
                f"✅ {user_info['name']}님, 발제문 제출 링크를 준비했습니다!\n\n'{target_seminar['week_name']} {target_seminar['day']}' 세미나의 발제 내용을 정성껏 작성해 주세요.",
                "📝 발제 폼으로 이동하기", url)

        else:
            return make_kakao_response("준비되지 않은 기능입니다.")

    except Exception as e:
        print(f"Error: {e}")
        return make_kakao_response("서버 처리 중 오류가 발생했습니다.")


# ==========================================
# 🌐 웹페이지 라우팅
# ==========================================
@app.get("/submit_topic", response_class=HTMLResponse)
def submit_topic_page(request: Request, att_id: int):
    return templates.TemplateResponse("topic_submit.html", {"request": request, "att_id": att_id})


@app.post("/api/submit_topic")
def api_submit_topic(data: dict):
    try:
        att_id = data.pop('att_id')
        topic_json = {"book_title": data.get('book_title'), "author": data.get('author'), "range": data.get('range'),
                      "questions": [data.get('q1'), data.get('q2'), data.get('q3')]}
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
    return templates.TemplateResponse("admin_edit_teams.html", {"request": request, "day": day, "attendees": attendees})


@app.post("/api/admin/save_teams")
def api_save_teams(body: dict):
    try:
        for u in body.get("updates", []): supabase.table("attendances").update({"team_name": u["team_name"]}).eq("id",
                                                                                                                 u[
                                                                                                                     "att_id"]).execute()
        return {"status": "success"}
    except Exception as e:
        return {"status": "error", "message": str(e)}


@app.get("/admin/history", response_class=HTMLResponse)
def admin_history_page(request: Request, admin_key: str):
    if admin_key != SUPABASE_KEY[:10]: return HTMLResponse("⛔ 권한 없음")
    seminars_res = supabase.table("seminars").select("*").eq("is_active", False).order("created_at", desc=True).limit(
        10).execute()
    seminar_ids = [s["id"] for s in seminars_res.data]
    if not seminar_ids: return templates.TemplateResponse("admin_history.html", {"request": request, "history_data": [],
                                                                                 "admin_key": admin_key})

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
        for att in sem_att: teams[att["team_name"]].append(
            {"name": user_dict.get(att["kakao_id"], "알수없음"), "is_fac": att["is_facilitator"],
             "topic": att["topic_content"]})
        history_data.append({"sem_id": sem_id, "week": sem["week_name"], "day": sem["day"],
                             "date": datetime.datetime.fromisoformat(sem["created_at"]).strftime('%Y-%m-%d'),
                             "teams": dict(sorted(teams.items()))})

    return templates.TemplateResponse("admin_history.html",
                                      {"request": request, "history_data": history_data, "admin_key": admin_key})


# 📄 워드 파일 다운로드
@app.get("/admin/history/{seminar_id}/download_word")
def download_topics_word(seminar_id: int, admin_key: str):
    if admin_key != SUPABASE_KEY[:10]: return HTMLResponse("⛔ 권한 없음")
    try:
        seminar = supabase.table('seminars').select('*').eq('id', seminar_id).single().execute().data
        att_res = supabase.table('attendances').select('*').eq('seminar_id', seminar_id).eq('is_facilitator',
                                                                                            True).execute()
        users_res = supabase.table('users').select('kakao_id, name, department').in_('kakao_id', [a['kakao_id'] for a in
                                                                                                  att_res.data]).execute()
        user_dict = {u['kakao_id']: u for u in users_res.data}

        template_path = os.path.join(app.root_path, 'templates', 'template.docx')
        doc = DocxTemplate(template_path)

        submissions_list = []
        for sub in att_res.data:
            user = user_dict.get(sub['kakao_id'], {})
            topic_data = sub.get('topic_content', {}) or {}
            topics_list = [{'topic': q, 'page': topic_data.get('range', ''), 'reference': ''} for q in
                           topic_data.get('questions', []) if q]
            submissions_list.append({'department': user.get('department', ''), 'author_name': user.get('name', ''),
                                     'book_title': topic_data.get('book_title', ''),
                                     'book_author': topic_data.get('author', ''), 'topics': topics_list})

        context = {'meeting_date': f"{seminar['week_name']} ({seminar['day']})", 'book_title': "이번 주 세미나 통합 발제문",
                   'submissions': submissions_list}
        doc.render(context)

        f = BytesIO();
        doc.save(f);
        f.seek(0)
        filename = f"발제문_{seminar['week_name']}_{seminar['day']}.docx"
        return StreamingResponse(f,
                                 media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                 headers={
                                     "Content-Disposition": f"attachment; filename={filename.encode('utf-8').decode('latin1')}"})
    except Exception as e:
        return HTMLResponse(f"문서 생성 중 오류: {e}")


# ==========================================
# 🛠️ 수동 출석 관리 (추가/삭제)
# ==========================================
@app.get("/admin/manual_manage", response_class=HTMLResponse)
def admin_manual_manage(request: Request, admin_key: str, day: str = "월요일"):
    if admin_key != SUPABASE_KEY[:10]: return HTMLResponse("⛔ 권한 없음")

    seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
    if not seminars_res.data: return HTMLResponse("활성화된 세미나가 없습니다.")

    target_seminar = next((s for s in seminars_res.data if s["day"] == day), None)

    all_users = supabase.table("users").select("kakao_id, name").execute().data
    att_res = supabase.table("attendances").select("*").eq("seminar_id", target_seminar["id"]).eq("status",
                                                                                                  "attending").execute()

    att_kakao_ids = {a["kakao_id"]: a["id"] for a in att_res.data}
    attendees = [{"att_id": att_kakao_ids[u["kakao_id"]], "name": u["name"]} for u in all_users if
                 u["kakao_id"] in att_kakao_ids]
    non_attendees = [{"kakao_id": u["kakao_id"], "name": u["name"]} for u in all_users if
                     u["kakao_id"] not in att_kakao_ids]

    return templates.TemplateResponse("admin_manual.html", {
        "request": request, "admin_key": admin_key, "day": day,
        "seminar_id": target_seminar["id"], "attendees": attendees, "non_attendees": non_attendees
    })


@app.post("/api/admin/manual_add")
def api_manual_add(body: dict):
    try:
        supabase.table("attendances").insert(
            {"seminar_id": body["seminar_id"], "kakao_id": body["kakao_id"], "status": "attending"}).execute()
        return {"status": "success"}
    except Exception as e:
        return {"status": "error"}


@app.post("/api/admin/manual_delete")
def api_manual_delete(body: dict):
    try:
        supabase.table("attendances").delete().eq("id", body["att_id"]).execute()
        return {"status": "success"}
    except Exception as e:
        return {"status": "error"}


# ==========================================
# 📚 현재 발제문 모아보기
# ==========================================
@app.get("/admin/current_topics", response_class=HTMLResponse)
def admin_current_topics(request: Request, admin_key: str, day: str = "월요일"):
    if admin_key != SUPABASE_KEY[:10]: return HTMLResponse("⛔ 권한 없음")

    seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
    if not seminars_res.data: return HTMLResponse("활성화된 세미나가 없습니다.")
    target_seminar = next((s for s in seminars_res.data if s["day"] == day), None)

    att_res = supabase.table('attendances').select('*').eq('seminar_id', target_seminar["id"]).eq('is_facilitator',
                                                                                                  True).execute()
    users_res = supabase.table('users').select('kakao_id, name, department').in_('kakao_id', [a['kakao_id'] for a in
                                                                                              att_res.data]).execute()
    user_dict = {u['kakao_id']: u for u in users_res.data}

    submissions = []
    for sub in att_res.data:
        user = user_dict.get(sub['kakao_id'], {})
        topic_data = sub.get('topic_content', {}) or {}
        topics_list = [{'topic': q, 'page': topic_data.get('range', ''), 'reference': ''} for q in
                       topic_data.get('questions', []) if q]

        submissions.append({
            'author_name': user.get('name', '알수없음'),
            'department': user.get('department', ''),
            'created_at': sub.get('created_at', ''),
            'topics': topics_list
        })

    event_data = {
        "id": target_seminar["id"],
        "book_title": f"[{target_seminar['week_name']}] {day} 세미나 발제문",
        "meeting_date": target_seminar["open_time"][:10]
    }

    return templates.TemplateResponse("admin_topic_view.html", {
        "request": request, "event": event_data, "submissions": submissions, "admin_key": admin_key, "day": day
    })