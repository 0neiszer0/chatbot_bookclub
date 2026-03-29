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
        if intent_name == "연결 테스트": return make_kakao_response("연결 성공! 🚀\n\n✅ 서버 연결 테스트가 정상적으로 완료되었습니다.")

        # ==========================================
        # 🛡️ [핵심 가드레일] 유저 DB 사전 조회
        # ==========================================
        user_db = supabase.table("users").select("*").eq("kakao_id", user_id).execute()
        user_info = user_db.data[0] if user_db.data else None

        # 1. 회원가입 로직 (여기는 오픈빌더 파라미터가 정상 작동해야 함)
        if intent_name == "회원가입":
            if user_info:
                return make_kakao_response(f"✅ 이미 가입된 회원입니다! ({user_info['name']}님)\n\n(정보 수정이 필요하다면 관리자에게 문의해 주세요.)")

            name, student_id, dept, gender = params.get("name"), params.get("student_id"), params.get(
                "department"), params.get("gender")
            supabase.table("users").upsert(
                {"kakao_id": user_id, "name": name, "student_id": student_id, "department": dept, "gender": gender,
                 "role": "member"}).execute()

            success_msg = (
                f"환영합니다, {name}님! 🎉\n"
                f"입력해주신 소중한 정보로 동아리원 등록이 정상적으로 완료되었습니다.\n\n"
                f"📝 [반영된 등록 정보]\n"
                f"▪️ 이름: {name}\n"
                f"▪️ 학번: {student_id}\n"
                f"▪️ 소속: {dept}\n"
                f"▪️ 성별: {gender}\n\n"
                f"이제부터 하단 메뉴를 통해 투표와 발제문 제출을 이용하실 수 있습니다! 😊"
            )
            return make_kakao_response(success_msg)

        if not user_info and intent_name in ["참석투표", "투표취소", "내상태", "발제문제출"]:
            return make_kakao_response("⛔ 아직 동아리원으로 등록되지 않았습니다.\n먼저 [회원가입]을 진행해 주세요!")

        if intent_name in ["세미나생성", "조편성", "명단확인", "기록열람", "관리자"]:
            if not user_info or user_info.get("role") != "admin":
                return make_kakao_response("⛔ 운영진(관리자) 전용 기능입니다.")

        # ==========================================
        # 📖 일반 부원 가이드
        # ==========================================
        if intent_name == "도움말":
            guide_msg = (
                "📖 [독서모임 챗봇 이용 가이드] 📖\n\n"
                "환영합니다! 챗봇을 통해 동아리 활동을 편하게 신청하고 관리하세요.\n\n"
                "1️⃣ 참석투표\n"
                "▪️ 매주 오픈되는 세미나(월/목) 중 원하는 요일을 선착순으로 신청합니다.\n\n"
                "2️⃣ 투표취소\n"
                "▪️ 사정이 생겨 못 오실 때 꼭 눌러주세요! 대기자에게 자리가 양보됩니다.\n\n"
                "3️⃣ 내상태\n"
                "▪️ 나의 투표 현황과 이번 주 배정된 조를 확인합니다.\n\n"
                "4️⃣ 발제문제출\n"
                "▪️ 참석이 확정된 후, 세련된 웹페이지에서 발제 내용을 편하게 제출합니다.\n\n"
                "👇 아래 버튼을 누르거나 채팅창 하단 메뉴를 이용해 보세요!"
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
                "👑 [운영진 전용 명령어 가이드] 👑\n\n"
                "동아리 운영을 100% 자동화해 주는 마법의 명령어들입니다. 채팅창에 아래 명령어들을 편하게 입력해 보세요!\n\n"
                "1️⃣ 세미나생성\n"
                "▪️ 기능: 새로운 주차의 세미나(월/목) 투표를 엽니다.\n"
                "▪️ 방법: '세미나생성' ➡️ 안내에 따라 주차 입력\n\n"
                "2️⃣ 조편성\n"
                "▪️ 기능: 1만 회 시뮬레이션으로 최적의 조를 짜고 리포트를 줍니다.\n"
                "▪️ 방법: '조편성' ➡️ 요일 선택\n"
                "▪️ 💡 팁: 결과 아래의 [🛠️ 조원 수정하기] 버튼을 누르면 웹에서 쉽게 핀셋 교체가 가능합니다!\n\n"
                "3️⃣ 명단확인\n"
                "▪️ 기능: 이번 주 월/목 참석 확정자 명단을 쫙 뽑아줍니다.\n"
                "▪️ 방법: '명단확인' 입력\n\n"
                "4️⃣ 기록열람\n"
                "▪️ 기능: 역대 세미나의 조 편성 및 제출된 발제 기록을 한눈에 봅니다. (워드 다운로드 지원)\n"
                "▪️ 방법: '기록열람' 입력\n\n"
                "운영진 여러분, 이번 주도 화이팅입니다! 💪"
            )
            replies = [
                {"label": "세미나생성", "action": "message", "messageText": "세미나생성"},
                {"label": "조편성", "action": "message", "messageText": "조편성"},
                {"label": "명단확인", "action": "message", "messageText": "명단확인"},
                {"label": "기록열람", "action": "message", "messageText": "기록열람"}
            ]
            return make_kakao_response(admin_guide, replies)

        elif intent_name == "세미나생성":
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

            admin_msg = (
                f"✅ 관리자님, 요청하신 세미나 생성이 완벽하게 처리되었습니다!\n\n"
                f"📝 [반영된 세팅 내역]\n"
                f"▪️ 주차: {week_name}\n"
                f"▪️ 요일: 월요일 / 목요일 동시 오픈\n"
                f"▪️ 정원: 각 세션당 30명\n"
                f"▪️ 투표 오픈: {open_time.strftime('%m/%d(금) 18:00')}\n"
                f"▪️ 투표 마감: {close_time.strftime('%m/%d(일) 23:59')}"
            )
            return make_kakao_response(admin_msg)

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
                f"📋 [이번 주 참석 확정 명단]\n\n[월요일]\n{', '.join(mon_list) if mon_list else '없음'}\n\n[목요일]\n{', '.join(thu_list) if thu_list else '없음'}")

        elif intent_name == "기록열람":
            history_url = f"{get_server_host()}/admin/history?admin_key={SUPABASE_KEY[:10]}"
            replies = [{"label": "📋 기록 열람 페이지 오픈", "action": "webLink", "webLinkUrl": history_url}]
            return make_kakao_response("✅ 과거 기록 열람 페이지가 준비되었습니다.\n\n아래 버튼을 눌러 지난 주차들의 조 편성 및 발제 기록을 편하게 확인하세요!",
                                       replies)

        # 💡 [핵심 스마트 파싱 로직 적용] 조편성
        elif intent_name == "조편성":
            utterance = body.get("userRequest", {}).get("utterance", "")
            if "월" in utterance:
                day_choice = "월요일"
            elif "목" in utterance:
                day_choice = "목요일"
            else:
                # 사용자가 요일을 안 쳤을 때, 직접 만든 퀵 리플라이 버튼 제공
                replies = [
                    {"label": "월요일 조편성", "action": "message", "messageText": "월요일 조편성"},
                    {"label": "목요일 조편성", "action": "message", "messageText": "목요일 조편성"}
                ]
                return make_kakao_response("어느 요일 조를 편성할까요?\n아래 버튼을 선택해주세요.", replies)

            seminar_res = supabase.table("seminars").select("*").eq("is_active", True).eq("day", day_choice).execute()
            if not seminar_res.data: return make_kakao_response(f"활성화된 {day_choice} 세미나가 없습니다.")
            target_seminar = seminar_res.data[0]
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

            past_groups = defaultdict(list)
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
                        pair = tuple(sorted([members[i], members[j]]))
                        pair_max_penalty[pair] = max(pair_max_penalty[pair], penalty)

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
                            score += pair_max_penalty[pair]
                if score < best_score: best_score, best_teams = score, teams
                if score == 0: break

            report = f"✅ 요청하신 [{target_seminar['week_name']} - {day_choice}] 조 편성이 1만 회 시뮬레이션을 통해 최적화 완료되었습니다!\n(성비 및 과거 만남 차감형 페널티 적용됨)\n\n✨ [조 편성 초안 리포트]\n\n"
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
                        if pair in pair_max_penalty and pair_max_penalty[pair] > 1:
                            warnings.append(f"{team[idx1]['name']}&{team[idx2]['name']}")
                if warnings:
                    report += f"⚠️ 특이사항(최근만남): {', '.join(warnings)}\n\n"
                else:
                    report += f"⚠️ 특이사항: 최근 중복자 없음!\n\n"

                for m_att in team: supabase.table("attendances").update({"team_name": t_name}).eq("id", m_att[
                    "att_id"]).execute()

            edit_url = f"{get_server_host()}/admin/edit_teams?seminar_id={target_id}&day={day_choice}&admin_key={SUPABASE_KEY[:10]}"
            replies = [{"label": "🛠️ 조원 핀셋 수정하기", "action": "webLink", "webLinkUrl": edit_url}]
            return make_kakao_response(report, replies)

        # ==========================================
        # 🔵 일반 부원 전용 기능들
        # ==========================================
        # 💡 [핵심 스마트 파싱 로직 적용] 참석투표
        elif intent_name == "참석투표":
            utterance = body.get("userRequest", {}).get("utterance", "")
            if "월" in utterance:
                day_choice = "월요일"
            elif "목" in utterance:
                day_choice = "목요일"
            else:
                # 요일이 없으면 자체 버튼 띄우기
                replies = [
                    {"label": "월요일 신청", "action": "message", "messageText": "월요일 참석투표"},
                    {"label": "목요일 신청", "action": "message", "messageText": "목요일 참석투표"}
                ]
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

            status_text = "참석 확정 🎉" if status == "attending" else "대기자 등록 ⏳ (취소자 발생 시 자동 승급)"
            vote_msg = (
                f"✅ 소중한 투표가 정상적으로 서버에 반영되었습니다!\n\n"
                f"📝 [나의 투표 내역]\n"
                f"▪️ 주차: {target_seminar['week_name']}\n"
                f"▪️ 선택 요일: {day_choice}\n"
                f"▪️ 현재 상태: {status_text}"
            )
            return make_kakao_response(vote_msg)

        elif intent_name == "투표취소":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
            active_ids = [s["id"] for s in seminars_res.data]

            my_vote = supabase.table("attendances").select("*").in_("seminar_id", active_ids).eq("kakao_id",
                                                                                                 user_id).execute()
            if not my_vote.data: return make_kakao_response("이번 주에 취소할 투표 내역이 없습니다.")
            vote_data = my_vote.data[0]

            target_seminar = next((s for s in seminars_res.data if s["id"] == vote_data["seminar_id"]), None)
            voted_day = target_seminar["day"] if target_seminar else "해당 요일"
            week_name = target_seminar["week_name"] if target_seminar else "이번 주"

            supabase.table("attendances").delete().eq("id", vote_data["id"]).execute()

            if vote_data["status"] == "attending":
                next_in_line = supabase.table("attendances").select("id").eq("seminar_id", vote_data["seminar_id"]).eq(
                    "status", "waitlisted").order("created_at").limit(1).execute()
                if next_in_line.data: supabase.table("attendances").update({"status": "attending"}).eq("id",
                                                                                                       next_in_line.data[
                                                                                                           0][
                                                                                                           "id"]).execute()

            cancel_msg = (
                f"✅ 참석 취소 처리가 정상적으로 완료되었습니다.\n\n"
                f"📝 [취소 내역]\n"
                f"▪️ 주차: {week_name}\n"
                f"▪️ 취소 요일: {voted_day}\n\n"
                f"명단에서 성공적으로 제외되었으며, 대기자가 있었다면 다른 부원에게 자리가 양보되었습니다. 다음 세미나 때 꼭 함께해요! 😊"
            )
            return make_kakao_response(cancel_msg)

        elif intent_name == "내상태":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("진행 중인 세미나가 없습니다.")
            my_vote = supabase.table("attendances").select("*").in_("seminar_id",
                                                                    [s["id"] for s in seminars_res.data]).eq("kakao_id",
                                                                                                             user_id).execute()
            if not my_vote.data: return make_kakao_response(f"👋 {user_info['name']}님, 이번 주 세미나에 아직 투표하지 않으셨습니다!")

            target_seminar = next(s for s in seminars_res.data if s["id"] == my_vote.data[0]["seminar_id"])
            status_kr = "✅ 참석 확정" if my_vote.data[0]["status"] == "attending" else "⏳ 대기 중"
            team_info = f"\n▪️ 소속 조: {my_vote.data[0].get('team_name', '아직 편성 안 됨')}" if my_vote.data[0][
                                                                                             "status"] == "attending" else ""

            return make_kakao_response(
                f"[{target_seminar['week_name']}]\n▪️ 예약 요일: {target_seminar['day']}\n▪️ 현재 상태: {status_kr}{team_info}\n\n사정이 생기면 '참석 취소'를 진행해주세요.")

        elif intent_name == "발제문제출":
            seminars_res = supabase.table("seminars").select("*").eq("is_active", True).execute()
            if not seminars_res.data: return make_kakao_response("활성화된 세미나가 없습니다.")
            active_ids = [s["id"] for s in seminars_res.data]
            my_vote = supabase.table("attendances").select("id, seminar_id").in_("seminar_id", active_ids).eq(
                "kakao_id", user_id).eq("status", "attending").execute()
            if not my_vote.data: return make_kakao_response("이번 주 '참석 확정' 부원만 발제 제출이 가능합니다. (대기자는 제출 불가)")

            target_seminar = next(s for s in seminars_res.data if s["id"] == my_vote.data[0]["seminar_id"])
            submit_url = f"{get_server_host()}/submit_topic?att_id={my_vote.data[0]['id']}"
            replies = [{"label": "📝 발제 폼으로 이동하기", "action": "webLink", "webLinkUrl": submit_url}]

            return make_kakao_response(
                f"✅ {user_info['name']}님, 발제문 제출 링크를 준비했습니다!\n\n아래 버튼을 눌러 '{target_seminar['week_name']} {target_seminar['day']}' 세미나의 발제 내용을 정성껏 작성해 주세요. 😊",
                replies)

        else:
            return make_kakao_response("준비되지 않은 기능입니다.")

    except Exception as e:
        print(f"Error: {e}")
        return make_kakao_response("서버 처리 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.")


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

    if not seminar_ids:
        return templates.TemplateResponse("admin_history.html",
                                          {"request": request, "history_data": [], "admin_key": admin_key})

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
            teams[att["team_name"]].append(
                {"name": user_dict.get(att["kakao_id"], "알수없음"), "is_fac": att["is_facilitator"],
                 "topic": att["topic_content"]})

        history_data.append({
            "sem_id": sem_id,
            "week": sem["week_name"],
            "day": sem["day"],
            "date": datetime.datetime.fromisoformat(sem["created_at"]).strftime('%Y-%m-%d'),
            "teams": dict(sorted(teams.items()))
        })

    return templates.TemplateResponse("admin_history.html",
                                      {"request": request, "history_data": history_data, "admin_key": admin_key})


# ==========================================
# 📄 워드(Word) 파일 다운로드 라우터
# ==========================================
@app.get("/admin/history/{seminar_id}/download_word")
def download_topics_word(seminar_id: int, admin_key: str):
    if admin_key != SUPABASE_KEY[:10]: return HTMLResponse("⛔ 권한 없음")
    try:
        seminar_res = supabase.table('seminars').select('*').eq('id', seminar_id).single().execute()
        seminar = seminar_res.data

        att_res = supabase.table('attendances').select('*').eq('seminar_id', seminar_id).eq('is_facilitator',
                                                                                            True).execute()
        users_res = supabase.table('users').select('kakao_id, name, department').in_('kakao_id', [a['kakao_id'] for a in
                                                                                                  att_res.data]).execute()
        user_dict = {u['kakao_id']: u for u in users_res.data}

        template_path = os.path.join(app.root_path, 'templates', 'template.docx')
        if not os.path.exists(template_path):
            return HTMLResponse("⛔ 템플릿 워드 파일(template.docx)을 templates 폴더에 넣어주세요!", status_code=404)

        doc = DocxTemplate(template_path)

        submissions_list = []
        for sub in att_res.data:
            user = user_dict.get(sub['kakao_id'], {})
            topic_data = sub.get('topic_content', {}) or {}

            topics_list = []
            for q in topic_data.get('questions', []):
                if q: topics_list.append({'topic': q, 'page': topic_data.get('range', ''), 'reference': ''})

            submissions_list.append({
                'department': user.get('department', ''),
                'author_name': user.get('name', ''),
                'book_title': topic_data.get('book_title', ''),
                'book_author': topic_data.get('author', ''),
                'topics': topics_list
            })

        context = {
            'meeting_date': f"{seminar['week_name']} ({seminar['day']})",
            'book_title': "이번 주 세미나 통합 발제문",
            'submissions': submissions_list
        }

        doc.render(context)

        f = BytesIO()
        doc.save(f)
        f.seek(0)

        filename = f"발제문_{seminar['week_name']}_{seminar['day']}.docx"

        return StreamingResponse(
            f,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename={filename.encode('utf-8').decode('latin1')}"}
        )
    except Exception as e:
        print(f"Word download error: {e}")
        return HTMLResponse("문서 생성 중 오류 발생.")