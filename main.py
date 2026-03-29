import os
from fastapi import FastAPI, Request
from supabase import create_client, Client
from dotenv import load_dotenv

# .env 파일에서 환경 변수 불러오기
load_dotenv()

app = FastAPI()

# 코드에 직접 적지 않고, 환경 변수에서 가져옵니다!
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

# Supabase 클라이언트 연결
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)


@app.post("/kakao")
async def kakao_response(request: Request):
    body = await request.json()
    print("수신된 데이터:", body)

    # 1. 카카오톡 유저 고유 ID 가져오기 (이 ID로 사람을 구분합니다)
    user_request = body.get("userRequest", {})
    user_id = user_request.get("user", {}).get("id", "알수없음")

    # 2. 어떤 블록(의도)에서 온 요청인지 확인
    intent_name = body.get("intent", {}).get("name", "")

    # 3. 챗봇에서 보낸 파라미터(사용자 입력값) 가져오기
    action = body.get("action", {})
    params = action.get("params", {})

    # ==========================================
    # 🟢 [기능 1] 회원가입 처리 로직
    # ==========================================
    if intent_name == "회원가입":
        # 파라미터에서 정보 추출
        name = params.get("name")
        student_id = params.get("student_id")
        department = params.get("department")
        gender = params.get("gender")

        try:
            # Supabase DB의 'users' 테이블에 데이터 넣기 (upsert는 덮어쓰기/새로넣기 모두 가능)
            supabase.table("users").upsert({
                "kakao_id": user_id,
                "name": name,
                "student_id": student_id,
                "department": department,
                "gender": gender
            }).execute()

            response_text = f"환영합니다, {name}님! 🎉\n학번: {student_id}\n학과: {department}\n성별: {gender}\n\n성공적으로 동아리원으로 등록되었습니다. 이제 세미나 투표가 가능합니다!"

        except Exception as e:
            print("DB 저장 에러:", e)
            response_text = "회원가입 중 오류가 발생했습니다. 관리자에게 문의해주세요."

    # ==========================================
    # 🟡 [기능 2] 연결 테스트 로직 (기존 기능 유지)
    # ==========================================
    elif intent_name == "연결 테스트":
        response_text = "안녕! 파이썬 서버와 성공적으로 연결되었어! 🚀"

    else:
        response_text = "아직 준비되지 않은 기능입니다."

    # 카카오톡 응답 양식으로 포장해서 보내기
    response = {
        "version": "2.0",
        "template": {
            "outputs": [
                {
                    "simpleText": {
                        "text": response_text
                    }
                }
            ]
        }
    }
    return response