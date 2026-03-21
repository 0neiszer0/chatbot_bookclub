from fastapi import FastAPI, Request

app = FastAPI()

# 카카오톡 오픈빌더에서 우리 서버로 신호를 보낼 주소 (엔드포인트)
@app.post("/kakao")
async def kakao_response(request: Request):
    # 카카오에서 보낸 데이터 읽기
    body = await request.json()
    print("수신된 데이터:", body) # 나중에 에러 확인할 때 유용합니다.

    # 카카오톡 챗봇으로 다시 보낼 응답 양식 (카카오 i 오픈빌더 규격)
    response = {
        "version": "2.0",
        "template": {
            "outputs": [
                {
                    "simpleText": {
                        "text": "안녕! 파이썬 서버와 성공적으로 연결되었어!! 🚀"
                    }
                }
            ]
        }
    }
    return response