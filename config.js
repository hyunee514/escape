// ★ 선생님이 한 번만 설정하는 파일이에요. (만든이: 성현쌤)
// 1) 구글 앱스 스크립트에서 복사한 '웹 앱 URL'을 SERVER_URL 따옴표 안에 붙여 넣으세요.
//    ※ 전기기능사 RPG 서버와는 별개로, '설비보전 방탈출' 용 새 서버(새 구글 시트)를 만들어 그 주소를 넣어야 해요.
//    비워 두면(또는 주소 형식이 아니면) '게스트 체험'만 할 수 있어요.
// 2) MAX_FLOOR : 지금 열어 둘 층 수 (처음엔 10). 층을 늘릴 땐 숫자만 바꾸면 돼요. (서버 Code.gs 의 TOP_FLOOR 도 같이 바꿔 주세요)
// 3) LIST_NEEDS_ITEMS : true(기본) = 📋 장치 목록에 '이미 가 본 방의 장치'만 보이고, 🔍 탐지기/🧭 나침반을 얻어야 전체 장치·출구가 보여요.
//    false 로 바꾸면 예전처럼 처음부터 모든 장치와 출구가 목록에 보여요.
// 4) EXTRA_QUESTIONS : 문제를 더 넣고 싶을 때 아래 목록에 추가하세요. (첫 번째 보기가 정답, 보기는 게임에서 자동으로 섞여요)
//    ※ id 는 문제마다 겹치지 않는 고정 이름이에요 (예: "n001"). 나중에 문장을 고쳐도 id 만 같으면 오답노트·통계가 이어져요. 한 번 쓴 id 는 다시 쓰지 마세요.
window.ESC_CONFIG = {
  SERVER_URL: "https://script.google.com/macros/s/AKfycbwsZEQ2orj2BQ7bNZFb30HpXv_svmDIoIxPhPaXiakflmaWwh6U7LKAtEKK6Knp6bY/exec",
  MAX_FLOOR: 30,
  LIST_NEEDS_ITEMS: true,
  EXTRA_QUESTIONS: [
    // { id: "n001", s: "안전", lv: 1, q: "문제 내용", o: ["정답", "오답1", "오답2", "오답3"], e: "해설" },
    //   s(과목): 안전 / 제도 / 요소 / 공유압 / 보전 / 측정      lv(난이도): 1 쉬움 · 2 보통 · 3 어려움
  ]
};
