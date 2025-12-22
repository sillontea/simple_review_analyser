# simple_review_analyser
A tool that extracts keywords from reviews using TF-IDF and clusters sentences by related concepts.


## 1. 키워드 추출기
How to use?
    "1. input_text 폴더 안에 리뷰 파일(엑셀)을 준비하세요.\n"
    "2. 실행버튼을 누르세요.\n"
    "3. 현재 위치에 'result.csv' 파일이 생겼는지 확인하세요.\n"
    "4. 작업이 다 끝나면 창을 닫으세요."

결과는 TF-IDF 점수가 높은 순으로 10개가 추출되며, result.csv로 저장됩니다.


## 2. 문장 분류기
1) https://ollama.com/download 에서 Ollama 설치

2) 설치 후 터미널에서:
   ollama pull qwen2.5:1.5b
 * 터미널: 윈도우 키+r 입력후 cmd 실행

3) 이후 main.exe 실행
