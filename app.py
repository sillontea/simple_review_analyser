import customtkinter as ctk
from extract import main

# -----------------------------
# GUI 설정
# -----------------------------
ctk.set_appearance_mode("light")   # "dark" 가능
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("키워드 추출 프로그램")
app.geometry("150X200")

# 실행 함수
def run_main():
    try:
        main()  # ← main() 함수 실행
        result_label.configure(text="✅ 완료되었습니다!")
    except Exception as e:
        result_label.configure(text=f"⚠️ 오류 발생: {e}")

# ✅ 버튼 크기 및 글자 크기 조정
run_button = ctk.CTkButton(
    app,
    text="실행",
    width=200,          # 기본 150 → 200
    height=60,          # 기본 40 → 60
    font=ctk.CTkFont(size=30, weight="bold"),  # 버튼 내 텍스트 크기
    command=run_main
)
run_button.pack(pady=40)


# ✅ 안내 문구
guide_text = (
    "1. input_text 폴더 안에 리뷰 파일(엑셀)을 준비하세요.\n"
    "2. 실행버튼을 누르세요.\n"
    "3. 현재 위치에 'result.csv' 파일이 생겼는지 확인하세요.\n"
    "4. 작업이 다 끝나면 창을 닫으세요."
)
guide_label = ctk.CTkLabel(app, text=guide_text, justify="left", anchor="w")
guide_label.pack(pady=(15, 10), padx=20)

# 상태 표시 레이블
result_label = ctk.CTkLabel(app, text="")
result_label.pack(pady=20)

app.mainloop()