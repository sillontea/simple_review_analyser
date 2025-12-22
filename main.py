import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os

from langchain_ollama import OllamaLLM
from langchain_core.prompts import ChatPromptTemplate

llm = OllamaLLM(model="qwen2.5:1.5b")

# ---------- LLM 함수 (여기에 너의 Ollama/Qwen용 함수 연결) ----------



def classify_sentiment(text:pd.DataFrame):
    """
    단일 문장 감성 분류 함수 (Qwen)
    return: "positive" / "negative" / "neutral"
    """
    prompt = ChatPromptTemplate.from_template("""너는 콜센터 고객 만족도 분석 전문가다.

    아래 문장을 고객의 "만족 여부" 기준으로 분류하라.
    대답은 반드시 '긍정', '부정', '모호함' 중 하나만 출력하라.
    
    분류 기준:
    - '긍정' = 실제로 만족, 칭찬, 감사, 좋은 경험을 표현한 문장
      예) 좋아요, 만족합니다, 친절했어요, 해결됐어요
    
    - '부정' = 불만, 불편, 화남, 문제 발생, 해결 안 됨 등
      예) 불친절했음, 불만족, 짜증남, 해결 안됨
    
    - '모호함' = 명확한 만족/불만 표현이 없는 문장
      예)
        • 요청/희망/건의 ("~해주면 좋겠습니다", "~되면 좋겠어요")
        • 정보 전달/사실 설명만 있는 문장
        • 중립적인 감정 표현
    
    아래 문장을 분류하라:
    
    문장: "{text}"

    답변:
    """)
    
    chain_senti = prompt | llm
    return chain_senti.invoke(text)

def summarize_text(text):
    """
    선택된 감성 문장 전체 요약 (Qwen)
    """
    prompt = ChatPromptTemplate.from_template("""
        너는 콜센터 상담 품질 분석 전문가이다.  
        아래에는 고객 만족도 조사에서 '긍정적인 리뷰'들만 모아놓은 전체 텍스트가 주어진다.  
        
        너의 업무는 긍정 리뷰 전체를 읽고,
        고객들이 어떤 점을 긍정적으로 평가했는지를 **대표적인 5가지 핵심 포인트**로 정리하는 것이다.
        
        요약 시 다음 지침을 반드시 따르라:
        
        1. **부정적인 내용은 절대 포함하지 않는다.**  
        2. **핵심 포인트는 서로 중복되지 않게 명확히 구분한다.**  
        3. **표현은 실무 보고서처럼 간결하고 명확하게 쓴다.**  
        4. **리뷰에서 자주 등장하거나 인상 깊었던 긍정적 요소만 요약한다.**  
        5. **모든 출력은 한국어 bullet point 5개로만 구성한다.**
        
        <고객 리뷰 전체>
        {input}
        
        <최종 출력 형식>
        - (긍정 포인트 1)
        - (긍정 포인트 2)
        - (긍정 포인트 3)
        - (긍정 포인트 4)
        - (긍정 포인트 5)
        """)
        
    chain = prompt | llm
    return chain.invoke(text)


# ----------------------- GUI 시작 -----------------------

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("리뷰 감성 분석기")
app.geometry("900x650")


# 탭 생성
tabview = ctk.CTkTabview(app)
tabview.pack(fill="both", expand=True, padx=10, pady=10)

tab1 = tabview.add("1. 감성분류")
tab2 = tabview.add("2. 요약")


# ----------------------- 탭 1: 감성 분류 -----------------------

file_path_var = ctk.StringVar(value="")

def load_excel_for_classification():
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        file_path_var.set(path)
        log1.insert("end", f"[파일 선택] {path}\n")

def run_sentiment_classification():

    path = file_path_var.get().strip()
    if not path or not os.path.exists(path):
        messagebox.showerror("오류", "엑셀 파일을 먼저 선택하세요.")
        return

    df = pd.read_excel(path)
    if "리뷰" not in df.columns:
        messagebox.showerror("오류", "'리뷰' 라는 열이 엑셀에 있어야 합니다.")
        return

    log1.insert("end", "감성 분류 시작...\n")
    app.update()

    results = []
    total = len(df)
    
    for i, text in enumerate(df["리뷰"]):
        sentiment = classify_sentiment(str(text))
        results.append(sentiment)

        # 진행률 업데이트
        progress1.set((i+1) / total)
        app.update()

        if (i+1) % 50 == 0:
            log1.insert("end", f"{i+1}개 처리...\n")
            app.update()

    df["감성"] = results

    save_path = path.replace(".xlsx", "_감성분류.xlsx")
    df.to_excel(save_path, index=False)

    log1.insert("end", f"[완료] 감성분류 저장됨 → {save_path}\n")
    messagebox.showinfo("완료", "감성 분류가 완료되었습니다.")


# UI 구성 (탭1)
ctk.CTkLabel(tab1, text="리뷰 엑셀 파일 선택").pack(anchor="w", padx=10, pady=(10, 0))

file_frame = ctk.CTkFrame(tab1)
file_frame.pack(fill="x", padx=10)

ctk.CTkEntry(file_frame, textvariable=file_path_var, width=500).pack(side="left", padx=5, pady=10)
ctk.CTkButton(file_frame, text="파일 선택", command=load_excel_for_classification).pack(side="left", padx=5)

ctk.CTkButton(tab1, text="감성 분류 실행", command=run_sentiment_classification).pack(pady=10)

progress1 = ctk.CTkProgressBar(tab1)
progress1.pack(fill="x", padx=10)
progress1.set(0)

log1 = ctk.CTkTextbox(tab1, width=800, height=300)
log1.pack(padx=10, pady=10)


# ----------------------- 탭 2: 요약 -----------------------

summary_file_path_var = ctk.StringVar(value="")
selected_sentiment_var = ctk.StringVar(value="긍정")

def load_excel_for_summary():
    path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if path:
        summary_file_path_var.set(path)
        log2.insert("end", f"[파일 선택] {path}\n")

def filter_sentences():
    """선택 감성의 문장만 불러오기"""
    path = summary_file_path_var.get()
    if not path or not os.path.exists(path):
        messagebox.showerror("오류", "감성분류된 엑셀을 선택하세요.")
        return

    df = pd.read_excel(path)

    if "감성" not in df.columns or "리뷰" not in df.columns:
        messagebox.showerror("오류", "'리뷰'와 '감성' 열이 있어야 합니다.")
        return

    sentiment = selected_sentiment_var.get()
    filtered = df[df["감성"].str.contains(sentiment)]["리뷰"].tolist()

    text = "\n".join(filtered)
    review_textbox.delete("0.0", "end")
    review_textbox.insert("end", text)

    log2.insert("end", f"{sentiment} 문장 {len(filtered)}개 불러옴.\n")


def run_summary():

    text = review_textbox.get("0.0", "end").strip()

    if not text:
        messagebox.showerror("오류", "요약할 문장이 없습니다.")
        return

    log2.insert("end", "요약 시작...\n")
    app.update()

    result = summarize_text(text)

    summary_textbox.delete("0.0", "end")
    summary_textbox.insert("end", result)

    log2.insert("end", "요약 완료.\n")
    messagebox.showinfo("완료", "요약이 완료되었습니다.")


# UI 구성 (탭2)
ctk.CTkLabel(tab2, text="감성분류된 엑셀 선택").pack(anchor="w", padx=10, pady=(10, 0))

file_frame2 = ctk.CTkFrame(tab2)
file_frame2.pack(fill="x", padx=10)

ctk.CTkEntry(file_frame2, textvariable=summary_file_path_var, width=500).pack(side="left", padx=5, pady=10)
ctk.CTkButton(file_frame2, text="파일 선택", command=load_excel_for_summary).pack(side="left", padx=5)


# 감성 선택 라디오 버튼
select_frame = ctk.CTkFrame(tab2)
select_frame.pack(anchor="w", padx=10, pady=10)

ctk.CTkLabel(select_frame, text="요약할 감성 선택:").pack(side="left", padx=5)
ctk.CTkRadioButton(select_frame, text="긍정", value="긍정", variable=selected_sentiment_var).pack(side="left")
ctk.CTkRadioButton(select_frame, text="모호", value="모호", variable=selected_sentiment_var).pack(side="left")
ctk.CTkRadioButton(select_frame, text="부정", value="부정", variable=selected_sentiment_var).pack(side="left")

ctk.CTkButton(tab2, text="문장 불러오기", command=filter_sentences).pack(pady=5)

# 요약 실행 & 저장
def save_summary_to_txt():
    result = summary_textbox.get("0.0", "end").strip()

    if not result:
        messagebox.showerror("오류", "저장할 요약 결과가 없습니다.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".txt",
        filetypes=[("Text files", "*.txt")],
        title="요약 결과 저장"
    )

    if save_path:
        with open(save_path, "w", encoding="utf-8") as f:
            f.write(result)
        messagebox.showinfo("저장 완료", f"요약 결과가 저장되었습니다.\n{save_path}")


# 텍스트 박스
review_textbox = ctk.CTkTextbox(tab2, width=800, height=200)
review_textbox.pack(padx=10, pady=10)

# --- 요약 실행 + 요약 저장 버튼 Frame (텍스트박스 바로 아래) ---
summary_button_frame = ctk.CTkFrame(tab2)
summary_button_frame.pack(pady=5)

ctk.CTkButton(summary_button_frame, text="요약 실행", command=run_summary).pack(side="left", padx=5)
ctk.CTkButton(summary_button_frame, text="요약 결과 저장", command=save_summary_to_txt).pack(side="left", padx=5)

# 텍스트 박스
summary_textbox = ctk.CTkTextbox(tab2, width=800, height=180)
summary_textbox.pack(padx=10, pady=10)

log2 = ctk.CTkTextbox(tab2, width=800, height=120)
log2.pack(padx=10, pady=10)




# -------------------------------------------------------

app.mainloop()