"""
Teramath - PDF 시험지 → HWP 네이티브 수식 변환기
한글(HWP) 자동화를 사용해 Ctrl+N+M 수식 객체로 직접 삽입
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, json, os, sys, subprocess, base64, re

# ── 패키지 자동 설치 ─────────────────────────────────────
def install_if_needed():
    needed = []
    try: import anthropic
    except: needed.append('anthropic')
    try: import win32com.client
    except: needed.append('pywin32')
    if needed:
        root = tk.Tk(); root.withdraw()
        messagebox.showinfo("Setup", f"Installing: {', '.join(needed)}\nPlease wait...")
        for p in needed:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', p, '-q'])
        messagebox.showinfo("Done", "Complete! Please restart the program.")
        sys.exit(0)

install_if_needed()
import anthropic
import win32com.client as win32

# ════════════════════════════════════════════════════════
#  LaTeX → HWP 수식 편집기 문법 변환기
# ════════════════════════════════════════════════════════
class Latex2HWP:
    """LaTeX를 HWP 수식 편집기 문법으로 변환"""
    
    GREEK = {
        'alpha':'alpha','beta':'beta','gamma':'gamma','Gamma':'GAMMA',
        'delta':'delta','Delta':'DELTA','epsilon':'epsilon','varepsilon':'epsilon',
        'zeta':'zeta','eta':'eta','theta':'theta','Theta':'THETA',
        'iota':'iota','kappa':'kappa','lambda':'lambda','Lambda':'LAMBDA',
        'mu':'mu','nu':'nu','xi':'xi','Xi':'XI',
        'pi':'pi','Pi':'PI','rho':'rho','sigma':'sigma','Sigma':'SIGMA',
        'tau':'tau','upsilon':'upsilon','phi':'phi','Phi':'PHI',
        'chi':'chi','psi':'psi','Psi':'PSI','omega':'omega','Omega':'OMEGA',
    }
    FUNCS = {
        'sin':'sin','cos':'cos','tan':'tan','cot':'cot','sec':'sec','csc':'csc',
        'log':'log','ln':'ln','exp':'exp','max':'max','min':'min',
        'det':'det','deg':'deg','gcd':'gcd','lcm':'lcm',
        'arcsin':'arcsin','arccos':'arccos','arctan':'arctan',
        'sinh':'sinh','cosh':'cosh','tanh':'tanh',
        'lim':'lim','sup':'sup','inf':'inf',
    }
    SYMBOLS = {
        'leq':'<=','le':'<=','geq':'>=','ge':'>=','neq':'<>','ne':'<>',
        'approx':'~~','equiv':'=def=','sim':'~','simeq':'~=',
        'times':'times','div':'div','pm':'+-','mp':'-+','cdot':'.',
        'infty':'inf','partial':'partial','nabla':'nabla',
        'to':'->','rightarrow':'->','leftarrow':'<-',
        'Rightarrow':'=>','Leftarrow':'<=','Leftrightarrow':'<=>','leftrightarrow':'<->',
        'therefore':'therefore','because':'because',
        'in':'in','notin':'notin','subset':'subset','supset':'supset',
        'subseteq':'subeq','supseteq':'supeq',
        'cup':'union','cap':'inter','emptyset':'empty','varnothing':'empty',
        'perp':'perp','parallel':'para','angle':'angle',
        'int':'int','iint':'iint','iiint':'iiint','oint':'oint',
        'sum':'sum','prod':'prod',
        'forall':'forall','exists':'exists',
        'ldots':'...','cdots':'cdots','vdots':'vdots','ddots':'ddots',
    }

    def __init__(self, s):
        self.s = s.strip()
        self.i = 0

    def convert(self):
        result = self._seq()
        return re.sub(r'  +', ' ', result).strip()

    def _seq(self):
        parts = []
        while self.i < len(self.s):
            ch = self.s[self.i]
            if ch == '}':
                break
            elif ch == '\\':
                parts.append(self._cmd())
            elif ch == '{':
                self.i += 1
                inner = self._seq()
                if self.i < len(self.s) and self.s[self.i] == '}':
                    self.i += 1
                parts.append('{' + inner + '}')
            elif ch == '^':
                self.i += 1
                base = parts.pop() if parts else ' '
                exp = self._arg()
                parts.append(f'{base}^{{{exp}}}')
            elif ch == '_':
                self.i += 1
                base = parts.pop() if parts else ' '
                sub = self._arg()
                # Check if next is ^
                if self.i < len(self.s) and self.s[self.i] == '^':
                    self.i += 1
                    sup = self._arg()
                    parts.append(f'{base}_{{{sub}}}^{{{sup}}}')
                else:
                    parts.append(f'{base}_{{{sub}}}')
            elif ch in ' \t\n':
                self.i += 1
                if parts and not parts[-1].endswith(' '):
                    parts.append(' ')
            else:
                parts.append(ch)
                self.i += 1
        return ''.join(parts)

    def _arg(self):
        while self.i < len(self.s) and self.s[self.i] == ' ':
            self.i += 1
        if self.i >= len(self.s):
            return ''
        if self.s[self.i] == '{':
            self.i += 1
            inner = self._seq()
            if self.i < len(self.s) and self.s[self.i] == '}':
                self.i += 1
            return inner
        if self.s[self.i] == '\\':
            return self._cmd()
        ch = self.s[self.i]
        self.i += 1
        return ch

    def _cmd(self):
        self.i += 1  # skip backslash
        name = ''
        if self.i < len(self.s) and self.s[self.i].isalpha():
            while self.i < len(self.s) and self.s[self.i].isalpha():
                name += self.s[self.i]; self.i += 1
            while self.i < len(self.s) and self.s[self.i] == ' ':
                self.i += 1
        elif self.i < len(self.s):
            name = self.s[self.i]; self.i += 1

        # 분수
        if name == 'frac':
            num = self._arg()
            den = self._arg()
            return f'{{{num}}} over {{{den}}}'

        # 제곱근
        if name == 'sqrt':
            deg = None
            if self.i < len(self.s) and self.s[self.i] == '[':
                self.i += 1
                deg = ''
                while self.i < len(self.s) and self.s[self.i] != ']':
                    deg += self.s[self.i]; self.i += 1
                self.i += 1
            rad = self._arg()
            if deg:
                return f'nroot{{{deg}}}{{{rad}}}'
            return f'sqrt{{{rad}}}'

        # 적분 (하한/상한 처리)
        if name == 'int':
            return self._with_limits('int')
        if name == 'iint':
            return self._with_limits('iint')
        if name == 'oint':
            return self._with_limits('oint')

        # 합/곱 (하한/상한 처리)
        if name == 'sum':
            return self._with_limits('sum')
        if name == 'prod':
            return self._with_limits('prod')

        # 극한
        if name == 'lim':
            result = 'lim'
            if self.i < len(self.s) and self.s[self.i] == '_':
                self.i += 1
                sub = self._arg()
                sub = sub.replace('\\to', '->').replace('to', '->')
                result += f' from{{{sub}}}'
            return result

        # 그리스 문자
        if name in self.GREEK:
            return self.GREEK[name]

        # 함수 이름
        if name in self.FUNCS:
            return self.FUNCS[name]

        # 연산자/기호
        if name in self.SYMBOLS:
            return self.SYMBOLS[name]

        # 괄호
        if name == 'left':
            if self.i < len(self.s):
                br = self.s[self.i]
                if br == '(':   self.i += 1; return 'left ('
                elif br == '[': self.i += 1; return 'left ['
                elif br == '|': self.i += 1; return 'left |'
                elif br == '\\' and self.i+1 < len(self.s) and self.s[self.i+1] == '{':
                    self.i += 2; return 'left {'
                elif br == '.': self.i += 1; return ''
            return 'left ('

        if name == 'right':
            if self.i < len(self.s):
                br = self.s[self.i]
                if br == ')':   self.i += 1; return 'right )'
                elif br == ']': self.i += 1; return 'right ]'
                elif br == '|': self.i += 1; return 'right |'
                elif br == '\\' and self.i+1 < len(self.s) and self.s[self.i+1] == '}':
                    self.i += 2; return 'right }'
                elif br == '.': self.i += 1; return ''
            return 'right )'

        # 중괄호
        if name == '{': return '{'
        if name == '}': return '}'
        if name == '|': return '|'

        # 간격 (무시)
        if name in (',', ';', '!', 'quad', 'qquad', ' ', ','): return ' '

        # 텍스트 명령 (내용만 반환)
        if name in ('text','mathrm','mathbf','mathit','mathbb','mbox',
                    'overline','underline','hat','bar','vec','tilde',
                    'widehat','widetilde','overbrace','underbrace'):
            return self._arg()

        # 기타
        if name == 'not':
            inner = self._arg()
            return 'not ' + inner

        if name == '\\': return ' '
        if name == ',': return ''

        # 알 수 없는 명령어 → 그대로
        return name

    def _with_limits(self, op):
        """하한/상한이 있는 연산자 처리 (int, sum 등)"""
        result = op
        from_part = to_part = None
        while self.i < len(self.s) and self.s[self.i] in ('_', '^', ' '):
            if self.s[self.i] == ' ':
                self.i += 1; continue
            if self.s[self.i] == '_':
                self.i += 1; from_part = self._arg()
            elif self.s[self.i] == '^':
                self.i += 1; to_part = self._arg()
        if from_part: result += f' from{{{from_part}}}'
        if to_part:   result += f' to{{{to_part}}}'
        return result


def latex_to_hwp(latex):
    """LaTeX → HWP 수식 편집기 문법"""
    try:
        return Latex2HWP(latex).convert()
    except Exception:
        return latex


# ════════════════════════════════════════════════════════
#  HWP COM 자동화
# ════════════════════════════════════════════════════════
class HWPController:
    def __init__(self):
        self.hwp = None

    def start(self):
        try:
            self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        except Exception:
            self.hwp = win32.Dispatch("HWPFrame.HwpObject")
        try:
            self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass
        self.hwp.HAction.Run("FileNew")
        return self

    def insert_text(self, text):
        if not text: return
        act = self.hwp.HAction
        pset = self.hwp.HParameterSet
        act.GetDefault("InsertText", pset.HInsertText.HSet)
        pset.HInsertText.Text = str(text)
        act.Execute("InsertText", pset.HInsertText.HSet)

    def insert_equation(self, eq_hwp):
        """HWP 수식 편집기 문법으로 수식 삽입 (HAction 직접 호출)"""
        eq_hwp = str(eq_hwp).strip()
        if not eq_hwp:
            return False
        try:
            act  = self.hwp.HAction
            pset = self.hwp.HParameterSet
            act.GetDefault("InsertEquation", pset.HEquation.HSet)
            pset.HEquation.FormText = eq_hwp
            pset.HEquation.BaseLineY = 0
            act.Execute("InsertEquation", pset.HEquation.HSet)
            return True
        except Exception as e:
            # 텍스트 fallback
            self.insert_text(f" {eq_hwp} ")
            return False

    def set_font_size(self, pt):
        act = self.hwp.HAction
        pset = self.hwp.HParameterSet
        act.GetDefault("CharShape", pset.HCharShape.HSet)
        pset.HCharShape.Height = pt * 100
        act.Execute("CharShape", pset.HCharShape.HSet)

    def insert_para(self):
        self.hwp.HAction.Run("BreakPara")

    def insert_line(self, text=''):
        if text:
            self.insert_text(text)
        self.insert_para()

    def save_as_hwp(self, path):
        # HWP는 백슬래시 경로 필요
        path = path.replace("/", "\\")
        self.hwp.SaveAs(path, "HWP", "")

    def quit(self):
        try:
            self.hwp.Quit()
        except Exception:
            pass


def insert_mixed_content(hwpc, text):
    """$...$가 섞인 텍스트를 텍스트+수식으로 삽입"""
    parts = re.split(r'(\$[^$]+\$)', text)
    for part in parts:
        if part.startswith('$') and part.endswith('$') and len(part) > 2:
            latex = part[1:-1]
            hwp_eq = latex_to_hwp(latex)
            hwpc.insert_equation(hwp_eq)
        elif part:
            hwpc.insert_text(part)


def create_hwp_document(data, output_path, progress_cb=None):
    """추출된 데이터로 HWP 파일 생성"""
    def p(msg):
        if progress_cb: progress_cb(msg)

    p("HWP 실행 중...")
    hwpc = HWPController()
    hwpc.start()

    try:
        # 제목
        title = data.get('title') or (data.get('subject', '수학') + ' 시험지')
        meta = ' | '.join(filter(None, [data.get('subject'), data.get('grade')]))

        p("제목 삽입 중...")
        hwpc.set_font_size(16)
        hwpc.insert_line(title)
        if meta:
            hwpc.set_font_size(11)
            hwpc.insert_line(meta)
        hwpc.insert_line()

        total = len(data.get('questions', []))
        for idx, q in enumerate(data.get('questions', []), 1):
            p(f"문제 삽입 중... ({idx}/{total})")
            hwpc.set_font_size(11)

            score_str = f" [{q['score']}점]" if q.get('score') else ''
            hwpc.insert_text(f"{q['number']}.{score_str}  ")

            # 문제 내용 (텍스트 + 수식 혼합)
            insert_mixed_content(hwpc, q.get('content', ''))
            hwpc.insert_para()

            # 그림 표시
            if q.get('has_figure'):
                hwpc.insert_line('   ※ [그림 삽입 필요]')

            # 선지
            if q.get('options'):
                for opt in q['options']:
                    hwpc.insert_text('   ')
                    insert_mixed_content(hwpc, opt)
                    hwpc.insert_para()

            hwpc.insert_line()

        p("HWP 파일 저장 중...")
        hwpc.save_as_hwp(output_path)

    finally:
        try:
            hwpc.quit()
        except Exception:
            pass


# ════════════════════════════════════════════════════════
#  Claude API 호출
# ════════════════════════════════════════════════════════
PROMPT = """당신은 수학 시험지에서 문제를 추출하는 전문가입니다.
수학 시험지 PDF를 분석하여 반드시 아래 JSON 형식으로만 응답하세요.
마크다운 코드블록 없이 순수 JSON만 출력하세요.

[수식 표기 규칙]
수식 부분은 달러 기호로 감싸고 LaTeX 형식으로 작성하세요.
일반 한글/숫자는 감싸지 마세요.

예시:
- "함수 $f(x)=x^3-3x+5$ 가 감소하는 구간이 $[a,b]$ 일 때"
- "정적분 $\\int_{-1}^{1}(5x^4+8)\\,dx$ 의 값은?"
- "① $-\\frac{2}{3}$"  (분수인 경우에만 달러)
- "남학생 5명과 여학생 4명 중에서 3명을 뽑을 때"  (순수 숫자는 달러 안 씀)

JSON:
{"title":"시험 제목 또는 null","subject":"과목명","grade":"학년 또는 null","questions":[{"number":1,"content":"문제 내용","options":["① $-\\frac{2}{3}$","② $-\\frac{1}{3}$"],"has_figure":false,"score":null}]}

규칙: 모든 문제 빠짐없이 / 그래프·도형 있으면 has_figure:true + [그림] / 객관식 아니면 options:null"""


def convert_pdf(pdf_path, api_key, progress_cb, done_cb, error_cb):
    def run():
        try:
            progress_cb("PDF 파일 읽는 중...")
            with open(pdf_path, 'rb') as f:
                pdf_data = base64.standard_b64encode(f.read()).decode('utf-8')

            progress_cb("AI가 시험지 분석 중... (30~60초 소요)")
            client = anthropic.Anthropic(api_key=api_key)
            msg = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=8000,
                messages=[{"role":"user","content":[
                    {"type":"document","source":{"type":"base64","media_type":"application/pdf","data":pdf_data}},
                    {"type":"text","text":PROMPT}
                ]}]
            )

            raw = msg.content[0].text.strip().replace('```json','').replace('```','').strip()
            data = json.loads(raw)

            # HWP 파일 경로
            base = os.path.splitext(pdf_path)[0]
            out_path = base + '_변환.hwp'

            create_hwp_document(data, out_path, progress_cb)
            done_cb(out_path, len(data.get('questions', [])))

        except Exception as e:
            error_cb(str(e))

    threading.Thread(target=run, daemon=True).start()


# ════════════════════════════════════════════════════════
#  GUI
# ════════════════════════════════════════════════════════
KEY_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.teramath_key')

def load_key():
    try:
        if os.path.exists(KEY_FILE):
            with open(KEY_FILE) as f: return f.read().strip()
    except: pass
    return ''

def save_key_file(k):
    try:
        with open(KEY_FILE, 'w') as f: f.write(k)
    except: pass


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Teramath - HWP 시험지 변환기")
        self.geometry("560x500")
        self.resizable(False, False)
        self.configure(bg='#F4F3EF')
        self.result_path = None
        self._build()

    def _build(self):
        # 헤더
        hdr = tk.Frame(self, bg='#1B4FD8', pady=16)
        hdr.pack(fill='x')
        tk.Label(hdr, text="Teramath  HWP 시험지 변환기",
                 bg='#1B4FD8', fg='white', font=('맑은 고딕',15,'bold')).pack()
        tk.Label(hdr, text="PDF 시험지 → HWP 네이티브 수식 (.hwp)",
                 bg='#1B4FD8', fg='#B5D4F4', font=('맑은 고딕',9)).pack(pady=(2,0))

        body = tk.Frame(self, bg='#F4F3EF', padx=24, pady=18)
        body.pack(fill='both', expand=True)

        # API 키
        tk.Label(body, text="Anthropic API 키", bg='#F4F3EF',
                 font=('맑은 고딕',10,'bold')).pack(anchor='w')
        kf = tk.Frame(body, bg='#F4F3EF')
        kf.pack(fill='x', pady=(4,10))
        self.key_var = tk.StringVar(value=load_key())
        tk.Entry(kf, textvariable=self.key_var, show='*',
                 font=('Consolas',10), relief='flat', bg='white', bd=1
                 ).pack(side='left', fill='x', expand=True, ipady=6, ipadx=6)
        tk.Button(kf, text="저장", command=self._save_key,
                  bg='#1B4FD8', fg='white', font=('맑은 고딕',9),
                  relief='flat', padx=12).pack(side='left', padx=(6,0))
        self.kst = tk.Label(body, text="", bg='#F4F3EF', font=('맑은 고딕',9))
        self.kst.pack(anchor='w')
        self._update_kst()

        # PDF 선택
        tk.Label(body, text="PDF 파일", bg='#F4F3EF',
                 font=('맑은 고딕',10,'bold')).pack(anchor='w', pady=(12,0))
        pf = tk.Frame(body, bg='#F4F3EF')
        pf.pack(fill='x', pady=(4,12))
        self.pdf_var = tk.StringVar()
        tk.Entry(pf, textvariable=self.pdf_var,
                 font=('맑은 고딕',9), relief='flat', bg='white', bd=1,
                 state='readonly').pack(side='left', fill='x', expand=True, ipady=6, ipadx=6)
        tk.Button(pf, text="찾아보기", command=self._browse,
                  bg='#555', fg='white', font=('맑은 고딕',9),
                  relief='flat', padx=10).pack(side='left', padx=(6,0))

        # 변환 버튼
        self.conv_btn = tk.Button(body, text="AI 변환 시작 (HWP 수식으로 저장)",
                                   command=self._convert,
                                   bg='#1B4FD8', fg='white',
                                   font=('맑은 고딕',12,'bold'),
                                   relief='flat', pady=12, state='disabled')
        self.conv_btn.pack(fill='x', pady=(4,0))

        # 상태
        self.status_var = tk.StringVar(value="PDF 파일을 선택하세요")
        tk.Label(body, textvariable=self.status_var, bg='#F4F3EF',
                 font=('맑은 고딕',9), wraplength=500, justify='left'
                 ).pack(anchor='w', pady=(10,4))

        self.progress = ttk.Progressbar(body, mode='indeterminate')
        self.progress.pack(fill='x')

        # 결과
        self.res_frame = tk.Frame(body, bg='#EEF2FD', pady=2)
        self.res_lbl = tk.Label(self.res_frame, text="", bg='#EEF2FD',
                                 font=('맑은 고딕',10), wraplength=490,
                                 justify='left', padx=12, pady=8)
        self.res_lbl.pack(fill='x')
        self.open_btn = tk.Button(self.res_frame, text="HWP 파일 열기",
                                   command=self._open,
                                   bg='#16A34A', fg='white',
                                   font=('맑은 고딕',10,'bold'),
                                   relief='flat', pady=8)
        self.open_btn.pack(fill='x', padx=12, pady=(0,10))

        tk.Label(body, text="console.anthropic.com 에서 API 키 발급 · 시험지 1장 약 5~15원",
                 bg='#F4F3EF', fg='#9ca3af', font=('맑은 고딕',8)
                 ).pack(pady=(14,0))

    def _save_key(self):
        k = self.key_var.get().strip()
        if not k.startswith('sk-ant'):
            messagebox.showerror("오류", "올바른 Anthropic API 키를 입력해주세요\n(sk-ant-... 로 시작)")
            return
        save_key_file(k)
        self._update_kst()
        messagebox.showinfo("저장", "API 키가 저장됐어요!")
        self._check_ready()

    def _update_kst(self):
        k = self.key_var.get().strip()
        if k.startswith('sk-ant'):
            self.kst.config(text="✓ 키 저장됨", fg='#16A34A')
        else:
            self.kst.config(text="키를 입력하고 저장해주세요", fg='#7A7870')

    def _browse(self):
        p = filedialog.askopenfilename(
            title="시험지 PDF 선택",
            filetypes=[("PDF", "*.pdf"), ("모든 파일", "*.*")])
        if p:
            self.pdf_var.set(p)
            self.status_var.set(f"선택: {os.path.basename(p)}")
            self._check_ready()

    def _check_ready(self):
        ok = self.key_var.get().strip().startswith('sk-ant') and self.pdf_var.get().strip()
        self.conv_btn.config(state='normal' if ok else 'disabled')

    def _convert(self):
        k = self.key_var.get().strip()
        p = self.pdf_var.get().strip()
        if not k.startswith('sk-ant'):
            messagebox.showerror("오류", "API 키를 먼저 저장해주세요.")
            return
        if not p or not os.path.exists(p):
            messagebox.showerror("오류", "PDF 파일을 선택해주세요.")
            return
        self.conv_btn.config(state='disabled')
        self.res_frame.pack_forget()
        self.progress.start(10)

        convert_pdf(p, k,
            progress_cb=lambda m: self.after(0, lambda: self.status_var.set(m)),
            done_cb=lambda path, cnt: self.after(0, lambda: self._done(path, cnt)),
            error_cb=lambda err: self.after(0, lambda: self._error(err)))

    def _done(self, path, cnt):
        self.progress.stop()
        self.result_path = path
        self.res_lbl.config(text=f"변환 완료!  {cnt}개 문제  →  {os.path.basename(path)}")
        self.res_frame.pack(fill='x', pady=(10,0))
        self.status_var.set("완료! HWP 파일에 수식이 삽입됐어요.")
        self.conv_btn.config(state='normal')

    def _error(self, err):
        self.progress.stop()
        self.status_var.set(f"오류: {err[:80]}")
        messagebox.showerror("오류", f"{err}\n\n확인사항:\n• API 키 확인\n• 인터넷 연결 확인\n• HWP 설치 확인")
        self.conv_btn.config(state='normal')

    def _open(self):
        if self.result_path and os.path.exists(self.result_path):
            os.startfile(self.result_path)


if __name__ == '__main__':
    app = App()
    app.mainloop()
