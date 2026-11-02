#!/usr/bin/env python3
"""自动获取并更新免费资源 Excel。
数据源：open-free-llm-api/awesome-freellm-apis 的 README（上游每日自动刷新）→ 解析 → 与本地精选资源合并 →
存活探测 → 与上次快照对比生成变更日志 → 输出 free-resources.xlsx / providers.csv / README.md。
用法：python build.py [--no-probe]"""
import json,re,ssl,sys,csv,os,math,time,datetime,urllib.request,urllib.error
import catalog
from concurrent.futures import ThreadPoolExecutor
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule,DataBarRule
from openpyxl.utils import get_column_letter
UP="https://raw.githubusercontent.com/open-free-llm-api/awesome-freellm-apis/main/README.md"
UA="Mozilla/5.0 (compatible; free-resources-auto/1.0)"
FAIL_LIMIT=3
NOW=datetime.datetime.now(datetime.timezone.utc)
TODAY=NOW.strftime("%Y-%m-%d");STAMP=NOW.strftime("%Y-%m-%d %H:%M UTC")
NOPROBE="--no-probe" in sys.argv
XLSX="free-resources.xlsx"

def fetch(url):
    for i in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(url,headers={"User-Agent":UA}),timeout=30) as r: return r.read().decode("utf-8")
        except Exception as e:
            err=e;time.sleep(3)
    sys.exit(f"获取上游失败：{err}")
def tables(md):
    cur="";out={}
    for ln in md.splitlines():
        if ln.startswith("#"): cur=ln.lstrip("# ").strip();continue
        if ln.startswith("|"):
            cells=[c.strip() for c in ln.strip().strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?",c) for c in cells): continue
            out.setdefault(cur,[]).append(cells)
    return out
def url_of(c):
    m=re.search(r'href="(https?://[^"]+)"',c) or re.search(r"\((https?://[^)\s]+)\)",c);return m.group(1) if m else ""
def txt(c):
    c=re.sub(r"<[^>]+>","",c)   # 上游原始 README 用 HTML 锚点，也兼容 Markdown 链接
    return re.sub(r"\[([^\]]*)\]\([^)]*\)",r"\1",c).replace("`","").strip()
def num(c): return int(re.sub(r"\D","",c) or 0)
CARD={"No":"无需信用卡","Registration":"需注册","Phone verification":"需手机验证"}
def parse(md):
    T=tables(md);P={};M=[]
    def find(k):
        for h,v in T.items():
            if k in h: return v[1:]
        return []
    for key,tier in (("Permanent","永久免费层"),("Renewable","可更新额度")):
        for r in find(key):
            if len(r)>=6:
                n=txt(r[0]);c=txt(r[2])
                P[n]={"name":n,"tier":tier,"models":num(r[1]),"card":CARD.get(c,c),"ctx":txt(r[3]),"mod":txt(r[4]),"key":url_of(r[5]),"base":"","top":[]}
    for r in find("Quick Reference"):
        n=txt(r[0])
        if n in P:
            P[n]["base"]=txt(r[1]) if txt(r[1])!="-" else ""
            P[n]["key"]=P[n]["key"] or url_of(r[2])
    prov=""
    for r in find("Best Free Models"):
        if len(r)<5: continue
        if txt(r[0]): prov=txt(r[0])
        if prov in P: P[prov]["top"].append((txt(r[1]),txt(r[2]),txt(r[3]),txt(r[4])));M.append((prov,txt(r[1]),txt(r[2]),txt(r[3]),txt(r[4])))
    m=re.search(r"Last updated:\s*(\d{4}-\d{2}-\d{2})",md)
    return P,M,(m.group(1) if m else "未知")
def probe(url):
    if NOPROBE or not url: return None,"未探测"
    u=re.sub(r"\{[^}]*\}","x",url);last=""
    for _ in range(2):
        try:
            with urllib.request.urlopen(urllib.request.Request(u,headers={"User-Agent":UA}),timeout=15,context=ssl.create_default_context()) as r: return True,r.status
        except urllib.error.HTTPError as e:
            if e.code<500: return True,e.code   # 任何 <500 的 HTTP 响应都说明主机还活着
            last=f"HTTP {e.code}"
        except Exception as e: last=type(e).__name__
        time.sleep(2)
    return False,last
def judge(S,name,res,log):
    s=S.setdefault(name,{"fails":0,"last_ok":None,"first_seen":TODAY,"status":"未探测"})
    ok,_=res;old=s.get("status")
    if ok is None: st="未探测"
    else:
        s["fails"]=0 if ok else s["fails"]+1
        if ok: s["last_ok"]=STAMP
        st="有效" if s["fails"]==0 else ("观察中" if s["fails"]<FAIL_LIMIT else "已移除")
    if st=="已移除" and old!="已移除": log.append([TODAY,"探测失败移除",name,f"连续 {s['fails']} 次探测失败"])
    if old=="已移除" and st in("有效","观察中"): log.append([TODAY,"恢复",name,"探测恢复正常"])
    s["status"]=st;return st,(s["last_ok"] or "-")

# ---------- 能力 / 场景分类（按提供商模态 + 模型名启发式推断，仅供参考） ----------
CN_PROV={"ModelScope","SiliconFlow","Z AI (Zhipu AI)","Alibaba Cloud Model Studio","DeepSeek"}
FAM=[("llama","Llama"),("qwen","Qwen"),("deepseek","DeepSeek"),("glm","GLM"),("kimi|moonshot","Kimi"),("gemini","Gemini"),("gemma","Gemma"),
("mistral|mixtral|codestral|devstral|magistral","Mistral"),("nemotron","Nemotron"),("gpt-oss","GPT-OSS"),("minimax","MiniMax"),("grok","Grok"),
("command","Command"),("jamba","Jamba"),("phi","Phi"),("laguna","Laguna"),("step","StepFun"),("mimo","MiMo"),("aion","Aion"),("agnes","Agnes"),("granite","Granite"),("big-pickle","Big Pickle")]
CNFAM={"Qwen","DeepSeek","GLM","Kimi","MiniMax","MiMo","StepFun"}
def ctxn(c):
    m=re.match(r"([\d.]+)\s*([KkMm]?)",c or "")
    if not m: return 0
    return float(m.group(1))*{"":1,"k":1e3,"m":1e6}[m.group(2).lower()]
def classify(name,mid,ctx):
    t=(name+" "+mid).lower();tags=set()
    if re.search(r"code|coder|codestral|devstral|laguna|swe|kimi-k2-7-code",t): tags.add("code")
    if re.search(r"r1|reason|think|qwq|magistral",t): tags.add("reason")
    if re.search(r"\bvl\b|-vl|vl-|vision|\dv-|\d\.\dv|pixtral|vlm|omni|multimodal",t): tags.add("vision")
    if re.search(r"embed",t): tags.add("embed")
    if "rerank" in t: tags.add("rerank")
    if re.search(r"whisper|audio|speech|tts|asr|voice",t): tags.add("audio")
    if re.search(r"flux|diffusion|sdxl|imagen|image-\d|-image|text-to-image",t): tags.add("imggen")
    fam=next((v for k,v in FAM if re.search(k,t)),"其他")
    size="未标注";m=re.findall(r"(?<![a-z])(\d+(?:\.\d+)?)b(?![a-z])",t)
    if m:
        b=float(m[0]);size="小（≤10B）" if b<=10 else "中（11–40B）" if b<=40 else "大（41–150B）" if b<=150 else "超大（>150B）"
    elif re.search(r"(?<![a-z])(mini|lite|flash|small|nano|tiny|lightning|haiku|fast)(?![a-z])",t): size="轻量（按名称）"
    elif re.search(r"(?<![a-z])(ultra|max|pro|large|plus|opus|medium)(?![a-z])",t): size="旗舰（按名称）"
    if size.startswith(("小","轻量")): tags.add("light")
    if size.startswith(("超大","旗舰")): tags.add("big")
    if fam in CNFAM: tags.add("cn")
    if ctxn(ctx)>=500000: tags.add("long")
    return fam,size,tags
TAGTXT=[("code","💻编程"),("reason","🧠推理"),("vision","👁视觉"),("long","📄长文档"),("light","⚡轻量"),("big","🏆旗舰"),("cn","🇨🇳中文友好"),("audio","🎙音频"),("imggen","🖼生图"),("embed","🔎嵌入"),("rerank","🔎重排")]
SCEN=[("💻 编程助手","写代码、补全、Agent 编程","code",None),("🧠 复杂推理 / 数学","多步推理、数学、分析","reason",None),
("🏆 高质量通用","追求回答质量，不太在意速度","big",None),("📄 超长上下文（≥500K）","长文档、整仓代码","long",None),
("👁 看图 / 多模态","图片理解、截图问答","vision",None),("⚡ 轻量快速 / 低延迟","简单对话、批量处理","light",None),
("🇨🇳 中文场景","中文写作与问答","cn",None),("🔎 向量检索 / RAG","嵌入与重排序（按提供商）",None,"embedding|rerank"),
("🎙 语音 / 音频","语音识别与合成（按提供商）",None,"audio|speech"),("🖼 图像生成","含 image 模态的提供商（可能是图像输入或输出，请到官方核对）",None,"image")]
CAPS=[("文本","text"),("推理","reasoning"),("代码","code"),("视觉","vision"),("图像","image"),("音频/语音","audio|speech"),("视频","video"),("嵌入","embedding"),("重排序","rerank"),("PDF","pdf")]
CAT={"免费服务器":"2F5BEA","免费GPU/算力":"7A4DE0","免费数据库/存储":"0E9F8E","免费域名/DNS/证书":"2E9F55","免费路由":"E0772B","免费VPN":"C93B6B","免费AI工具":"B7791F","免费图像/语音":"B03AB0"}
def old_user():
    out={}
    if not os.path.exists(XLSX): return out
    try:
        wb=load_workbook(XLSX)
        for sh,key in (("免费API","提供商"),("其他免费资源","名称")):
            if sh not in wb.sheetnames: continue
            ws=wb[sh];h=[c.value for c in ws[3]]
            if key in h and "我的状态" in h:
                k,a,b=h.index(key),h.index("我的状态"),h.index("我的备注")
                for row in ws.iter_rows(min_row=4,values_only=True): out[(sh,row[k])]=(row[a] or "未注册",row[b] or "")
    except Exception as e: print("读取旧表失败，用户标记将重置：",e)
    return out
# ---------- Excel 样式 ----------
F="Microsoft YaHei";NAVY="14213D";ZEBRA="F6F8FC"
def ft(**k): return Font(name=F,**k)
thin=Side(style="thin",color="E1E7F0")
CH={"永久免费层":("16895C","E6F6EF"),"可更新额度":("2F5BEA","E8EEFD"),"无需信用卡":("16895C","E6F6EF"),"需注册":("B7791F","FDF3E0"),"需手机验证":("C2410C","FDEBDD"),
"长期":("16895C","E6F6EF"),"有限制":("B7791F","FDF3E0"),"试用":("5B6781","EDF0F6"),"有效":("16895C","E6F6EF"),"观察中":("B7791F","FDF3E0"),"未探测":("5B6781","EDF0F6"),
"🇨🇳 国内":("C93B6B","FCE8EF"),"🌐 海外":("2F5BEA","E8EEFD"),"●":("16895C","D5F0E3")}
def chip(v): return CH.get(v) or (("C2410C","FDEBDD") if isinstance(v,str) and v.startswith("Free") else None)
def mk(wb,title,banner,sub,headers,widths,rows,tab,links=(),chips=(),status=None,mine=None,bars=(),cat=None,center=()):
    ws=wb.create_sheet(title);n=len(headers);L=get_column_letter(n);ws.sheet_properties.tabColor=tab;ws.sheet_view.showGridLines=False;ws.sheet_view.zoomScale=95
    ws.merge_cells(f"A1:{L}1");ws["A1"]=banner;ws["A1"].font=ft(size=16,bold=True,color="FFFFFF");ws["A1"].alignment=Alignment(vertical="center",indent=1)
    for i in range(1,n+1): ws.cell(1,i).fill=PatternFill("solid",fgColor=NAVY)
    ws.row_dimensions[1].height=36
    ws.merge_cells(f"A2:{L}2");ws["A2"]=sub;ws["A2"].font=ft(size=9,italic=True,color="5B6781");ws["A2"].alignment=Alignment(vertical="center",indent=1,wrap_text=True)
    for i in range(1,n+1): ws.cell(2,i).fill=PatternFill("solid",fgColor="EEF2F7")
    ws.row_dimensions[2].height=24 if len(sub)<110 else 42
    for i,(h,w) in enumerate(zip(headers,widths),1):
        c=ws.cell(3,i,h);c.font=ft(bold=True,color="FFFFFF",size=10);c.fill=PatternFill("solid",fgColor="2B3A67");c.alignment=Alignment(vertical="center",horizontal="center" if i in center else "left",indent=1)
        c.border=Border(bottom=Side(style="medium",color=tab));ws.column_dimensions[get_column_letter(i)].width=w
    ws.row_dimensions[3].height=28
    for r,row in enumerate(rows,4):
        lines=1
        for i,v in enumerate(row,1):
            c=ws.cell(r,i,v);c.font=ft(size=10);c.alignment=Alignment(wrap_text=True,vertical="center",horizontal="center" if i in center else "left",indent=1);c.border=Border(bottom=thin)
            if r%2==1: c.fill=PatternFill("solid",fgColor=ZEBRA)
            if i in links and v:
                c.value="打开 ↗";c.hyperlink=v;c.font=ft(size=10,color="2F5BEA",underline="single")
            elif isinstance(v,str):
                lines=max(lines,math.ceil(len(v)*1.7/widths[i-1]))
                if i in chips and chip(v): fg,bg=chip(v);c.font=ft(size=10,bold=True,color=fg);c.fill=PatternFill("solid",fgColor=bg);c.alignment=Alignment(vertical="center",horizontal="center",wrap_text=True)
        if cat and row[cat-1] in CAT:
            c=ws.cell(r,cat);c.font=ft(size=10,bold=True,color=CAT[row[cat-1]]);c.border=Border(left=Side(style="thick",color=CAT[row[cat-1]]),bottom=thin)
        ws.row_dimensions[r].height=min(130,max(26,lines*15+8))
    last=max(len(rows)+3,4);ws.freeze_panes="B4";ws.auto_filter.ref=f"A3:{L}{last}"
    rule=lambda col,txt,fg,bg:ws.conditional_formatting.add(f"{col}4:{col}1000",FormulaRule(formula=[f'${col}4="{txt}"'],font=Font(color=fg,bold=True),fill=PatternFill("solid",bgColor=bg)))
    if mine:
        dv=DataValidation(type="list",formula1='"未注册,已注册,在用,已失效"',allow_blank=True);ws.add_data_validation(dv);dv.add(f"{mine}4:{mine}1000")
        rule(mine,"在用","16895C","D5F0E3");rule(mine,"已失效","B42318","FBE4E4");rule(mine,"已注册","2F5BEA","E8EEFD")
    for b in bars: ws.conditional_formatting.add(f"{b}4:{b}{last}",DataBarRule(start_type="num",start_value=0,end_type="max",color="9DB7F5",showValue=True))
    ws.page_setup.orientation="landscape"
    return ws
def main():
    md=fetch(UP);P,M,upd=parse(md)
    if len(P)<10: sys.exit("上游格式可能已变化：解析到的提供商过少，已中止以免覆盖现有数据")
    recs,srcs=catalog.build(set(P)|set(catalog.EXTRA),M)
    for n,x in catalog.EXTRA.items():
        k=sum(1 for r in recs if r["prov"]==n)
        if k: P[n]={"name":n,"tier":x["tier"],"models":k,"card":x["card"],"ctx":"","mod":"text","key":x["key"],"base":x["base"],"top":[]}
    try: S=json.load(open("state.json",encoding="utf-8"))
    except FileNotFoundError: S={}
    S.setdefault("api",{});S.setdefault("other",{});S.setdefault("gone",{});S.setdefault("log",[])
    log=[];first=not S["api"]
    for n,p in P.items():
        if n in S["gone"]: del S["gone"][n];log.append([TODAY,"回归",n,"重新出现在上游列表"])
        elif not first and n not in S["api"]: log.append([TODAY,"新增提供商",n,f"{p['models']} 个免费模型，{p['card']}"])
        elif n in S["api"] and S["api"][n].get("models") not in (None,p["models"]): log.append([TODAY,"模型数变化",n,f"{S['api'][n]['models']} → {p['models']}"])
    for n in list(S["api"]):
        if n not in P: S["gone"][n]={"date":TODAY,"reason":"上游已移除"};log.append([TODAY,"下架提供商",n,"不再出现在上游列表"]);del S["api"][n]
    R=json.load(open("resources.json",encoding="utf-8"))
    jobs=[p["base"] or p["key"] for p in P.values()]+[r["probe"] for r in R]
    with ThreadPoolExecutor(8) as ex: res=list(ex.map(probe,jobs))
    ra,rb=res[:len(P)],res[len(P):]
    old=old_user();api=[];dead=[];other=[];mat=[]
    def capline(mod): return " ".join(e for e,k in (("💬","text"),("🧠","reasoning"),("💻","code"),("👁","vision"),("🖼","image"),("🎙","audio|speech"),("🎬","video"),("🔎","embedding|rerank"),("📄","pdf")) if re.search(k,mod))
    for p,rs in zip(P.values(),ra):
        st,lo=judge(S["api"],p["name"],rs,log);S["api"][p["name"]]["models"]=p["models"]
        us=old.get(("免费API",p["name"]),("未注册",""))
        top="；".join(f"{a}" for a,b,_,_ in p["top"][:3]);rate=next((t[3] for t in p["top"] if t[3] and t[3]!="See provider"),"")
        reg="🇨🇳 国内" if p["name"] in CN_PROV else "🌐 海外"
        row=[p["name"],reg,p["tier"],p["models"],p["card"],p["ctx"],capline(p["mod"]),p["base"] or "-",p["key"],top,rate,st,lo,*us]
        if st=="已移除": dead.append([p["name"],"探测连续失败",TODAY]);continue
        api.append(row);mat.append([p["name"],*["●" if re.search(k,p["mod"]) else "" for _,k in CAPS],p["models"],p["card"]])
    for r,rs in zip(R,rb):
        st,lo=judge(S["other"],r["name"],rs,log);us=old.get(("其他免费资源",r["name"]),("未注册",""))
        if st=="已移除": dead.append([r["name"],"探测连续失败",TODAY]);continue
        other.append([r["category"],r["name"],r["quota"],r["term"],r["fit"],r["url"],r["steps"],st,lo,*us])
    api.sort(key=lambda x:-x[3]);mat.sort(key=lambda x:-x[-2]);dead+=[[n,g["reason"],g["date"]] for n,g in S["gone"].items()]
    other.sort(key=lambda x:list(CAT).index(x[0]) if x[0] in CAT else 99)
    S["log"]=(log+S["log"])[:300];S["last_run"]=STAMP
    try: age=(NOW.date()-datetime.date.fromisoformat(upd)).days
    except ValueError: age=0
    if age>7 and not any(l[1]=="上游陈旧" and l[0]==TODAY for l in S["log"]): S["log"].insert(0,[TODAY,"上游陈旧","awesome-freellm-apis",f"README 标注更新日期 {upd}，已超过 7 天"])
    json.dump(S,open("state.json","w",encoding="utf-8"),ensure_ascii=False,indent=1)
    live={r[0]:r for r in api};mods=[];SO={"✅ 确认免费":0,"🆓 标价为0":1,"❔ 需核对":2}
    for r in recs:
        if r["prov"] not in live: continue
        fam,size,tags=classify(r["name"],r["id"],catalog.fmt(r["ctx"]))
        if r["vision"]: tags.add("vision")
        if r["reasoning"]: tags.add("reason")
        if r["ctx"]>=500000: tags.add("long")
        pr=live[r["prov"]];mods.append({**r,"fam":fam,"size":size,"tags":tags,"card":pr[4],"key":pr[8],"n":r["ctx"]})
    mods.sort(key=lambda m:(SO[m["status"]],m["prov"],-m["n"]))
    ids=sorted(f"{m['prov']}|{catalog.bkey(m['id'])}" for m in mods if m["status"]!="❔ 需核对")
    if all(srcs.values()):
        prev=S.get("free_ids")
        if prev is not None:
            new,gone=set(ids)-set(prev),set(prev)-set(ids)
            if new: S["log"].insert(0,[TODAY,"新增免费模型",f"{len(new)} 个","、".join(sorted(new))[:200]])
            if gone: S["log"].insert(0,[TODAY,"免费模型消失",f"{len(gone)} 个","、".join(sorted(gone))[:200]])
        S["free_ids"]=ids
    else: S["log"].insert(0,[TODAY,"数据源不可用","、".join(k for k,v in srcs.items() if not v),"本次未做模型级对比，保留上次快照"])
    json.dump(S,open("state.json","w",encoding="utf-8"),ensure_ascii=False,indent=1)
    wb=Workbook();ov=wb.active;ov.title="总览"
    # 场景推荐
    CH2={"✅ 确认免费":("16895C","E6F6EF"),"🆓 标价为0":("0E9F8E","DDF5F1"),"❔ 需核对":("B7791F","FDF3E0"),"✔":("16895C","D5F0E3")};CH.update(CH2)
    sc=[];chatm=[m for m in mods if m["mode"]=="对话"]
    SC2=[("💻 编程助手","写代码、补全、Agent 编程",lambda m:"code" in m["tags"],"对话"),("🧠 复杂推理 / 数学","支持推理的模型",lambda m:"reason" in m["tags"],"对话"),
    ("🏆 高质量通用","旗舰或大参数模型",lambda m:"big" in m["tags"],"对话"),("📄 超长上下文（≥500K）","长文档、整仓代码",lambda m:"long" in m["tags"],"对话"),
    ("👁 看图 / 多模态","图片、截图、PDF 理解",lambda m:"vision" in m["tags"],"对话"),("🛠 工具调用 / Agent","支持函数调用",lambda m:m["tools"],"对话"),
    ("⚡ 轻量快速 / 低延迟","简单对话、批量处理",lambda m:"light" in m["tags"],"对话"),("🇨🇳 中文场景","中文写作与问答",lambda m:"cn" in m["tags"],"对话"),
    ("🔎 向量检索 / RAG","嵌入与重排序模型",lambda m:True,"嵌入|重排序"),("🎙 语音识别 / 合成","ASR 与 TTS 模型",lambda m:True,"语音识别|语音合成"),("🖼 图像生成","文生图模型",lambda m:True,"生图")]
    for title,desc,fn,mode in SC2:
        cand=sorted([m for m in mods if re.fullmatch(mode,m["mode"]) and fn(m)],key=lambda m:(SO[m["status"]],m["card"]!="无需信用卡",-m["n"]));picks=[];cnt={}
        for m in cand:
            if cnt.get(m["prov"],0)>=2: continue
            cnt[m["prov"]]=cnt.get(m["prov"],0)+1;picks.append([title,desc,m["name"],m["prov"],m["status"],catalog.fmt(m["ctx"]),m["rate"] or "见官方",m["card"],m["key"]])
            if len(picks)>=6: break
        good=[p for p in picks if p[4]!="❔ 需核对"]
        if len(good)>=4: picks=good
        sc+=picks or [[title,desc,"（当前数据源中暂无匹配）","","","","","",""]]
    mk(wb,"场景推荐","🎯 按应用场景选模型","能力来自 LiteLLM / models.dev / OpenRouter 的结构化字段，编程、中文、轻量等标签按模型名推断。排序：确认免费 > 标价为 0 > 需核对，再看是否免信用卡与上下文；同一提供商最多 2 个。",
       ["场景","适用说明","推荐模型","提供商","免费状态","上下文","限速","信用卡/验证","获取 Key"],[22,24,38,20,13,10,22,14,10],sc,"E0772B",links=(9,),chips=(5,8),center=(6,))
    mk(wb,"免费API","⚡ 免费 LLM API 提供商","数据来源：awesome-freellm-apis（freellm.net 每日刷新）。“我的状态/我的备注”可自行填写，重新生成时按名称保留。",
       ["提供商","地区","类型","免费模型数","信用卡/验证","最大上下文","能力","Base URL","获取 Key","代表模型","限速参考","状态","最近正常","我的状态","我的备注"],[20,10,12,12,13,10,16,36,10,40,22,9,19,11,22],api,"2F5BEA",links=(9,),chips=(2,3,5,12),mine="N",bars=("D",),center=(2,3,5,6,12))
    ws=mk(wb,"能力矩阵","🧩 提供商 × 能力矩阵","● 表示该提供商的免费模型中包含此类能力（来自上游模态字段）。",
       ["提供商",*[c for c,_ in CAPS],"免费模型数","信用卡/验证"],[22]+[9]*len(CAPS)+[12,14],mat,"7A4DE0",chips=(len(CAPS)+3,),bars=(get_column_letter(len(CAPS)+2),),center=tuple(range(2,len(CAPS)+3)))
    for row in ws.iter_rows(min_row=4,max_row=3+len(mat),min_col=2,max_col=len(CAPS)+1):
        for c in row:
            c.alignment=Alignment(horizontal="center",vertical="center")
            if c.value=="●": c.font=ft(size=12,bold=True,color="16895C");c.fill=PatternFill("solid",fgColor="D5F0E3")
    T=lambda b:"✔" if b else ""
    pr=lambda x:None if x is None else round(x,4)
    mk(wb,"全量模型库","📚 全量模型库（免费层提供商的模型 × 能力）",f"数据源：{'；'.join(k+('✓' if v else '✗') for k,v in srcs.items())}，另含上游确认清单。免费状态：✅确认免费 / 🆓标价为0 / ❔该提供商有免费层，但此模型是否在免费层内需到官网核对。价格单位：美元 / 百万 token。已弃用的模型已自动剔除。",
       ["提供商","模型","免费状态","类型","上下文","最大输出","视觉","推理","工具","JSON","PDF","音频输入","输入价","输出价","系列","规模","标签","来源","发布"],[18,38,13,10,9,9,7,7,7,7,7,9,9,9,11,17,32,20,11],
       [[m["prov"],m["name"],m["status"],m["mode"],catalog.fmt(m["ctx"]),catalog.fmt(m["maxout"]),T(m["vision"]),T(m["reasoning"]),T(m["tools"]),T(m["json"]),T(m["pdf"]),T(m["audio"]),pr(m["pin"]),pr(m["pout"]),m["fam"],m["size"],
         "  ".join(t for k,t in TAGTXT if k in m["tags"]),"+".join(sorted(m["src"])),m["released"]] for m in mods],"0E9F8E",chips=(3,7,8,9,10,11,12),center=(3,4,5,6,7,8,9,10,11,12,13,14))
    mk(wb,"其他免费资源","🧰 服务器 · 算力 · 数据库 · 域名 · 路由 · VPN · AI 工具","手工整理的补充资源，额度与政策变化快，以官网为准；状态来自每日存活探测。",
       ["类别","名称","免费额度","时效","适合","官网","使用方法","状态","最近正常","我的状态","我的备注"],[16,26,44,9,24,10,70,9,19,11,22],other,"C93B6B",links=(6,),chips=(4,8),mine="J",cat=1,center=(4,8))
    mk(wb,"变更日志","📝 变更日志","新增/下架/模型数变化/探测失败与恢复，最近 300 条。",["日期","类型","对象","详情"],[12,16,28,60],S["log"],"B7791F")
    mk(wb,"已下架","🗑 已下架","上游移除，或连续 3 次探测失败的条目。恢复后会自动回到有效表。",["名称","原因","日期"],[30,24,12],dead,"5B6781")
    # ---- 总览仪表盘 ----
    ov.sheet_view.showGridLines=False;ov.sheet_properties.tabColor=NAVY
    for r in range(1,40):
        for c in range(1,11): ov.cell(r,c).fill=PatternFill("solid",fgColor="EEF2F7")
    ov.column_dimensions["A"].width=3;ov.column_dimensions["J"].width=3
    for c in "BCDEFGHI": ov.column_dimensions[c].width=15
    ov.merge_cells("B2:I2");ov["B2"]="免费资源中心";ov["B2"].font=ft(size=26,bold=True,color=NAVY);ov.row_dimensions[2].height=44
    ov.merge_cells("B3:I3");ov["B3"]=f"生成于 {STAMP}　·　上游数据日期 {upd}　·　数据源：上游"+"".join(f" · {k}{'✓' if v else '✗'}" for k,v in srcs.items());ov["B3"].font=ft(size=10,color="5B6781")
    if age>7: ov.merge_cells("B4:I4");ov["B4"]=f"⚠ 上游数据已 {age} 天未更新，内容可能不再准确";ov["B4"].font=ft(size=10,bold=True,color="B42318")
    cards=[("有效 API 提供商","=COUNTA('免费API'!A4:A1000)","2F5BEA","个"),("确认免费的模型",'=COUNTIF(\'全量模型库\'!C4:C5000,"✅ 确认免费")',"7A4DE0","个"),
           ("无需信用卡的提供商",'=COUNTIF(\'免费API\'!E4:E1000,"无需信用卡")',"16895C","个"),("国内服务商",'=COUNTIF(\'免费API\'!B4:B1000,"🇨🇳 国内")',"C93B6B","个"),
           ("其他免费资源","=COUNTA('其他免费资源'!B4:B1000)","E0772B","条"),("全量模型库条目","=COUNTA('全量模型库'!A4:A5000)","0E9F8E","条"),
           ("已下架",'=COUNTA(\'已下架\'!A4:A1000)',"5B6781","条"),("我标记为在用",'=COUNTIF(\'免费API\'!N4:N1000,"在用")+COUNTIF(\'其他免费资源\'!J4:J1000,"在用")',"B7791F","个")]
    for i,(lab,f,col,unit) in enumerate(cards):
        r0=6+(i//4)*5;c0=2+(i%4)*2
        for rr in range(r0,r0+4):
            ov.merge_cells(start_row=rr,start_column=c0,end_row=rr,end_column=c0+1)
            for cc in (c0,c0+1):
                cell=ov.cell(rr,cc);cell.fill=PatternFill("solid",fgColor="FFFFFF")
                if rr==r0: cell.border=Border(top=Side(style="thick",color=col))
        ov.cell(r0+1,c0,lab).font=ft(size=10,color="5B6781");ov.cell(r0+1,c0).alignment=Alignment(horizontal="left",indent=1)
        v=ov.cell(r0+2,c0,f);v.font=ft(size=28,bold=True,color=col);v.alignment=Alignment(horizontal="left",indent=1,vertical="center");ov.row_dimensions[r0+2].height=42
        ov.cell(r0+3,c0,unit).font=ft(size=9,color="9AA7C4");ov.cell(r0+3,c0).alignment=Alignment(horizontal="left",indent=1)
    ov.merge_cells("B17:I17");ov["B17"]="📚 工作表导航（点击跳转）";ov["B17"].font=ft(size=13,bold=True,color=NAVY)
    nav=[("场景推荐","按编程、推理、看图、长文档、工具调用、中文、RAG、语音、生图等场景挑模型"),("免费API","30 家提供商的额度、信用卡要求、Base URL 与 Key 入口"),("能力矩阵","一眼看各提供商支持哪些能力"),("全量模型库","每个模型的类型、上下文、视觉/推理/工具等能力、价格与免费状态"),("其他免费资源","服务器、GPU、数据库、域名、路由、VPN、AI 工具"),("变更日志","每天新增、下架与变动"),("已下架","失效条目归档")]
    for i,(sh,d) in enumerate(nav,18):
        ov.merge_cells(start_row=i,start_column=2,end_row=i,end_column=3);ov.merge_cells(start_row=i,start_column=4,end_row=i,end_column=9)
        ov.cell(i,2,f'=HYPERLINK("#\'{sh}\'!A1","→ {sh}")').font=ft(size=11,bold=True,color="2F5BEA",underline="single");ov.cell(i,4,d).font=ft(size=10,color="5B6781")
        for cc in range(2,10): ov.cell(i,cc).fill=PatternFill("solid",fgColor="FFFFFF");ov.cell(i,cc).border=Border(bottom=thin)
    for i,t in enumerate(["探测只能确认接口/站点还活着，不能确认免费额度未变；注册前请以官网为准。","场景、系列、规模标签由程序按名称与模态推断，属于参考分类。","第三方项目请先读源码或脚本再运行；使用 VPN 或跨境工具请遵守所在地法律法规。"],26):
        ov.merge_cells(start_row=i,start_column=2,end_row=i,end_column=9);ov.cell(i,2,"• "+t).font=ft(size=9,color="5B6781")
    wb.save(XLSX)
    with open("providers.csv","w",encoding="utf-8-sig",newline="") as f:
        w=csv.writer(f);w.writerow(["provider","region","tier","free_models","card","max_context","capabilities","base_url","key_url","top_models","rate_limit","status","last_ok"])
        w.writerows(r[:13] for r in api)
    L=[f"# 免费 LLM API 一览（自动更新）\n\n生成：{STAMP}　上游数据：{upd}　有效提供商 {len(api)} 个，免费模型 {sum(r[3] for r in api)} 个\n","| 提供商 | 地区 | 类型 | 免费模型 | 信用卡/验证 | 获取 Key |","|---|---|---|---|---|---|"]
    L+=[f"| {r[0]} | {r[1]} | {r[2]} | {r[3]} | {r[4]} | "+(f"[获取]({r[8]})" if r[8] else "-")+" |" for r in api]
    L+=["","## 最近变更",""]+[f"- {a} **{b}** {c}：{d}" for a,b,c,d in S["log"][:15]]
    open("README.md","w",encoding="utf-8").write("\n".join(L)+"\n")
    print(f"有效 {len(api)}，下架 {len(dead)}，其他资源 {len(other)}，模型 {len(mods)}，变更 {len(log)}")
if __name__=="__main__": main()
