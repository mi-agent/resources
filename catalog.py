"""全量模型目录：合并 LiteLLM 模型库 + models.dev + OpenRouter 实时接口 + 上游确认免费清单。
免费状态：✅确认免费（上游确认，或 OpenRouter 实时标价为 0）/ 🆓标价为0（某数据源标价为 0）/ ❔需核对（提供商有免费层，但该模型是否在免费层内未确认）。
任何数据源失败都不会中断构建，只会在总览和日志里标注。"""
import json,re,time,datetime,urllib.request
UA="Mozilla/5.0 (compatible; free-resources-auto/1.0)"
LITELLM="https://raw.githubusercontent.com/BerriAI/litellm/main/model_prices_and_context_window.json"
MODELSDEV="https://models.dev/api.json"
OPENROUTER="https://openrouter.ai/api/v1/models"
PMAP={"NVIDIA NIM":{"lite":["nvidia_nim"],"md":["nvidia"]},"Google Gemini":{"lite":["gemini"],"md":["google"]},"Groq":{"lite":["groq"],"md":["groq"]},
"Mistral AI":{"lite":["mistral"],"md":["mistral"]},"Cloudflare Workers AI":{"lite":["cloudflare"],"md":["cloudflare-workers-ai"]},"Cerebras":{"lite":["cerebras"],"md":["cerebras"]},
"Cohere":{"lite":["cohere","cohere_chat"],"md":["cohere"]},"DeepSeek":{"lite":["deepseek"],"md":["deepseek"]},"xAI":{"lite":["xai"],"md":["xai"]},
"SambaNova":{"lite":["sambanova"],"md":["sambanova"]},"Z AI (Zhipu AI)":{"lite":["zai"],"md":["zai","zhipuai"]},"Nebius":{"lite":["nebius"],"md":["nebius"]},
"Nscale":{"lite":["nscale"],"md":["nscale"]},"Hugging Face":{"lite":["huggingface"],"md":["huggingface"]},"OpenRouter":{"lite":["openrouter"],"md":["openrouter"]},
"Alibaba Cloud Model Studio":{"lite":["dashscope"],"md":["alibaba"]},"ModelScope":{"lite":[],"md":["modelscope"]},"SiliconFlow":{"lite":[],"md":["siliconflow"]},
"OVHcloud AI Endpoints":{"lite":["ovhcloud"],"md":["ovhcloud"]},"AI21 Labs":{"lite":["ai21"],"md":["ai21"]},"Ollama Cloud":{"lite":[],"md":["ollama-cloud"]}}
EXTRA={"GitHub Models":{"lite":[],"md":["github-models"],"tier":"永久免费层（补充）","card":"需 GitHub 账号","key":"https://github.com/settings/tokens","base":"https://models.github.ai/inference"}}
MODE={"chat":"对话","completion":"对话","responses":"对话","embedding":"嵌入","image_generation":"生图","image_edit":"生图","audio_transcription":"语音识别","audio_speech":"语音合成","rerank":"重排序","video_generation":"视频生成","ocr":"OCR"}
TODAY=datetime.date.today().isoformat()
def get_json(url):
    err=""
    for _ in range(2):
        try:
            with urllib.request.urlopen(urllib.request.Request(url,headers={"User-Agent":UA}),timeout=90) as r: return json.loads(r.read().decode("utf-8"))
        except Exception as e: err=e;time.sleep(2)
    print(f"[数据源不可用] {url}: {err}");return None
def bkey(s):
    s=str(s).lower().strip()
    if s.endswith(":free"): s=s[:-5]
    return s.rsplit("/",1)[-1]
def ctxn(c):
    m=re.match(r"([\d.]+)\s*([KkMm]?)",str(c or ""))
    return float(m.group(1))*{"":1,"k":1e3,"m":1e6}[m.group(2).lower()] if m else 0
def fmt(n):
    n=n or 0
    return "-" if n<=0 else (f"{n/1e6:g}M" if n>=1e6 else f"{round(n/1e3):g}K" if n>=1e3 else str(int(n)))
def num(x):
    try: return float(x)
    except (TypeError,ValueError): return None
def upsert(recs,prov,name,mid,src,**kw):
    r=recs.setdefault((prov,bkey(mid)),{"prov":prov,"name":name,"id":mid,"mode":"","ctx":0,"maxout":0,"vision":False,"reasoning":False,"tools":False,"json":False,"pdf":False,"audio":False,
        "pin":None,"pout":None,"released":"","dep":False,"confirmed":False,"rate":"","src":set()})
    r["src"].add(src)
    for k,v in kw.items():
        if k in("vision","reasoning","tools","json","pdf","audio","dep","confirmed"): r[k]=r[k] or bool(v)
        elif k in("pin","pout"):
            if r[k] is None and v is not None: r[k]=v
        elif v and not r.get(k): r[k]=v
    return r
def from_litellm(d,inv,recs):
    for key,v in d.items():
        if not isinstance(v,dict): continue
        up=inv.get(v.get("litellm_provider"));mode=MODE.get(v.get("mode"))
        if not up or not mode: continue
        pi,po=num(v.get("input_cost_per_token")),num(v.get("output_cost_per_token"))
        dep=bool(v.get("deprecation_date")) and str(v["deprecation_date"])<=TODAY
        upsert(recs,up,key.rsplit("/",1)[-1],key,"LiteLLM",mode=mode,ctx=v.get("max_input_tokens") or v.get("max_tokens") or 0,maxout=v.get("max_output_tokens") or 0,
            vision=v.get("supports_vision"),reasoning=v.get("supports_reasoning"),tools=v.get("supports_function_calling"),json=v.get("supports_response_schema"),
            pdf=v.get("supports_pdf_input"),audio=v.get("supports_audio_input"),pin=None if pi is None else pi*1e6,pout=None if po is None else po*1e6,dep=dep)
def from_modelsdev(d,inv,recs):
    for pid,pv in d.items():
        up=inv.get(pid)
        if not up or not isinstance(pv,dict): continue
        for mid,m in (pv.get("models") or {}).items():
            mo=m.get("modalities") or {};i=mo.get("input") or [];o=mo.get("output") or [];c=m.get("cost") or {};l=m.get("limit") or {}
            mode="生图" if "image" in o and "text" not in o else "对话"
            upsert(recs,up,m.get("name") or mid,mid,"models.dev",mode=mode,ctx=l.get("context") or 0,maxout=l.get("output") or 0,vision="image" in i,reasoning=m.get("reasoning"),
                tools=m.get("tool_call"),pdf="pdf" in i,audio="audio" in i,pin=num(c.get("input")) if c else None,pout=num(c.get("output")) if c else None,released=m.get("release_date") or "")
def from_openrouter(d,recs):
    for m in (d or {}).get("data",[]):
        pr=m.get("pricing") or {};pi,po=num(pr.get("prompt")),num(pr.get("completion"))
        if pi!=0 or po!=0: continue          # 只收实时标价为 0 的免费模型
        a=m.get("architecture") or {};i=a.get("input_modalities") or [];o=a.get("output_modalities") or [];sp=m.get("supported_parameters") or []
        created=datetime.datetime.fromtimestamp(m["created"],datetime.timezone.utc).strftime("%Y-%m-%d") if m.get("created") else ""
        r=upsert(recs,"OpenRouter",m.get("name") or m.get("id"),m.get("id",""),"OpenRouter",mode="生图" if "image" in o and "text" not in o else "对话",ctx=m.get("context_length") or 0,
            maxout=(m.get("top_provider") or {}).get("max_completion_tokens") or 0,vision="image" in i,reasoning=("reasoning" in sp or "include_reasoning" in sp),tools="tools" in sp,
            json=("structured_outputs" in sp or "response_format" in sp),pdf="file" in i,audio="audio" in i,pin=0.0,pout=0.0,released=created)
        r["confirmed"]=True
def build(provs,M):
    """provs: 上游提供商名集合；M: 上游代表模型 (prov,name,id,ctx,rate)。返回 (记录列表, 数据源状态 dict)"""
    cfg={**PMAP,**EXTRA};inv_l={};inv_m={}
    for up,c in cfg.items():
        if up in provs:
            for x in c.get("lite",[]): inv_l[x]=up
            for x in c.get("md",[]): inv_m[x]=up
    recs={};st={}
    d=get_json(LITELLM);st["LiteLLM"]=bool(d)
    if d: d.pop("sample_spec",None);from_litellm(d,inv_l,recs)
    d=get_json(MODELSDEV);st["models.dev"]=bool(d)
    if d: from_modelsdev(d,inv_m,recs)
    if "OpenRouter" in provs:
        d=get_json(OPENROUTER);st["OpenRouter"]=bool(d)
        if d: from_openrouter(d,recs)
    for prov,name,mid,ctx,rate in M:
        if re.search(r"deprecat|retire",name+mid,re.I): continue   # 上游已标注弃用/即将下线
        r=upsert(recs,prov,name,mid,"上游",ctx=ctxn(ctx))
        r["confirmed"]=True;r["rate"]=rate if rate!="See provider" else "";r["name"]=name
    out=[]
    for r in recs.values():
        if r["dep"]: continue
        r["mode"]=r["mode"] or "对话"
        r["status"]="✅ 确认免费" if r["confirmed"] else "🆓 标价为0" if (r["pin"]==0 and r["pout"]==0 and r["mode"] in ("对话","嵌入")) else "❔ 需核对"
        out.append(r)
    return out,st
