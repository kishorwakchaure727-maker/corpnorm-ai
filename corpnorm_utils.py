import re
import requests
import json
import time
import difflib
import openai
from urllib.parse import urlparse, unquote

# =========================================================
# 1. Normalization Constants & Logic
# =========================================================

LEGAL_SUFFIXES = [
    " CO LTD", " CO", " LTD", " LLC", " LLP", " INC", " CORP", " CORPORATION",
    " SDN BHD", " PVT LTD", " PVT", " BV", " GMBH", " AG", " SA", " SRL",
    " SAS", " HOLDING AG", " PLANT", " FACTORY", " BRANCH", " S P A", " AB", " OY",
    " K K", " G K", " SPOL S R O"
]

def strict_normalize_name(text: str) -> str:
    if not isinstance(text, str) or not text:
        return ""
    text = text.replace("&", " AND ")
    text = re.sub(r"[^A-Za-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = text.strip().upper()
    changed = True
    while changed and text:
        changed = False
        for suffix in LEGAL_SUFFIXES:
            if text.endswith(suffix):
                text = text[: -len(suffix)].rstrip()
                changed = True
    return text

# =========================================================
# 2. URL Validation & Classification
# =========================================================

BLOCKED_OFFICIAL = [
    "linkedin.com", "facebook.com", "instagram.com", "twitter.com", "x.com",
    "youtube.com", "wikipedia.org", "glassdoor.com", "indeed.com",
    "yellowpages", "bloomberg.com", "dnb.com", "kompass.com", "zoominfo",
    "tmall.com", "taobao.com", "1688.com", "godaddy.com", "namecheap.com"
]

THIRD_PARTY_DOMAINS = [
    "dnb.com", "pitchbook.com", "bloomberg.com", "taiwantrade.com",
    "opencorporates.com", "lei-lookup.com", "reuters.com", "zoominfo.com",
    "kompass.com", "techcrunch.com"
]

def clean_url(url: str) -> str:
    if not isinstance(url, str): return ""
    u = url.strip()
    if not u or " " in u or "." not in u: return ""
    if not re.match(r"^https?://", u): u = "https://" + u
    return u

def get_domain(url: str) -> str:
    try: return urlparse(url).netloc.lower()
    except Exception: return ""

# =========================================================
# 3. Search Providers (DDG & SerpAPI)
# =========================================================

def duckduckgo_search_api(query: str) -> list:
    url = "https://api.duckduckgo.com/"
    params = {"q": query, "format": "json", "no_html": 1, "skip_disambig": 1}
    headers = {"User-Agent": "Mozilla/5.0"}
    urls = []
    try:
        resp = requests.get(url, params=params, headers=headers, timeout=10)
        data = resp.json()
        if data.get("AbstractURL"): urls.append(data["AbstractURL"])
        for r in data.get("Results", []):
            if r.get("FirstURL"): urls.append(r["FirstURL"])
        for topic in data.get("RelatedTopics", []):
            if isinstance(topic, dict):
                if topic.get("FirstURL"): urls.append(topic["FirstURL"])
    except Exception: pass
    return urls

def serpapi_search(query: str, api_key: str) -> list:
    """
    Simulates Bing result structure but using SerpAPI (Google).
    Returns list of dicts: [{"name": Title, "url": URL, "snippet": Snippet}]
    """
    url = "https://serpapi.com/search.json"
    params = {
        "engine": "google",
        "q": query,
        "api_key": api_key,
        "num": 5
    }
    try:
        resp = requests.get(url, params=params, timeout=10)
        data = resp.json()
        
        if "error" in data:
            return {"error": data["error"]}
            
        results = []
        # Organic Results
        for r in data.get("organic_results", []):
            results.append({
                "name": r.get("title", ""),
                "url": r.get("link", ""),
                "snippet": r.get("snippet", "")
            })
        return {"webPages": {"value": results}} # Simulate Bing structure for compat
    except Exception as e:
        return {"error": str(e)}

# =========================================================
# 4. Agentic Verification Logic
# =========================================================

def fuzzy_match_score(str1: str, str2: str) -> float:
    return difflib.SequenceMatcher(None, str1.lower(), str2.lower()).ratio()

def match_domain_score(url: str, normalized_name: str) -> float:
    domain = get_domain(url)
    core_domain = re.sub(r"^www\.", "", domain)
    core_domain = core_domain.split(".")[0]
    flat_name = normalized_name.replace(" ", "").lower()
    
    if flat_name == core_domain: return 1.0 
    if flat_name in core_domain: return 0.9 
    if core_domain in flat_name and len(core_domain) > 3: return 0.8
    
    name_parts = normalized_name.split()
    if len(name_parts) > 0:
        first_word = name_parts[0].lower()
        if len(first_word) > 3 and first_word in core_domain: return 0.7
    return 0.0

def fetch_page_metadata(url: str) -> dict:
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
        resp = requests.get(url, headers=headers, timeout=5, verify=False)
        if resp.status_code != 200: return {"error": f"Status {resp.status_code}"}
        
        content = resp.text
        title_match = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE | re.DOTALL)
        title = title_match.group(1).strip() if title_match else ""
        desc_match = re.search(r'<meta\s+name=["\']description["\']\s+content=["\'](.*?)["\']', content, re.IGNORECASE)
        description = desc_match.group(1).strip() if desc_match else ""
        h1_match = re.search(r'<h1[^>]*>(.*?)</h1>', content, re.IGNORECASE | re.DOTALL)
        h1 = h1_match.group(1).strip() if h1_match else ""
        h1 = re.sub(r'<[^>]+>', '', h1)
        body_text = re.sub(r'<script.*?</script>', ' ', content, flags=re.DOTALL | re.IGNORECASE)
        body_text = re.sub(r'<style.*?</style>', ' ', body_text, flags=re.DOTALL | re.IGNORECASE)
        body_text = re.sub(r'<[^>]+>', ' ', body_text)
        body_text = re.sub(r'\s+', ' ', body_text).strip()
        
        return {"title": title, "description": description, "h1": h1, "body": body_text[:1000]}
    except Exception as e: return {"error": str(e)}

# =========================================================
# 5. Industry Inference
# =========================================================

INDUSTRY_KEYWORDS = [
    ("semiconductor", "Semiconductors"), ("pcb", "PCB manufacturing"), ("connector", "Connectors"),
    ("cable", "Cables & Wires"), ("wire", "Cables & Wires"), ("electronic component", "Electronic components"),
    ("fiber", "Optical fiber solutions"), ("fastener", "Fasteners manufacturing"), ("magnet", "Magnetic components"),
    ("cooling", "Cooling solutions"), ("motor", "Electric motors"), ("automation", "Industrial automation"),
    ("machinery", "Industrial machinery"), ("logistic", "Logistics services"), ("software", "Software solutions"),
    ("consult", "Business consulting"), ("automotive", "Automotive parts"), ("plastic", "Plastic manufacturing"),
    ("rubber", "Rubber products"), ("display", "Display modules"), ("sensor", "Sensors"), ("led", "LED lighting"),
    ("manufactur", "Manufacturing (General)"), ("supplier of", "Industrial Supply"), ("distributor", "Distribution"),
    ("production", "Manufacturing (General)"), ("assembly", "Manufacturing (General)"), ("technology", "Technology"),
    ("solution", "Technology Solutions")
]

def infer_industry_from_text(text: str) -> str:
    if not text: return ""
    text = text.lower()
    for kw, label in INDUSTRY_KEYWORDS:
        if kw in text: return label
    return "" 

# =========================================================
# 6. The AGENT
# =========================================================

class CompanyAgent:
    def verify_candidate(self, url: str, normalized_name: str) -> dict:
        url = clean_url(url)
        if not url: return {"score": 0, "industry": "", "reason": "Bad URL"}

        domain = get_domain(url)
        if any(b in domain for b in BLOCKED_OFFICIAL): return {"score": 0.1, "industry": "", "reason": "Blocked"}
        if any(tp in domain for tp in THIRD_PARTY_DOMAINS): return {"score": 0.2, "industry": "", "reason": "Third Party"}

        domain_score = match_domain_score(url, normalized_name)
        meta = fetch_page_metadata(url)
        title = meta.get("title", "")
        
        blob_lower = (title + " " + meta.get("description", "")).lower()
        parked_keywords = ["domain for sale", "buy this domain", "godaddy", "namecheap", "parked"]
        if any(pk in blob_lower for pk in parked_keywords): return {"score": 0, "industry": "", "reason": "Parked"}

        title_score = 0
        if "error" not in meta:
             clean_title = re.sub(r"(?i)\s*[-|]\s*(home|official|welcome|index).*", "", title)
             title_score = fuzzy_match_score(normalized_name, clean_title)

        final_score = max(domain_score, title_score)
        if domain_score > 0.7 and "error" not in meta: final_score = max(final_score, 0.9)
        elif domain_score > 0.4 and title_score > 0.4: final_score = max(final_score, 0.8)

        blob = f"{title} {meta.get('description', '')} {meta.get('h1', '')} {meta.get('body', '')}"
        industry = infer_industry_from_text(blob)
        if not industry and final_score > 0.6: industry = "Unclassified (Website Found)"
        
        return {"score": final_score, "industry": industry, "reason": f"S:{final_score:.2f} (D:{domain_score:.1f})"}

    def ask_openai(self, raw_name: str, address: dict, search_results: dict, rules: str, openai_key: str) -> dict:
        """Calls OpenAI for smart enrichment."""
        openai.api_key = openai_key
        # Note: We replaced 'bing_results' with 'search_results' to be generic
        messages = [
            {"role": "system", "content": rules},
            {
                "role": "user",
                "content": json.dumps({
                    "raw_name": raw_name,
                    "address": address,
                    "search_results": search_results
                })
            }
        ]
        try:
            response = openai.ChatCompletion.create(
                model="gpt-4o-mini",
                messages=messages,
                temperature=0
            )
            data = json.loads(response["choices"][0]["message"]["content"])
            data.setdefault("website", "")
            data.setdefault("industry", "")
            data.setdefault("remark", "Verified by OpenAI")
            return {
                "Normalized Company Name": strict_normalize_name(data.get("normalized_name", raw_name)),
                "Website": data["website"],
                "Industry": data["industry"],
                "Third Party Data Source Link": data.get("third_party_link", ""),
                "Remark": data["remark"],
                "Confidence Score": "High (AI)"
            }
        except Exception as e:
            # Fallback output
            return {
                "Normalized Company Name": strict_normalize_name(raw_name),
                "Website": "", "Industry": "", "Third Party Data Source Link": "",
                "Remark": f"AI Error: {str(e)}", "Confidence Score": "0" 
            }

    def process_premium(self, raw_name: str, address: dict, serpapi_key: str, openai_key: str, rules: str) -> dict:
        """Premium Pipeline: SerpAPI (Google) -> OpenAI"""
        # 1. SerpAPI Search
        search_data = serpapi_search(raw_name, serpapi_key)
        
        # 2. OpenAI Reason
        ai_res = self.ask_openai(raw_name, address, search_data, rules, openai_key)
        
        res = {
            "Raw Company Name": raw_name,
            "Normalized Company Name": ai_res.get("Normalized Company Name", ""),
            "Website": ai_res.get("Website", ""),
            "Industry": ai_res.get("Industry", ""),
            "Third Party Data Source Link": ai_res.get("Third Party Data Source Link", ""),
            "Remark": ai_res.get("Remark", ""),
            "Confidence Score": ai_res.get("Confidence Score", "")
        }
        return res

    def process(self, raw_name: str, address: dict) -> dict:
        """Free Pipeline: Guess -> Verify -> DDG API"""
        norm = strict_normalize_name(raw_name)
        if not norm: return {"Raw Company Name": raw_name, "Remark": "Invalid"}

        best_score = 0; best_cand = ""; best_ind = ""; best_reason = ""
        
        # Guessing
        guess_domain = norm.replace(" ", "").lower()
        guesses = [f"https://www.{guess_domain}.com", f"https://{guess_domain}.com"]
        for guess in guesses:
            res = self.verify_candidate(guess, norm)
            if res["score"] > 0.7: 
                best_score = res["score"]; best_cand = guess; best_ind = res["industry"]; best_reason = "Domain Guess"; break 

        # API Fallback
        if best_score < 0.7:
            urls = duckduckgo_search_api(norm)
            for url in urls[:3]:
                res = self.verify_candidate(url, norm)
                if res["score"] > best_score:
                    best_score = res["score"]; best_cand = url; best_ind = res["industry"]; best_reason = "API Match"
        
        official = best_cand if best_score >= 0.5 else ""
        remark = f"Verified Official ({best_reason})." if official else "Official site not found."
        
        return {
            "Raw Company Name": raw_name,
            "Normalized Company Name": norm,
            "Website": official,
            "Industry": best_ind,
            "Third Party Data Source Link": "",
            "Remark": remark,
            "Confidence Score": f"{best_score:.2f}"
        }
