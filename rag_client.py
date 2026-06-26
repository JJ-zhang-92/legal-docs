"""
Local RAG client for the legal knowledge base API.
Queries the local law-rag service at http://localhost:8720.
Falls back to statutes.py if the service is unavailable.

Usage:
    from rag_client import query
    result = query("劳动合同试用期最长多久", top=3)
    print(result)  # Returns full statute text, safe for AI context
"""

import urllib.request
import urllib.error
import urllib.parse
import http.client
import json
import os
import re
import sys

# Ensure skill directory is on sys.path for fallback import
_skill_dir = os.path.dirname(os.path.abspath(__file__))
if _skill_dir not in sys.path:
    sys.path.insert(0, _skill_dir)

RAG_URL = "http://localhost:8720"


def query(question, top=5, fmt="text"):
    """
    Query the local legal RAG service.

    Args:
        question: Natural language question or keywords
        top: Number of results (default 5)
        fmt: Response format — "text" for paste-into-prompt, "json" for programmatic use

    Returns:
        String: Statute text if successful, or None if service unreachable.
        When service is unreachable, falls back to statutes.py keyword search.
    """
    try:
        params = urllib.parse.urlencode({'q': question, 'top': top, 'format': fmt})
        url = f"{RAG_URL}?{params}"
        req = urllib.request.Request(url, headers={'User-Agent': 'legal-docs/1.0'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = resp.read().decode('utf-8')
            # Validate: reject empty or HTML error responses
            if not data or not data.strip():
                return _fallback_statutes(question, top)
            if data.strip().startswith('<') or data.strip().startswith('<!'):
                return _fallback_statutes(question, top)
            return data
    except (urllib.error.URLError, urllib.error.HTTPError,
            OSError, http.client.HTTPException) as e:
        return _fallback_statutes(question, top)


def _fallback_statutes(question, top=5):
    """Fallback: keyword-based search in statutes.py.
    Tokenizes natural-language questions into meaningful keywords."""
    try:
        from statutes import search_statutes

        # Tokenize: extract legal-relevant 2-4 char Chinese substrings
        # Remove common question words to get meaningful search terms
        stopwords = {'什么', '怎么', '多少', '如何', '哪个', '请问', '是否',
                     '最长', '最短', '多久', '的', '了', '吗', '呢', '啊',
                     '应该', '可以', '需要'}
        tokens = []
        for i in range(len(question)):
            for length in (2, 3, 4):
                if i + length <= len(question):
                    token = question[i:i+length]
                    if token not in stopwords and not all(c in '，。？！、：；""''（）' for c in token):
                        tokens.append(token)

        # Use longest unique tokens as keywords
        keywords = list(dict.fromkeys(sorted(tokens, key=len, reverse=True)))[:8]
        results = search_statutes(keywords, top_k=top)

        if not results:
            return None
        lines = ["[RAG 服务未启动，以下结果来自本地 statutes.py 关键词搜索]\n"]
        for r in results:
            lines.append(f"【{r['statute']}·{r['chapter']}】{r['article']}：{r['summary']}")
        return '\n'.join(lines)
    except ImportError:
        return None


def check_rag_status():
    """Check if the RAG service is running. Returns (bool, message)."""
    try:
        params = urllib.parse.urlencode({'q': 'ping', 'top': 1, 'format': 'text'})
        url = f"{RAG_URL}?{params}"
        req = urllib.request.Request(url, headers={'User-Agent': 'legal-docs/1.0'})
        with urllib.request.urlopen(req, timeout=3) as resp:
            data = resp.read().decode('utf-8')
            # Validate: response must be non-empty and not an HTML error page
            if not data or not data.strip():
                return False, "RAG 服务返回空响应"
            if data.strip().startswith('<'):
                return False, "RAG 服务返回 HTML 错误页"
            return True, f"RAG 服务运行中 ({RAG_URL})"
    except (urllib.error.URLError, urllib.error.HTTPError,
            OSError, http.client.HTTPException) as e:
        return False, f"RAG 服务未启动: {e}"


if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Usage: python rag_client.py <question> [top]")
        print("       python rag_client.py --status")
        sys.exit(1)

    if sys.argv[1] == '--status':
        ok, msg = check_rag_status()
        print(msg)
    else:
        q = sys.argv[1]
        k = int(sys.argv[2]) if len(sys.argv) > 2 else 5
        result = query(q, top=k)
        if result:
            print(result.encode('utf-8', errors='replace').decode('utf-8'))
        else:
            print("无法获取法规信息：RAG 服务未启动且 statutes 回退失败")
