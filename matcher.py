"""
Twitter/X Profile Matcher
-------------------------
Given a person's name, address, and employer, searches for their
Twitter/X profile and returns a confidence-scored match.

Confidence scoring model
------------------------
Each signal adds points toward a 0–100 score:

  Name match (exact)          +40
  Name match (partial)        +20
  Employer in bio             +25
  City in bio/location        +20
  State in bio/location       +10
  Bio keyword overlap         +5 per keyword (max 15)
  Verified account            +5
  Account age > 2 years       +5

Thresholds:
  HIGH    >= 75   (green)
  MEDIUM  45–74   (yellow)
  LOW     20–44   (orange)
  NO MATCH < 20   (red)
"""

import re
import time
import logging
from dataclasses import dataclass, field
from typing import Optional
import requests

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Data models
# ---------------------------------------------------------------------------

@dataclass
class Person:
    first_name: str
    last_name: str
    address: str = ""
    employer: str = ""

    @property
    def full_name(self) -> str:
        return f"{self.first_name.strip()} {self.last_name.strip()}".strip()

    @property
    def city(self) -> str:
        """Best-effort city extraction from address string."""
        parts = [p.strip() for p in self.address.split(",")]
        if len(parts) >= 2:
            return parts[-2]
        return parts[0] if parts else ""

    @property
    def state(self) -> str:
        """Best-effort US state extraction from address string."""
        US_STATES = {
            "AL","AK","AZ","AR","CA","CO","CT","DE","FL","GA","HI","ID","IL","IN",
            "IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV",
            "NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","RI","SC","SD","TN",
            "TX","UT","VT","VA","WA","WV","WI","WY","DC",
        }
        parts = [p.strip() for p in self.address.split(",")]
        last = parts[-1] if parts else ""
        for t in last.split():
            if t.upper() in US_STATES:
                return t.upper()
        return ""


@dataclass
class TwitterProfile:
    handle: str
    display_name: str
    bio: str
    location: str
    followers: int
    verified: bool
    created_year: int
    profile_url: str
    avatar_url: str = ""
    source: str = "search"


@dataclass
class MatchResult:
    person: Person
    profile: Optional[TwitterProfile]
    score: int
    confidence: str          # HIGH / MEDIUM / LOW / NO MATCH
    score_breakdown: dict = field(default_factory=dict)
    search_query: str = ""
    error: str = ""

    @property
    def color(self) -> str:
        return {
            "HIGH":     "#22c55e",
            "MEDIUM":   "#eab308",
            "LOW":      "#f97316",
            "NO MATCH": "#ef4444",
        }.get(self.confidence, "#6b7280")

    @property
    def emoji(self) -> str:
        return {
            "HIGH":     "✅",
            "MEDIUM":   "🟡",
            "LOW":      "🟠",
            "NO MATCH": "❌",
        }.get(self.confidence, "❓")


# ---------------------------------------------------------------------------
# Confidence scorer
# ---------------------------------------------------------------------------

def _normalise(text: str) -> str:
    return re.sub(r"[^a-z0-9 ]", " ", text.lower())


def score_match(person: Person, profile: TwitterProfile) -> tuple[int, dict]:
    breakdown = {}
    total = 0

    full = person.full_name.lower()
    disp = profile.display_name.lower()
    handle_clean = profile.handle.lstrip("@").lower()
    bio_norm = _normalise(profile.bio)
    loc_norm = _normalise(profile.location)
    combined = bio_norm + " " + loc_norm

    # ── Name matching ────────────────────────────────────────────────────────
    first_l = person.first_name.lower()
    last_l  = person.last_name.lower()

    if full in disp or disp in full:
        pts = 40
        breakdown["name_exact"] = pts
    elif last_l in disp and first_l in disp:
        pts = 35
        breakdown["name_both_parts"] = pts
    elif last_l in disp or last_l in handle_clean:
        pts = 20
        breakdown["name_last_only"] = pts
    elif first_l in disp:
        pts = 10
        breakdown["name_first_only"] = pts
    else:
        pts = 0
    total += pts

    # ── Employer ─────────────────────────────────────────────────────────────
    if person.employer:
        emp_words = [w for w in _normalise(person.employer).split() if len(w) > 3]
        matches = sum(1 for w in emp_words if w in bio_norm)
        if emp_words and matches / len(emp_words) >= 0.5:
            pts = 25
            breakdown["employer_match"] = pts
            total += pts
        elif matches > 0:
            pts = 10
            breakdown["employer_partial"] = pts
            total += pts

    # ── City ─────────────────────────────────────────────────────────────────
    if person.city:
        city_norm = _normalise(person.city)
        if city_norm and city_norm in combined:
            pts = 20
            breakdown["city_match"] = pts
            total += pts

    # ── State ────────────────────────────────────────────────────────────────
    if person.state:
        state_norm = person.state.lower()
        if state_norm in combined:
            pts = 10
            breakdown["state_match"] = pts
            total += pts

    # ── Bio keyword overlap ──────────────────────────────────────────────────
    kw_sources = [person.employer, person.city]
    keywords = set()
    for src in kw_sources:
        for word in _normalise(src).split():
            if len(word) > 4:
                keywords.add(word)

    kw_hits = sum(1 for kw in keywords if kw in bio_norm)
    if kw_hits:
        pts = min(kw_hits * 5, 15)
        breakdown["bio_keywords"] = pts
        total += pts

    # ── Verified ─────────────────────────────────────────────────────────────
    if profile.verified:
        breakdown["verified"] = 5
        total += 5

    # ── Account age ──────────────────────────────────────────────────────────
    if profile.created_year and profile.created_year <= 2022:
        breakdown["account_age"] = 5
        total += 5

    total = min(total, 100)

    if total >= 75:
        confidence = "HIGH"
    elif total >= 45:
        confidence = "MEDIUM"
    elif total >= 20:
        confidence = "LOW"
    else:
        confidence = "NO MATCH"

    return total, breakdown


# ---------------------------------------------------------------------------
# Search backends
# ---------------------------------------------------------------------------

def _build_queries(person: Person) -> list[str]:
    """Generate a ranked list of search queries from most to least specific."""
    queries = []
    name = person.full_name
    city = person.city
    employer = person.employer

    if employer and city:
        queries.append(f'"{name}" "{employer}" "{city}" site:twitter.com OR site:x.com')
        queries.append(f'"{name}" "{employer}" site:twitter.com OR site:x.com')
    if employer:
        queries.append(f'"{name}" "{employer}" twitter')
    if city:
        queries.append(f'"{name}" "{city}" twitter')
    queries.append(f'"{name}" twitter')
    queries.append(f'{name} site:twitter.com')
    return queries


def _parse_twitter_url(url: str) -> Optional[str]:
    """Extract @handle from a twitter.com or x.com URL."""
    m = re.search(r'(?:twitter\.com|x\.com)/([A-Za-z0-9_]{1,50})(?:/|$|\?)', url)
    if m:
        handle = m.group(1)
        if handle.lower() not in ("intent", "search", "hashtag", "share", "home",
                                   "explore", "notifications", "messages", "i"):
            return "@" + handle
    return None


def search_via_serpapi(person: Person, api_key: str) -> list[dict]:
    """Use SerpAPI Google Search to find candidate Twitter profiles."""
    candidates = []
    seen_handles = set()

    for query in _build_queries(person)[:3]:
        try:
            resp = requests.get(
                "https://serpapi.com/search",
                params={"q": query, "api_key": api_key, "num": 5, "engine": "google"},
                timeout=10
            )
            if resp.status_code != 200:
                continue
            data = resp.json()
            for result in data.get("organic_results", []):
                url = result.get("link", "")
                handle = _parse_twitter_url(url)
                if handle and handle not in seen_handles:
                    seen_handles.add(handle)
                    candidates.append({
                        "handle": handle,
                        "title": result.get("title", ""),
                        "snippet": result.get("snippet", ""),
                        "url": url,
                        "query": query,
                    })
            time.sleep(0.3)
        except Exception as e:
            logger.warning(f"SerpAPI error: {e}")
            continue

    return candidates


def search_via_google_cse(person: Person, api_key: str, cx: str) -> list[dict]:
    """Use Google Custom Search Engine API."""
    candidates = []
    seen_handles = set()

    for query in _build_queries(person)[:3]:
        try:
            resp = requests.get(
                "https://www.googleapis.com/customsearch/v1",
                params={"q": query, "key": api_key, "cx": cx, "num": 5},
                timeout=10
            )
            if resp.status_code != 200:
                continue
            data = resp.json()
            for item in data.get("items", []):
                url = item.get("link", "")
                handle = _parse_twitter_url(url)
                if handle and handle not in seen_handles:
                    seen_handles.add(handle)
                    candidates.append({
                        "handle": handle,
                        "title": item.get("title", ""),
                        "snippet": item.get("snippet", ""),
                        "url": url,
                        "query": query,
                    })
            time.sleep(0.3)
        except Exception as e:
            logger.warning(f"Google CSE error: {e}")
            continue

    return candidates


def search_via_claude(person: Person, anthropic_key: str) -> list[dict]:
    """
    Use Claude with web search tool to find Twitter profiles.
    This is the primary backend — works without any extra API keys
    beyond the Anthropic key the team already has.
    """
    import anthropic

    client = anthropic.Anthropic(api_key=anthropic_key)

    queries = _build_queries(person)
    candidates = []
    seen_handles = set()

    prompt = f"""Find the Twitter/X profile for this person using web search.

Person details:
- Full name: {person.full_name}
- Address/City: {person.address}
- Employer: {person.employer}

Search for their Twitter or X.com profile. Try multiple searches if needed.
Return ONLY a JSON array of candidate profiles found. Each object must have:
  - handle: the @username (e.g. "@johndoe")
  - display_name: the name shown on the profile
  - bio: their Twitter bio text
  - location: location shown on profile
  - followers: follower count as integer (0 if unknown)
  - verified: true/false
  - created_year: year account was created as integer (0 if unknown)
  - profile_url: full URL to profile
  - confidence_notes: brief note on why this might or might not be the right person

Return ONLY the JSON array, no other text. If no profiles found, return [].
Try at least 3 different search queries before giving up."""

    try:
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=2000,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": prompt}]
        )

        # Extract text from response
        full_text = ""
        for block in response.content:
            if hasattr(block, "text"):
                full_text += block.text

        # Parse JSON from response
        import json
        # Strip markdown fences if present
        clean = re.sub(r"```(?:json)?", "", full_text).strip().strip("`").strip()
        # Find JSON array
        m = re.search(r'\[.*\]', clean, re.DOTALL)
        if m:
            raw_list = json.loads(m.group(0))
            for item in raw_list:
                handle = item.get("handle", "").strip()
                if not handle.startswith("@"):
                    handle = "@" + handle
                if handle and handle not in seen_handles and handle != "@":
                    seen_handles.add(handle)
                    candidates.append({
                        "handle": handle,
                        "title": item.get("display_name", ""),
                        "snippet": item.get("bio", ""),
                        "url": item.get("profile_url", f"https://twitter.com/{handle.lstrip('@')}"),
                        "query": "claude_web_search",
                        "raw": item,
                    })
    except Exception as e:
        logger.warning(f"Claude search error for {person.full_name}: {e}")

    return candidates


# ---------------------------------------------------------------------------
# Profile builder from search snippet
# ---------------------------------------------------------------------------

def build_profile_from_candidate(candidate: dict) -> TwitterProfile:
    """
    Build a TwitterProfile from a search result candidate.
    When we have rich data from Claude's search (raw dict), use it directly.
    Otherwise infer from title/snippet.
    """
    raw = candidate.get("raw", {})
    handle = candidate.get("handle", "")

    if raw:
        return TwitterProfile(
            handle=handle,
            display_name=raw.get("display_name", "") or candidate.get("title", ""),
            bio=raw.get("bio", "") or candidate.get("snippet", ""),
            location=raw.get("location", ""),
            followers=int(raw.get("followers", 0) or 0),
            verified=bool(raw.get("verified", False)),
            created_year=int(raw.get("created_year", 0) or 0),
            profile_url=raw.get("profile_url", "") or candidate.get("url", ""),
            source="claude_search",
        )

    # Fallback: parse from search snippet
    title = candidate.get("title", "")
    snippet = candidate.get("snippet", "")
    display_name = re.sub(r'\s*\(@[^)]+\)', '', title).split(" | ")[0].strip()

    return TwitterProfile(
        handle=handle,
        display_name=display_name,
        bio=snippet,
        location="",
        followers=0,
        verified=False,
        created_year=0,
        profile_url=candidate.get("url", f"https://twitter.com/{handle.lstrip('@')}"),
        source="search_snippet",
    )


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def find_twitter(
    person: Person,
    anthropic_key: str = "",
    serpapi_key: str = "",
    google_cse_key: str = "",
    google_cx: str = "",
) -> MatchResult:
    """
    Find a person's Twitter profile. Tries available backends in order.
    Returns the best-scoring match above NO MATCH threshold, or the top
    candidate with a NO MATCH result if nothing scores well.
    """
    candidates = []

    # Try Claude web search first (best quality, uses existing Anthropic key)
    if anthropic_key:
        candidates.extend(search_via_claude(person, anthropic_key))

    # Fall back to SerpAPI
    if not candidates and serpapi_key:
        candidates.extend(search_via_serpapi(person, serpapi_key))

    # Fall back to Google CSE
    if not candidates and google_cse_key and google_cx:
        candidates.extend(search_via_google_cse(person, google_cse_key, google_cx))

    if not candidates:
        return MatchResult(
            person=person,
            profile=None,
            score=0,
            confidence="NO MATCH",
            error="No search backends available or no results returned.",
        )

    # Score all candidates, pick the best
    scored = []
    for cand in candidates:
        profile = build_profile_from_candidate(cand)
        score, breakdown = score_match(person, profile)
        scored.append((score, breakdown, profile, cand.get("query", "")))

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best_breakdown, best_profile, best_query = scored[0]

    if best_score < 20:
        confidence = "NO MATCH"
    elif best_score < 45:
        confidence = "LOW"
    elif best_score < 75:
        confidence = "MEDIUM"
    else:
        confidence = "HIGH"

    return MatchResult(
        person=person,
        profile=best_profile,
        score=best_score,
        confidence=confidence,
        score_breakdown=best_breakdown,
        search_query=best_query,
    )
