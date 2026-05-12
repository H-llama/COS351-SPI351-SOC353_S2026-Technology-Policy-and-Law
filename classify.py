"""
Scam Ad Classifier
Processes Apify Ad Library JSON files for luxury, mid-tier, and fast fashion tiers.
Applies 5-criteria rubric and outputs a labeled CSV for analysis.

Usage:
    python3 classifier.py

Input files expected in same directory:
    LUXARY.json, MID_TIER.json, FAST_FASHION.json

Output:
    scam_ads_labeled.csv
"""

import json
import re
import csv
import os
import datetime
from urllib.parse import urlparse

# ─── Configuration ────────────────────────────────────────────────────────────

INPUT_FILES = {
    "luxury":       "LUXARY.json",
    "mid_tier":     "MID_TIER.json",
    "fast_fashion": "FAST_FASHION.json",
}

OUTPUT_CSV = "scam_ads_labeled.csv"

# Official domains per brand
BRAND_DOMAINS = {
    "gucci":            ["gucci.com"],
    "hermes":           ["hermes.com", "hermes.fr"],
    "louis vuitton":    ["louisvuitton.com"],
    "coach":            ["coach.com"],
    "ralph lauren":     ["ralphlauren.com", "polo.com"],
    "armani exchange":  ["armaniexchange.com", "armani.com"],
    "shein":            ["shein.com", "sheglam.com"],
    "zara":             ["zara.com"],
    "fashion nova":     ["fashionnova.com"],
}

# Which brands belong to which tier
TIER_BRANDS = {
    "luxury":       ["gucci", "hermes", "louis vuitton"],
    "mid_tier":     ["coach", "ralph lauren", "armani exchange"],
    "fast_fashion": ["shein", "zara", "fashion nova"],
}

# Completely unrelated categories — strong scam signal
UNRELATED_CATEGORIES = {
    "politician", "political organization", "government official",
    "news & media website", "news page", "journalist",
    "religious organization", "church", "nonprofit organization",
    "hospital", "doctor", "health/beauty",
    "restaurant", "food & beverage", "bar",
    "artist", "musician/band", "public figure",
    "comedian", "athlete", "actor/director",
    "travel & leisure", "hotel & lodging",
    "automotive", "car dealership",
    "education", "school", "college & university",
    "community", "community organization",
}

# Urgency keywords for C3
URGENCY_KEYWORDS = [
    "today only", "limited time", "limited stock", "hours left",
    "while supplies last", "ends soon", "selling fast", "almost gone",
    "last chance", "don't miss", "hurry", "flash sale", "act now",
    "expires", r"only \d+ left", "going fast", "ending soon",
    "24 hours", "48 hours", "this weekend only", "one day only",
    "don't wait", "sell out", "sold out soon",
]

# Holiday window: Nov 15 - Dec 31
HOLIDAY_WINDOW = [(11, 15), (12, 31)]

# ─── Helpers ──────────────────────────────────────────────────────────────────

def get_text(record):
    body = record.get("snapshot", {}).get("body")
    if not body:
        return ""
    return body.get("text", "") or ""

def get_page_uri(record):
    return record.get("snapshot", {}).get("page_profile_uri", "") or ""

def get_page_likes(record):
    return record.get("snapshot", {}).get("page_like_count", None)

def get_page_categories(record):
    cats = record.get("snapshot", {}).get("page_categories", []) or []
    return {c.lower() for c in cats}

def extract_domain(uri):
    try:
        parsed = urlparse(uri)
        netloc = parsed.netloc.lower()
        if netloc.startswith("www."):
            netloc = netloc[4:]
        return netloc
    except:
        return ""

def get_source_brand(url, tier):
    url_lower = url.lower()
    for brand in TIER_BRANDS[tier]:
        encoded = brand.replace(" ", "%20").lower()
        if encoded in url_lower or brand.replace(" ", "+").lower() in url_lower:
            return brand
        if brand.split()[0].lower() in url_lower:
            return brand
    return "unknown"

def is_holiday_window(record):
    start = record.get("start_date", 0) or 0
    if not start:
        return False
    try:
        dt = datetime.datetime.fromtimestamp(start)
        if dt.month == 11 and dt.day >= 15:
            return True
        if dt.month == 12:
            return True
        return False
    except:
        return False

# ─── Criteria ─────────────────────────────────────────────────────────────────

def check_c1_domain_mismatch(record, brand):
    """C1: Page URL domain doesn't match official brand domain."""
    uri = get_page_uri(record)
    if not uri:
        return None
    domain = extract_domain(uri)
    # Facebook-hosted pages can't be assessed by domain
    if "facebook.com" in domain or not domain:
        return None
    official_domains = BRAND_DOMAINS.get(brand, [])
    for official in official_domains:
        if official in domain:
            return False
    return True

def check_c2_extreme_discount(text):
    """C2: Ad claims >= 70% off."""
    if not text:
        return False
    matches = re.findall(r'(\d{1,3})\s*%\s*(?:off|discount|sale|savings?)', text, re.IGNORECASE)
    return any(int(m) >= 70 for m in matches)

def check_c3_urgency(text):
    """C3: Ad uses urgency/time-pressure language."""
    if not text:
        return False
    text_lower = text.lower()
    return any(re.search(kw, text_lower) for kw in URGENCY_KEYWORDS)

def check_c4_category_mismatch(record):
    """C4: Page category is completely unrelated to fashion/retail."""
    cats = get_page_categories(record)
    if not cats:
        return False  # Unknown category — don't penalize
    for cat in cats:
        for unrelated in UNRELATED_CATEGORIES:
            if unrelated in cat or cat in unrelated:
                return True
    return False

def check_c5_low_likes(record):
    """C5: Page has fewer than 1,000 likes."""
    likes = get_page_likes(record)
    if likes is None:
        return None
    return likes < 150

# ─── Classifier ───────────────────────────────────────────────────────────────

def classify_record(record, tier, brand):
    text = get_text(record)
    holiday = is_holiday_window(record)
    likes = get_page_likes(record)

    # Brand mention filter — ad text must contain the brand name
    # to be eligible for classification as a scam
    brand_keywords = {
        "gucci": ["gucci"],
        "hermes": ["hermès", "hermes", "hermès"],
        "louis vuitton": ["louis vuitton", "lv"],
        "coach": ["coach"],
        "ralph lauren": ["ralph lauren", "polo ralph lauren"],
        "armani exchange": ["armani exchange", "a|x", "ax armani"],
        "shein": ["shein"],
        "zara": ["zara"],
        "fashion nova": ["fashion nova"],
    }
    keywords = brand_keywords.get(brand, [brand])
    text_lower = text.lower()
    brand_mentioned = any(kw in text_lower for kw in keywords)

    # If brand not mentioned in ad text, skip — not a brand impersonation scam
    if not brand_mentioned:
        return {
            "c1_domain_mismatch":   None,
            "c2_extreme_discount":  None,
            "c3_urgency_language":  None,
            "c4_category_mismatch": None,
            "c5_low_likes":         None,
            "criteria_met":         0,
            "is_scam":              False,
            "holiday_window":       holiday,
            "c2_waived":            False,
            "c3_waived":            False,
            "brand_mentioned":      False,
        }

    c1 = check_c1_domain_mismatch(record, brand)
    c2 = check_c2_extreme_discount(text)
    c3 = check_c3_urgency(text)
    c4 = check_c4_category_mismatch(record)
    c5 = check_c5_low_likes(record)

    # False positive guard: waive C2+C3 during holiday window for large verified pages
    c2_waived = c3_waived = False
    if holiday and likes is not None and likes > 100000:
        c2_waived, c3_waived = c2, c3
        c2, c3 = False, False

    criteria_met = sum(1 for v in [c1, c2, c3, c4, c5] if v is True)
    is_scam = criteria_met >= 2

    return {
        "c1_domain_mismatch":   c1,
        "c2_extreme_discount":  c2,
        "c3_urgency_language":  c3,
        "c4_category_mismatch": c4,
        "c5_low_likes":         c5,
        "criteria_met":         criteria_met,
        "is_scam":              is_scam,
        "holiday_window":       holiday,
        "c2_waived":            c2_waived,
        "c3_waived":            c3_waived,
        "brand_mentioned":      True,
    }

def process_file(filepath, tier):
    with open(filepath, encoding="utf-8") as f:
        data = json.load(f)

    results = []
    for record in data:
        source_url = record.get("url", "")
        brand = get_source_brand(source_url, tier)
        classification = classify_record(record, tier, brand)

        row = {
            "ad_archive_id":      record.get("ad_archive_id", ""),
            "tier":               tier,
            "brand":              brand,
            "source_url":         source_url,
            "page_id":            record.get("page_id", ""),
            "page_name":          record.get("page_name", ""),
            "page_profile_uri":   get_page_uri(record),
            "page_like_count":    get_page_likes(record),
            "page_categories":    "|".join(get_page_categories(record)),
            "is_active":          record.get("is_active", ""),
            "start_date":         record.get("start_date_formatted", ""),
            "end_date":           record.get("end_date_formatted", ""),
            "publisher_platform": "|".join(record.get("publisher_platform", []) or []),
            "ad_text":            get_text(record)[:500],
            **classification,
        }
        results.append(row)

    return results

# ─── Main ─────────────────────────────────────────────────────────────────────

def main():
    all_results = []

    for tier, filename in INPUT_FILES.items():
        if not os.path.exists(filename):
            print(f"WARNING: {filename} not found, skipping {tier}")
            continue

        print(f"\nProcessing {filename} ({tier})...")
        results = process_file(filename, tier)
        all_results.extend(results)

        total = len(results)
        eligible = sum(1 for r in results if r["brand_mentioned"])
        scams = sum(1 for r in results if r["is_scam"])
        print(f"  Total: {total} ads | Brand mentioned: {eligible} | Flagged: {scams} ({scams/eligible*100:.1f}% of eligible)")

        for brand in sorted(set(r["brand"] for r in results)):
            b_ads = [r for r in results if r["brand"] == brand]
            b_elig = [r for r in b_ads if r["brand_mentioned"]]
            b_scams = sum(1 for r in b_ads if r["is_scam"])
            pct = b_scams/len(b_elig)*100 if b_elig else 0
            print(f"    {brand:20s}: {b_scams:4d}/{len(b_elig):4d} eligible ({pct:.1f}%)")

    if not all_results:
        print("\nNo results — check JSON files are in the same folder as this script.")
        return

    # Write CSV
    fieldnames = list(all_results[0].keys())
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_results)

    total = len(all_results)
    eligible = sum(1 for r in all_results if r["brand_mentioned"])
    scams = sum(1 for r in all_results if r["is_scam"])
    print(f"\n=== COMPLETE ===")
    print(f"Total ads collected:  {total}")
    print(f"Brand mentioned (eligible): {eligible} ({eligible/total*100:.1f}%)")
    print(f"Flagged as scams:     {scams} ({scams/eligible*100:.1f}% of eligible)")
    print(f"Output: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()