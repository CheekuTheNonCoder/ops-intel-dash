"""
engine_normalize.py — Advanced Brand, Product Normalization & Registry Engine (v4.0)
Preprocesses raw names and resolves variants using 2-pass volume matching with RapidFuzz.
No "Manual Review" tags are ever prepended to product names.
"""
import re
import difflib
from collections import defaultdict
from rapidfuzz import fuzz

# ── BRAND ALIAS SYSTEM ────────────────────────────────────────
BRAND_ALIASES = {
    "zivx": "ZivX", "ziv-x": "ZivX", "ziv x": "ZivX",
    "go5": "Go5 Incorporation", "go5 incorporation": "Go5 Incorporation",
    "mavinclub": "Mavinclub", "mavin club": "Mavinclub",
    "crazybee": "Crazybee", "crazy bee": "Crazybee", "crazy-bee": "Crazybee",
    "tecsox": "TecSox", "tec sox": "TecSox",
    "texovera": "TexoVera", "texo vera": "TexoVera",
    "shopeeq": "ShopeeQ", "shopeeq.com": "ShopeeQ",
    "shopzone": "Shopzone", "shop zone": "Shopzone",
    "getsetwear": "GetsetWear", "getset wear": "GetsetWear",
    "getsetwear clothing": "GetsetWear",
    "alphawalk services llp": "ALPHAWALK SERVICES LLP",
    "digimate": "DIGIMATE",
    "heganwalk": "Heganwalk",
    "urban owl": "Urban Owl",
    "urban style": "Urban Style",
    "noyomi": "NOYOMI", "noymi": "NOYOMI",
    "werox": "Werox",
    "woggles": "Woggles",
    "beyoung": "Beyoung",
    "beyoung folks private limited": "Beyoung",
    "zyntrix lifestyle llp": "ZYNTRIX LIFESTYLE LLP",
    "trending youth technologies private limited": "TRENDING YOUTH",
    "super clone enterprises (opc) private limited": "Super Clone Enterprises",
    "super clone enterprises": "Super Clone Enterprises",
    "campussutra retail private limited": "CampusSutra",
    "dealcliq technology private limited": "DEALCLIQ",
    "aeloria trillion partners llp": "AELORIA",
    "geekverse ventures llp": "GEEKVERSE",
    "alluvium retail llp": "Alluvium Retail",
    "lili origin private limited": "Lili Origin",
    "the hatke": "THE HATKE",
    "apic inc.": "APIC Inc", "apic inc": "APIC Inc",
    "hichkie": "Hichkie",
    "boenjoy gifts": "Boenjoy Gifts",
    "vivek": "VIVEK",
    "yoga bar": "Yoga Bar",
    "dabster international private limited": "DABSTER",
    "kunsh cosmetics": "KUNSH COSMETICS",
    "haritu": "Haritu",
    "auxa": "Auxa",
    "deodap": "DeoDap",
    "dynamic marketting solution": "DYNAMIC MARKETING",
    "mayank kwatra": "MAYANK KWATRA",
    "sunmoon organics": "Sunmoon Organics",
    "by naked fact": "By Naked Fact",
    "dhanta wellness private limited": "DHANTA WELLNESS",
    "sproutlife foods pvt ltd": "Sproutlife Foods",
    "belogical wellness private limited": "Belogical Wellness",
    "savani dayaben": "SAVANI",
    "arogyasiddhi": "Arogyasiddhi",
    "beatfus products private limited": "BEATFUS",
    "nuts delish private limited": "NUTS DELISH",
    "gemies consumer private limited": "GEMIES",
    "almondzo": "Almondzo",
    "poptopia foods private ltd": "Poptopia Foods",
    "visage lines personal care pvt ltd": "Visage Lines",
    "manor rama care private limited": "Manor Rama Care",
    "bombay shaving company": "Bombay Shaving Company",
    "inlief": "In'lief", "in'lief": "In'lief",
    "branta": "Branta",
    "zibri india private limited": "ZIBRI INDIA",
    "lakshita gupta": "Lakshita Gupta",
    "jugal kant sharma": "JUGAL KANT",
    "mumma's life": "Mumma's Life",
    "hirolas": "Hirolas",
    "sneakare": "Sneakare",
    "bevzilla": "Bevzilla",
    "vrd spice private limited": "VRD SPICE",
    "kyari": "Kyari",
    "oleyy lifestyle": "Oleyy Lifestyle",
    "sanfe": "Sanfe",
    "redroomtechnology private limited": "Redroom Technology",
    "sanatan santaram gupta": "SANATAN SANTARAM GUPTA",
    "order status": None,
}

# ── PREPROCESSING REMOVAL STOPWORDS (STEP 1) ──────────────────
STOPWORDS = {
    "wireless", "bluetooth", "bt", "tws", "earbuds", "headphone", "headphones",
    "portable", "smart", "series", "edition", "black", "white", "blue", "red",
    "pro", "plus", "gen", "generation", "original", "premium", "latest", "version"
}


def normalize_brand_name(raw):
    if not isinstance(raw, str):
        return "Unmapped Brand"
    cleaned = raw.strip().strip('"').strip("'")
    if not cleaned or cleaned.lower() in ("not found", "nan", "none", "", "n/a", "order status"):
        return "Unmapped Brand"
        
    brand_lower = cleaned.lower()
    
    # Substring safeguard mapping for ZivX
    if "zivx" in brand_lower or "ziv-x" in brand_lower or ("ziv" in brand_lower and "x" in brand_lower):
        return "ZivX"
        
    key = brand_lower.strip()
    if key in BRAND_ALIASES:
        result = BRAND_ALIASES[key]
        return result if result else "Unmapped Brand"
    matches = difflib.get_close_matches(key, list(BRAND_ALIASES.keys()), n=1, cutoff=0.88)
    if matches:
        result = BRAND_ALIASES[matches[0]]
        return result if result else "Unmapped Brand"
    return cleaned


def clean_product_name(name):
    """Step 1 — Product Cleaning: lowercase, remove brackets, punctuation, dashes, extra spaces & stopwords."""
    if not isinstance(name, str):
        return ""
    # Convert to lowercase
    s = name.lower()
    # Remove brackets
    s = re.sub(r'[\(\)\[\]\{\}]', ' ', s)
    # Remove punctuation, dashes and replace with spaces
    s = re.sub(r'[^\w\s]', ' ', s)
    # Split, strip extra spaces, and filter stop words
    tokens = s.split()
    filtered_tokens = [t for t in tokens if t not in STOPWORDS]
    return ' '.join(filtered_tokens)


def has_keyword_overlap(s1, s2):
    """Keyword overlap validation: checks if both names share at least one alphabetic token > 1 char."""
    tokens1 = {t for t in s1.split() if t.isalpha() and len(t) > 1}
    tokens2 = {t for t in s2.split() if t.isalpha() and len(t) > 1}
    return len(tokens1 & tokens2) > 0


def extract_sku(raw):
    if not isinstance(raw, str):
        return None
    patterns = [
        r'\b([A-Z]{2,6}-[A-Z0-9]{2,12}-\d{2,3})\b',
        r'\b([A-Z]{2,6}-[A-Z0-9]{2,12})\b',
    ]
    for pat in patterns:
        m = re.search(pat, raw.upper())
        if m:
            return m.group(1)
    return None


class ProductRegistry:
    def __init__(self):
        # Temp raw inputs storage - Standard serializable dicts to prevent Streamlit caching crashes
        self.raw_delivered = {}
        self.raw_tickets = {}
        
        # Final resolved dictionary maps
        self.resolved_map = {}
        self.final_groups = {}
        self.debug_log = []
        
        # Validation Metrics (Step 7)
        self.total_before = 0
        self.total_after = 0
        self.auto_matched_count = 0
        self.manual_review_count = 0
        self.avg_confidence_score = 0.0

    def record_delivered(self, brand, raw_product, count=1):
        brand_str = str(brand) # Force brand strictly to string key
        if not isinstance(raw_product, str):
            return
        raw_clean = raw_product.strip().strip('"').strip("'")
        if raw_clean:
            if brand_str not in self.raw_delivered:
                self.raw_delivered[brand_str] = {}
            if raw_clean not in self.raw_delivered[brand_str]:
                self.raw_delivered[brand_str][raw_clean] = 0
            self.raw_delivered[brand_str][raw_clean] += count

    def record_ticket(self, brand, raw_product):
        brand_str = str(brand) # Force brand strictly to string key
        if not isinstance(raw_product, str):
            return
        raw_clean = raw_product.strip().strip('"').strip("'")
        if raw_clean:
            if brand_str not in self.raw_tickets:
                self.raw_tickets[brand_str] = set()
            self.raw_tickets[brand_str].add(raw_clean)

    def resolve(self):
        """Step 5 — 2-pass Brand-aware volume matching. Groups products based on highest delivered volume first."""
        self.debug_log = []
        self.resolved_map = {}
        self.final_groups = {}
        
        total_pairs_evaluated = 0
        total_confidence_sum = 0.0
        
        # Cast keys strictly to str
        all_brands = {str(b) for b in self.raw_delivered.keys()} | {str(b) for b in self.raw_tickets.keys()}
        
        for brand in all_brands:
            brand_str = str(brand)
            delivered_counts = self.raw_delivered.get(brand_str, {})
            ticket_products = self.raw_tickets.get(brand_str, set())
            
            all_raw = set(delivered_counts.keys()) | ticket_products
            if not all_raw:
                continue
                
            # Step 5 — Sort raw products by delivered volume (descending) 
            # Highest volume is evaluated first and naturally becomes the canonical name [R5].
            sorted_raw = sorted(list(all_raw), key=lambda x: delivered_counts.get(x, 0), reverse=True)
            
            # Temporary mapping bucket for this brand
            canonical_products = {} # cname -> { norm, variants, volume, sku }
            self.resolved_map[brand_str] = {}
            
            for raw_name in sorted_raw:
                norm = clean_product_name(raw_name)
                vol  = delivered_counts.get(raw_name, 0)
                sku  = extract_sku(raw_name)
                
                if not norm:
                    canonical_products[raw_name] = {
                        "norm": raw_name.lower(),
                        "variants": {raw_name},
                        "volume": vol,
                        "sku": sku
                    }
                    self.resolved_map[brand_str][raw_name] = raw_name
                    self.debug_log.append({
                        "Raw Product": raw_name,
                        "Canonical Product": raw_name,
                        "Confidence": 100.0,
                        "Mapping Method": "Direct Fallback (Empty Norm)"
                    })
                    continue
                
                best_score = 0.0
                best_match = None
                sku_matched = False
                
                # Try SKU match first
                if sku:
                    for cname, cdata in canonical_products.items():
                        if cdata.get("sku") == sku:
                            best_match = cname
                            best_score = 100.0
                            sku_matched = True
                            break
                
                if not sku_matched:
                    # Fuzzy matching evaluation using RapidFuzz
                    for cname, cdata in canonical_products.items():
                        t_set   = fuzz.token_set_ratio(norm, cdata["norm"])
                        t_sort  = fuzz.token_sort_ratio(norm, cdata["norm"])
                        p_ratio = fuzz.partial_ratio(norm, cdata["norm"])
                        
                        # Perfect Token Subset Shortcut rule: If one product's tokens are a perfect subset of another 
                        # and they share a significant keyword, they are an automatic 100% match!
                        if t_set == 100.0 and has_keyword_overlap(norm, cdata["norm"]):
                            score = 100.0
                        else:
                            # Standard weighted formula (Step 2)
                            score = (0.40 * t_set) + (0.30 * t_sort) + (0.30 * p_ratio)
                        
                        if score > best_score:
                            best_score = score
                            best_match = cname
                
                # Step 3 — Auto Match Threshold evaluations
                matched = False
                method = "New Product Created"
                
                if best_match:
                    if best_score >= 90:
                        matched = True
                        method = "Auto Match (>=90)"
                        self.auto_matched_count += 1
                    elif 80 <= best_score < 90:
                        if has_keyword_overlap(norm, canonical_products[best_match]["norm"]):
                            matched = True
                            method = "Auto Match (80-89: Overlap Passed)"
                            self.auto_matched_count += 1
                        else:
                            method = "Manual Review Required (80-89: No Overlap)"
                            self.manual_review_count += 1
                    else:
                        method = "Manual Review Required (<80)"
                        self.manual_review_count += 1
                
                if matched:
                    # Map to the group's established parent (guaranteed to have the highest delivered volume!) [R5]
                    canonical_products[best_match]["variants"].add(raw_name)
                    canonical_products[best_match]["volume"] += vol
                    self.resolved_map[brand_str][raw_name] = best_match
                    
                    total_pairs_evaluated += 1
                    total_confidence_sum += best_score
                    
                    self.debug_log.append({
                        "Raw Product": raw_name,
                        "Canonical Product": best_match,
                        "Confidence": round(best_score, 1),
                        "Mapping Method": method
                    })
                else:
                    # No Match: Map to itself (No manual prefixing. Name remains clean and standalone) [R3]
                    canonical_products[raw_name] = {
                        "norm": norm,
                        "variants": {raw_name},
                        "volume": vol,
                        "sku": sku
                    }
                    self.resolved_map[brand_str][raw_name] = raw_name
                    
                    self.debug_log.append({
                        "Raw Product": raw_name,
                        "Canonical Product": raw_name,
                        "Confidence": round(best_score, 1) if best_match else 0.0,
                        "Mapping Method": method
                    })
            
            # Save final compiled groups for Step 6 Output [R6]
            self.final_groups[brand_str] = {}
            for cname, cdata in canonical_products.items():
                self.final_groups[brand_str][cname] = {
                    "variants": cdata["variants"],
                    "delivered": cdata["volume"],
                    "tickets": 0,
                    "sku": cdata["sku"] if cdata["sku"] else "N/A"
                }
                
        # Calculate Validation metrics (Step 7)
        self.total_before = sum(len(self.raw_delivered.get(b, {})) | len(self.raw_tickets.get(b, set())) for b in all_brands)
        self.total_after = sum(len(self.final_groups.get(b, {})) for b in all_brands)
        self.avg_confidence_score = round(total_confidence_sum / max(total_pairs_evaluated, 1), 1)

    def summary_df(self):
        import pandas as pd
        rows = []
        for brand, groups in self.final_groups.items():
            for cname, data in groups.items():
                rows.append({
                    "Brand": str(brand),
                    "Canonical Product": cname,
                    "SKU": data.get("sku", "N/A"),
                    "Variants": " | ".join(sorted(data["variants"]))[:300],
                    "Delivered Orders": data["delivered"],
                    "Tickets": data["tickets"]
                })
        return pd.DataFrame(rows)