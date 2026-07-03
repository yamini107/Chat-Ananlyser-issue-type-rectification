"""
Chat Analyzer Dashboard — Shopee, Lazada & TikTok
==================================================
Graas.ai-themed Streamlit app for daily chat enquiry analysis.

CHANGES FROM PREVIOUS VERSION
──────────────────────────────
1. JSON Message Extraction (NEW)
   - Messages on all three platforms are JSON-encoded blobs.
   - New helper `extract_plain_text()` unwraps txt / text / content / ext
     fields before any keyword matching. Without this, keywords like
     "cancel" or "refund" were never found inside raw JSON strings,
     causing nearly every conversation to fall through to "Other".

2. Scoring-based Classification (IMPROVED)
   - Buyer text is matched against the ISSUE_KEYWORDS dictionary; every
     matching keyword contributes weight 1 to that issue type's score.
   - The issue type with the highest cumulative score wins.
   - Ties broken by priority order (High > Medium > Low).

3. Unresolved Chat Detection (IMPROVED)
   - Previous logic only checked seller messages for stalling/resolution
     patterns in raw JSON strings — patterns were never matching.
   - `conversation_is_unresolved()` now receives pre-parsed seller texts.
   - Added explicit resolution signals from the data: "ยินดีให้บริการ"
     (thank-you closings), "จัดส่งแล้ว" (shipped), "ดำเนินการแล้ว"
     (processed), "เรียบร้อย" (done), plus English equivalents.
   - Auto-replies (bot greetings, chatbot messages) are now excluded from
     both stall detection and resolution detection to avoid false positives.

4. Expanded Keyword Coverage & Refined Priority (NEW — this version)
   - ISSUE_KEYWORDS has been extended with a large multilingual keyword set
     (English / Thai / Bahasa / Tagalog / Chinese) sourced from the latest
     reference keyword sheet, covering Refund, Return, Cancellation, Delay,
     Damaged/Wrong Item, Missing Item, Payment Issue, Product Inquiry,
     Promotion Issue, Technical Issue and Complaint. This significantly
     reduces conversations falling through to "Other" when a genuine issue
     keyword is present in the buyer's message.
   - PRIORITY_MAP has been updated: "Missing Item" is now High priority
     (previously Medium), in line with updated business priority rules.
   - A new HIGH_PRIORITY_ESCALATION_KEYWORDS list + escalation check inside
     `get_priority()` promotes specific high-severity situations (fraud,
     account hacked, chargeback/legal threats, double charge, safety
     issues, "marked delivered but not received", etc.) to High priority
     even when their underlying Issue Type would normally be Medium/Low —
     without changing the underlying Issue Type or Issue Type keyword
     matching logic.

ASSUMPTIONS
───────────
- "too big" / "too small" / "doesn't fit" style keywords are mapped to
  "Return" (size-related returns) rather than "Product Inquiry".
- Messages that are purely media (images, order cards, item cards) with
  no extractable text are skipped during issue classification.
- Escalation keywords (fraud, hacked account, legal/chargeback threats,
  double charge, safety issues, false "delivered" status) bump a
  conversation's PRIORITY to "High" regardless of its detected Issue Type,
  since these situations require urgent handling even when they surface
  inside an otherwise Medium/Low-priority issue category (e.g. a double
  charge inside "Payment Issue").

Run:  streamlit run chat_analyzer_dashboard.py
Deps: pip install streamlit pandas openpyxl xlsxwriter
"""

import streamlit as st
import pandas as pd
import numpy as np
import re, io, warnings, gc, json, os
from datetime import datetime, timedelta

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG — Graas.ai theme
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Chat Analyzer Dashboard | Graas.ai",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM CSS — Graas.ai brand colours (#1B2A4A navy, #00C4B4 teal, #FF6B35 orange)
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Global ── */
html, body, [class*="css"] { font-family: 'Inter', 'Segoe UI', sans-serif; }
.main { background: #F4F6FB; }
.block-container { padding: 1.5rem 2rem; }

/* ── Top header bar ── */
.graas-header {
    background: linear-gradient(135deg, #1B2A4A 0%, #243554 100%);
    border-radius: 12px;
    padding: 1.2rem 1.8rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.graas-header h1 { color: #fff; margin: 0; font-size: 1.5rem; font-weight: 700; }
.graas-header p  { color: #A8C0D6; margin: 0; font-size: 0.85rem; }
.graas-logo { color: #00C4B4; font-size: 2rem; }

/* ── Metric cards ── */
.metric-row { display: flex; gap: 1rem; margin-bottom: 1.5rem; flex-wrap: wrap; }
.metric-card {
    background: #fff;
    border-radius: 10px;
    padding: 1rem 1.3rem;
    flex: 1;
    min-width: 150px;
    border-left: 4px solid #00C4B4;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.metric-card.orange { border-left-color: #FF6B35; }
.metric-card.red    { border-left-color: #E74C3C; }
.metric-card.navy   { border-left-color: #1B2A4A; }
.metric-card.green  { border-left-color: #27AE60; }
.metric-val { font-size: 1.9rem; font-weight: 800; color: #1B2A4A; }
.metric-label { font-size: 0.78rem; color: #7A8EA8; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px; }
.metric-sub { font-size: 0.75rem; color: #A0AEC0; margin-top: 2px; }

/* ── Section titles ── */
.section-title {
    font-size: 1rem;
    font-weight: 700;
    color: #1B2A4A;
    border-bottom: 2px solid #00C4B4;
    padding-bottom: 0.4rem;
    margin: 1.5rem 0 1rem;
}

/* ── Priority badges ── */
.badge-high   { background:#FDECEA; color:#C0392B; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-medium { background:#FEF9E7; color:#D68910; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }
.badge-low    { background:#EAF4FB; color:#2980B9; padding:2px 8px; border-radius:12px; font-size:0.75rem; font-weight:600; }

/* ── Sentiment ── */
.sent-pos { color:#27AE60; font-weight:600; }
.sent-neu { color:#7F8C8D; font-weight:600; }
.sent-neg { color:#C0392B; font-weight:600; }

/* ── Sidebar ── */
section[data-testid="stSidebar"] { background: #1B2A4A !important; }
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3 {
    color: #00C4B4 !important; font-size: 1rem !important; font-weight: 700 !important;
}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] .stMarkdown span { color: #FFFFFF !important; }
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stSelectbox > label,
section[data-testid="stSidebar"] .stMultiSelect > label,
section[data-testid="stSidebar"] .stDateInput > label,
section[data-testid="stSidebar"] .stTextInput > label {
    color: #FFFFFF !important;
    font-size: 0.85rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.3px !important;
}
section[data-testid="stSidebar"] .stSelectbox > div > div,
section[data-testid="stSidebar"] .stMultiSelect > div > div,
section[data-testid="stSidebar"] .stDateInput > div > div > input,
section[data-testid="stSidebar"] .stTextInput > div > div > input {
    background: #FFFFFF !important;
    color: #1B2A4A !important;
    border-radius: 6px !important;
    border: 1.5px solid #00C4B4 !important;
}
section[data-testid="stSidebar"] .stSelectbox svg,
section[data-testid="stSidebar"] .stMultiSelect svg { color: #1B2A4A !important; fill: #1B2A4A !important; }
section[data-testid="stSidebar"] .stMultiSelect span[data-baseweb="tag"] {
    background: #00C4B4 !important; color: #fff !important;
}
section[data-testid="stSidebar"] hr { border-color: #2E4A6A !important; }
section[data-testid="stSidebar"] strong { color: #00C4B4 !important; }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] { background: #fff; border-radius:8px; padding:4px; gap:4px; }
.stTabs [data-baseweb="tab"] { border-radius:6px; padding:6px 18px; font-weight:600; color:#7A8EA8; }
.stTabs [aria-selected="true"] { background:#00C4B4 !important; color:#fff !important; }

/* ── Suggested reply box ── */
.reply-box {
    background: #F0FBF9;
    border: 1px solid #00C4B4;
    border-radius: 8px;
    padding: 0.9rem 1rem;
    font-size: 0.85rem;
    color: #1B2A4A;
    line-height: 1.6;
    margin-top: 0.5rem;
}
.reply-label { font-size:0.75rem; color:#00C4B4; font-weight:700; text-transform:uppercase; margin-bottom:4px; }

/* ── TikTok platform card ── */
.tiktok-card { border-left-color: #010101 !important; }

/* ── Upload area ── */
.upload-area {
    background: #fff;
    border: 2px dashed #00C4B4;
    border-radius: 12px;
    padding: 2rem;
    text-align: center;
    margin-bottom: 1rem;
}

/* ── Store search result info box ── */
.store-search-info {
    background: #E8F8F6;
    border: 1px solid #00C4B4;
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 0.78rem;
    color: #1B2A4A;
    margin-top: 4px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

# ── Issue keyword dictionary (weight=1 per match) ─────────────────────────────
# CHANGE: expanded keyword lists based on real message samples observed in the
# uploaded chat data (JSON-unwrapped buyer messages).
#
# NEW (this version): merged in the full multilingual keyword reference sheet
# (English / Thai / Bahasa / Tagalog / Chinese) so that far fewer genuine
# buyer issues fall through to "Other". Every keyword below is matched
# case-insensitively as a substring of the extracted buyer text, and plural /
# tense variants (e.g. "cancel", "cancelled", "cancelling", "canceling") are
# listed explicitly to support singular/plural and common variations.
ISSUE_KEYWORDS = {
    "Refund": [
        "refund", "คืนเงิน", "pengembalian dana", "dana kembali", "ibalik", "irefund",
        "bayar balik", "money back", "reimburse", "reimbursement",
        "kembalikan uang", "uang kembali",
        "refunded", "refunds", "money return", "return money", "refund request",
        "bayaran balik", "pulangkan duit", "คืนเงินให้",
        "not refunded yet", "refund status", "refund pending", "waiting for refund",
        "belum refund", "belum dapat refund", "refund belum masuk",
        "dana belum masuk", "refund not received", "when will i get my refund",
        "refund delay", "yet to receive refund", "refund amount wrong",
        "partial refund", "how long does refund take", "เงินยังไม่เข้า",
        "ยังไม่ได้รับเงินคืน", "รอเงินคืน",
    ],
    "Return": [
        "return", "คืนสินค้า", "retur", "rma", "send back", "ส่งคืน", "kembalikan",
        "return item", "product return", "too big", "too small", "doesn't fit",
        "wrong fit", "size issue", "exchange", "change of size",
        "want to return", "how to return", "return policy", "returning", "returned",
        "return request", "return process", "pengembalian barang",
        "ส่งของคืน", "คืนของ", "ส่งคืนสินค้า",
        "how do i send back", "drop off return", "courier pickup",
        "collect the item", "schedule pickup", "return label", "return shipping",
        "print label", "restock fee", "unopened item", "unused item",
        "cara retur", "proses pengembalian", "ต้องคืนของยังไง",
        "ส่งของคืนยังไง", "นัดรับของคืน",
    ],
    "Cancellation": [
        "cancel", "cancelled", "ยกเลิก", "batalkan", "batal", "cancellation",
        "cancel order", "ยกเลิกคำสั่งซื้อ", "revoke", "withdraw", "i cancelled",
        "change address", "change size", "change color", "change mobile",
        # NEW — reference keyword sheet additions
        "want to cancel", "how to cancel", "canceling", "cancelling", "pembatalan",
        "batalkan pesanan", "cancel my order", "stop order", "cancel purchase",
        "ยกเลิกออเดอร์",
        # NEW — additional real-world phrasing
        "order cancelled by seller", "seller cancelled", "auto cancel",
        "system cancelled", "order got cancelled", "why was my order cancelled",
        "unable to cancel", "can't cancel", "no cancel button",
        "tidak bisa batal", "ไม่สามารถยกเลิก", "ยกเลิกไม่ได้",
    ],
    "Delay": [
        "delay", "late", "slow", "ช้า", "lambat", "belum sampai", "haven't received",
        "not arrived", "waiting", "รอนาน", "still waiting", "lama", "terlambat",
        "overdue", "not delivered yet", "ยังไม่ได้รับ", "belum diterima",
        # NEW — shipment-status keywords
        "ship", "delivery", "shipping", "status", "when will receive",
        "receive by today", "ship out", "when will my item be shipped",
        "has it been shipped", "shipped out", "shipped yet",
        "follow up", "update", "when will i receive", "track",
        # NEW — reference keyword sheet additions
        "where is my order", "where is my package", "where is my parcel",
        "when can i receive", "shipment status", "order status",
        "estimated delivery", "still not received", "have not received",
        "did not arrive", "no delivery", "not dispatched", "pending delivery",
        "pending shipment", "belum terima", "kapan sampai", "lama sekali",
        "makakarating", "makaabot", "kailan kaya", "delivery status",
        "order update", "ยังไม่มา", "ยังไม่ถึง", "สถานะพัสดุ",
        "เมื่อไหร่จะได้รับ", "รอสินค้า",
        # NEW — additional real-world phrasing
        "out for delivery", "in transit", "stuck at warehouse", "stuck in transit",
        "no movement", "not moving", "no update on tracking",
        "tracking not updating", "held at customs", "customs clearance",
        "returned to sender", "failed delivery attempt", "courier didn't come",
        "driver did not arrive", "not yet dispatch", "belum dikirim",
        "belum di proses", "รอจัดส่ง", "พัสดุค้าง", "พัสดุไม่ขยับ",
    ],
    "Damaged/Wrong Item": [
        "wrong item", "wrong product", "damaged", "broken", "defective",
        "สินค้าผิด", "ของเสีย", "ของแตก", "ของชำรุด", "rusak", "cacat",
        "salah barang", "salah produk", "not as described", "different item",
        "wrong size", "wrong colour", "wrong color", "different from picture",
        # NEW — product quality keywords
        "defected", "sole separation", "deviate", "damaged box", "damaged product",
        "dah pakai", "mcm dah pakai", "kotor", "looks used", "old stock",
        # NEW — reference keyword sheet additions
        "wrong variation", "wrong variant", "received wrong", "got wrong",
        "incorrect item", "incorrect product", "not what i ordered",
        "tama po ang variation", "make sure tama", "tampered", "torn", "dented",
        "cracked", "leaking", "expired", "kadaluarsa", "expire", "salah variasi",
        "หมดอายุ", "สินค้าไม่ตรง", "ผิดสินค้า", "สินค้าเสีย", "สินค้าแตก",
        # NEW — additional real-world phrasing
        "item arrived broken", "packaging damaged", "item spoiled", "moldy",
        "smell bad", "missing parts", "parts missing", "doesn't work properly",
        "malfunctioning", "stitching came off", "peeling", "discoloured",
        "discolored", "scratched", "dirty item", "rusty", "barang rosak",
        "barang cacat", "สินค้ามีตำหนิ", "สินค้าเป็นรอย",
    ],
    "Missing Item": [
        "missing", "not received", "didn't receive", "never received",
        "ไม่ได้รับ", "ของหาย", "hilang", "tidak diterima", "tidak ada", "kurang",
        "incomplete", "item missing", "package empty",
        # NEW — reference keyword sheet additions
        "empty box", "no item", "not inside", "not in the box", "missing product",
        "barang tidak ada", "barang kurang", "tidak lengkap", "never got",
        "ไม่ครบ", "ของไม่ครบ", "ขาด",
        # NEW — additional real-world phrasing
        "only received one item", "received less than ordered", "quantity wrong",
        "short shipped", "one item missing from order", "box was light",
        "seal broken", "kurang barang", "jumlah tidak sesuai", "ได้ของไม่ครบ",
        "ได้รับของไม่ครบตามจำนวน",
    ],
    "Payment Issue": [
        "payment", "ชำระเงิน", "bayar", "pembayaran", "charge", "double charge",
        "overcharged", "wrong charge", "billing", "invoice", "โอนเงิน", "จ่ายเงิน",
        "pay", "transfer", "deducted", "not paid", "receipt", "official receipt",
        # NEW — reference keyword sheet additions
        "payment failed", "payment error", "cannot pay", "can't pay",
        "charged twice", "double payment", "extra charge", "wrong amount",
        "payment issue", "payment problem", "billing issue", "bank transfer",
        "credit card", "gagal bayar", "pembayaran gagal", "tax invoice",
        "การชำระเงิน", "ชำระเงินไม่ได้", "ชำระไม่ได้", "จ่ายไม่ได้", "โอนไม่ได้",
        "ใบกำกับภาษี",
        # NEW — additional real-world phrasing
        "money deducted but no order", "payment deducted but order not placed",
        "refund to bank", "bank not received", "wrong bank account",
        "payment pending", "still processing payment", "cod problem",
        "cash on delivery issue", "belum dapat resit", "หักเงินแล้วแต่ไม่ได้ออเดอร์",
        "เงินหักแล้วแต่ไม่มีออเดอร์",
    ],
    "Product Inquiry": [
        "how to", "how do", "วิธีใช้", "ราคา", "price", "size", "ขนาด",
        "สี", "colour", "color", "spec", "specification", "ingredient",
        "cara pakai", "ukuran", "warna", "harga", "stok", "stock", "available",
        "variant", "model", "version",
        # NEW — restock & size recommendation keywords
        "when will be available", "product available", "size available",
        "color available", "restock", "uk", "us", "euro",
        "what size should i take", "waterproof", "will be fit", "warranty",
        # NEW — material/product questions from data samples
        "suitable for", "ผลิตจาก", "ประเทศอะไร", "tali", "budak", "umur",
        # NEW — reference keyword sheet additions
        "size chart", "uk size", "us size", "eu size", "same size",
        "actual picture", "actual photo", "is this suitable", "is it suitable",
        "do you have", "do you sell", "is it available", "still available",
        "how to use", "how to wear", "tersedia", "ada stok", "masih ada",
        "ada size", "apakah ada", "travelling", "expiry date", "shelf life",
        "manufacture date", "authentic", "genuine", "dimension", "dimensions",
        "recommend", "สอบถาม", "อยากรู้", "อยากทราบ", "ขอถาม", "ยังมี",
        "มีไหม", "มีขาย", "มีสต็อก", "สต็อค", "สินค้ามีไหม", "库存", "有库存",
        # NEW — additional real-world phrasing
        "compatible with", "what material", "material of", "washable",
        "machine wash", "care instructions", "how many pieces", "set of how many",
        "included in the box", "what's inside the box", "bahan apa",
        "cuci pakai apa", "ส่วนประกอบ", "ซักได้ไหม", "ผ้าอะไร",
    ],
    "Promotion Issue": [
        "voucher", "promo", "discount", "coupon", "code", "sale", "offer",
        "โปรโมชั่น", "ส่วนลด", "โค้ด", "diskon", "kode promo", "cashback",
        "flash sale", "deal", "bundle",
        # NEW — promotion keywords
        "live", "vouchers",
        # NEW — reference keyword sheet additions
        "voucher code", "promo code", "discount code", "coupon code",
        "cannot apply", "voucher not working", "code not working",
        "promo not valid", "discount not applied", "kode tidak bisa",
        "kupon tidak valid", "diskon tidak berlaku", "free gift", "free item",
        "freebie", "bundle deal", "buy 1 get 1", "b1g1", "free shipping",
        "reward", "rebate", "campaign", "ใช้โค้ดไม่ได้", "โค้ดใช้ไม่ได้",
        "ของแถม", "แถม", "กิจกรรม",
        # NEW — additional real-world phrasing
        "member discount", "first purchase discount", "app exclusive",
        "seller voucher", "shop voucher", "platform voucher", "minimum spend",
        "min spend not met", "voucher expired", "code expired", "kode kadaluarsa",
        "โค้ดหมดอายุ", "ส่วนลดหมดอายุ",
    ],
    "Technical Issue": [
        "error", "bug", "cannot", "can't", "unable", "failed", "not working",
        "app issue", "website", "login", "checkout problem", "system",
        "ไม่สามารถ", "เกิดข้อผิดพลาด", "tidak bisa", "gagal", "eror",
        # NEW — reference keyword sheet additions
        "technical issue", "technical problem", "page not found", "server error",
        "link not working", "button not working", "cannot access",
        "access denied", "aplikasi error", "tidak bisa login",
        "tidak bisa checkout", "glitch", "crash", "loading",
        "ระบบขัดข้อง", "ระบบล่ม", "เข้าไม่ได้",
        # NEW — additional real-world phrasing
        "app keeps crashing", "page keeps loading", "can't upload",
        "cannot upload", "can't checkout", "payment gateway error",
        "otp not received", "verification code not received",
        "tidak menerima otp", "ไม่ได้รับ otp", "ระบบค้าง",
    ],
    "Complaint": [
        "complain", "complaint", "terrible", "horrible", "awful", "worst",
        "ร้องเรียน", "ไม่พอใจ", "รำคาญ", "โกรธ", "disappointed",
        "frustrated", "unacceptable", "poor service", "bad service",
        "kecewa", "mengecewakan", "tidak puas", "buruk", "parah",
        # NEW — escalation language observed in real data
        "consumer court", "file a complaint", "report", "sue", "legal",
        # NEW — reference keyword sheet additions (incl. fraud/scam signals)
        "false advertising", "waste of time", "no response", "not responding",
        "no reply", "terrible experience", "bad experience", "worst experience",
        "pelayanan buruk", "tidak membantu", "sangat kecewa", "proper packaging",
        "request_proper_packaging", "unhappy", "angry", "upset", "dissatisfied",
        "rude", "unprofessional", "scam", "fraud", "cheat", "ignored",
        "แย่มาก", "บริการแย่",
        # NEW — additional real-world phrasing
        "will not buy again", "never shopping here again", "never buying again",
        "reporting to platform", "escalate to platform", "taking this further",
        "very unprofessional", "extremely disappointed", "not satisfied at all",
        "sangat tidak puas", "จะไม่ซื้ออีกแล้ว", "ไม่ซื้อแล้ว",
    ],
}

# ── Priority classification ───────────────────────────────────────────────────
# UPDATED: "Missing Item" moved from Medium → High per the refreshed priority
# rules. High/Medium/Low buckets otherwise keep their previous membership.
PRIORITY_MAP = {
    "High":   ["Refund", "Complaint", "Damaged/Wrong Item", "Missing Item"],
    "Medium": ["Delay", "Return", "Cancellation", "Payment Issue"],
    # "Other" (no keyword matched any defined Issue Type) is explicitly Low
    # priority — it should never be treated as urgent.
    "Low":    ["Product Inquiry", "Promotion Issue", "Technical Issue", "Other"],
}

# ── High-priority escalation keywords (NEW) ───────────────────────────────────
# These cover urgent/high-severity situations called out in the updated
# priority rules — Fraud, Double Charge, Account Hacked, Legal Threat,
# Chargeback, Safety Issue, and "order marked delivered but not received" —
# that can surface inside an otherwise Medium/Low-priority Issue Type (e.g. a
# double charge is still classified as "Payment Issue", but must still be
# escalated to High priority). `get_priority()` checks these BEFORE falling
# back to the standard PRIORITY_MAP lookup, so the underlying Issue Type
# classification, reply templates, and action steps are left untouched.
HIGH_PRIORITY_ESCALATION_KEYWORDS = [
    # Fraud / scam
    "fraud", "scam", "unauthorized transaction", "unauthorised transaction",
    "unauthorized charge", "unauthorised charge", "fake seller", "被骗", "诈骗",
    # Double charge
    "double charge", "charged twice", "double payment", "charged 2 times",
    "duplicate charge", "duplicate payment",
    # Account hacked / security
    "account hacked", "account was hacked", "hacked my account", "hacked",
    "unauthorized access", "unauthorised access", "someone accessed my account",
    "account compromised", "account breach", "แฮก", "โดนแฮก",
    # Legal threat / chargeback
    "legal action", "take legal action", "sue you", "lawyer", "consumer court",
    "file a police report", "police report", "chargeback", "dispute charge",
    "bank dispute", "credit card dispute", "report to authorities",
    "consumer protection", "ฟ้องร้อง", "ดำเนินคดี",
    # Safety issue
    "safety issue", "safety hazard", "injury", "injured", "got hurt",
    "allergic reaction", "caused injury", "fire hazard", "explode", "exploded",
    "electric shock", "chemical burn", "อันตราย", "บาดเจ็บ",
    # Order marked delivered but not received (false delivery status)
    "marked as delivered", "shows as delivered", "shows delivered",
    "delivered but not received", "delivered but i didn't receive",
    "delivered but i did not receive", "says delivered but", "false delivery",
    "gps shows delivered", "system shows delivered",
]

# ── Team Member → Store Code mapping (effective 30 March 2026) ────────────────
TEAM_ASSIGNMENTS = {
    "Yeria":      ["AACMH", "FFH", "IKU",
                   "GED MY", "GEDMY", "GED_MY"],
    "Syahira":    ["EWG", "HFC", "AAISS",
                   "GED SG", "GEDSG", "GED_SG"],
    "Keerthana":  ["AABIY", "AABIW", "AAFTP",
                   "GED PH", "GEDPH", "GED_PH"],
    "Alfian":     ["IGZ", "AADMJ", "AAEDD", "AADWP",
                   "IGZ ID", "IGZID", "IGZ_ID"],
    "Jaye":       ["GSK", "DBC", "IEI", "FYW", "ILL"],
    "Ratchakorn": ["AABWU", "AAFHU", "AAFHB"],
}

STORE_TO_AGENT = {
    store.upper(): agent
    for agent, stores in TEAM_ASSIGNMENTS.items()
    for store in stores
}

# GED is a shared store code — assign by COUNTRY_CODE
GED_COUNTRY_TO_AGENT = {
    "MY": "Yeria",
    "SG": "Syahira",
    "PH": "Keerthana",
}

AGENT_SHIFT = {
    "Yeria":      "GED MY · AACMH / FFH / IKU",
    "Syahira":    "GED SG · EWG / HFC / AAISS",
    "Keerthana":  "GED PH · AABIY / AABIW / AAFTP",
    "Alfian":     "IGZ ID · AADMJ / AAEDD / AADWP",
    "Jaye":       "GSK / DBC / IEI / FYW / ILL",
    "Ratchakorn": "Full-time · AABWU / AAFHU / AAFHB",
}

# ── Stalling patterns — seller has acknowledged but not resolved ──────────────
# CHANGE: expanded to cover patterns seen in TikTok/Shopee seller messages.
STALLING_PATTERNS = [
    r"will (check|look|get back|follow up|investigate|verify|review|update)",
    r"let me (check|look into|verify|confirm|see)",
    r"(checking|looking into|investigating|following up|reviewing)",
    r"please (wait|hold on|allow us|bear with)",
    r"i will (check|get back|follow up|update)",
    r"we (are|will) (checking|looking|investigating|getting back|following up)",
    r"get back to you",
    r"bear with us",
    r"kindly (wait|allow|hold)",
    r"we'?ll? (check|look|get back|follow up)",
    r"akan (kami|segera) (cek|periksa|tindak lanjut|proses|hubungi)",
    r"mohon (tunggu|ditunggu|bersabar)",
    r"kami (sedang|akan) (cek|periksa|proses|tindak lanjut)",
    r"จะตรวจสอบ", r"กำลังตรวจสอบ", r"จะแจ้งกลับ",
    r"จะดำเนินการ", r"ขอตรวจสอบ", r"ขอเวลา",
    r"จะติดต่อกลับ", r"ติดตามให้", r"กำลังประสานงาน",
    r"escalat",
    # NEW — common Malay stalling phrases seen in TikTok data
    r"boleh (tunggu|check|semak|tanya)",
    r"kami akan (semak|tindak|hubungi|proses)",
    r"sedang (semak|proses|check)",
    r"follow up",
    r"bole follow up",
]

# ── Resolution patterns — seller has given a concrete answer/action ───────────
# CHANGE: significantly expanded with patterns seen in real seller messages.
RESOLUTION_PATTERNS = [
    r"refund (has been|was|is) (processed|completed|done|issued|approved)",
    r"(your|the) (order|item|package) (has been|was|is) (shipped|dispatched|replaced|delivered)",
    r"(issue|problem|case) (has been|was|is) (resolved|fixed|closed|sorted|handled)",
    r"(cancellation|cancel) (has been|was|is) (processed|done|completed|approved)",
    r"(we have|we've) (processed|completed|resolved|fixed|issued|sent)",
    r"please (expect|allow) (\d|few|some|a couple)",
    r"track.*link.*sent", r"tracking (number|id|code) (is|was|has been)",
    r"you (should|will) (receive|get) (it|your order|the item)",
    r"ดำเนินการเรียบร้อย", r"จัดการเรียบร้อย", r"แก้ไขเรียบร้อย",
    r"คืนเงินเรียบร้อย", r"ยกเลิกเรียบร้อย",
    r"sudah (diproses|selesai|dikirim|dikembalikan|dibatalkan)",
    r"telah (diproses|selesai|diselesaikan|dikirimkan)",
    # NEW — resolution signals from real data samples
    r"ยินดีให้บริการ",           # "happy to serve" — polite closing
    r"ยินดีต้อนรับ",             # welcome / acknowledgement
    r"จัดส่งแล้ว",               # shipped already
    r"เรียบร้อยแล้ว",            # completed already
    r"ดำเนินการแล้ว",            # done/processed already
    r"สินค้าของทางเรา.*มีอายุ",  # product detail answered
    r"can cancel.*follow.*steps", # cancellation instructions given
    r"you can cancel your order", # direct resolution instruction
    r"i'?m? (keer|robbie|agent)", # agent identified themselves
    r"connected to (support|live chat)",
    r"(our|the) (agent|team) will (respond|reply|contact|assist)",
    r"sudah dijawab", r"sudah selesai",
    r"telah dikirim", r"telah diproses",
    r"telah diselesaikan",
    r"barang (telah|sudah) (dikirim|dihantar)",
    # NEW — TikTok-specific closings
    r"terima kasih.*kerana",
    r"maaf atas (kesulitan|masalah)",
    r"harap maklum",
]

# ── Auto-reply / bot patterns — exclude from stall/resolution detection ────────
# CHANGE: added chatbot JSON markers seen in Lazada/Shopee/TikTok data.
AUTO_REPLY_PATTERNS = [
    r"(thank you for contacting|thanks for reaching out).*auto",
    r"auto.?reply", r"automated (response|message|reply)",
    r"we'?ll? (get back|respond) (to you )?(within|in|shortly|soon)",
    r"our (team|agent).*(will|shall) (respond|reply|contact)",
    r"welcome to .*(official store|store).*\nhow (can|may) (we|i) help",
    r"สวัสดีค่ะ.*แอดมิน.*ยินดีให้บริการ",
    r"ยินดีต้อนรับ.*ร้าน",
    r"hi.{0,30}welcome to.{0,40}store",
    # NEW — chatbot signals observed in JSON ext fields
    r"generateby.*chatgpt",
    r"level2intentname",
    r"chatbot_replied",
    r"requestid.*issync",
    r"pass_through_data",
    r"shopee_chatbot",
]

POSITIVE_KWS = [
    "thank", "thanks", "great", "excellent", "awesome", "perfect", "love",
    "good", "nice", "happy", "satisfied", "wonderful", "amazing", "fantastic",
    "superb", "appreciate", "helpful", "fast", "quick", "well done", "recommend",
    "ขอบคุณ", "ดีมาก", "ประทับใจ", "พอใจ", "ยอดเยี่ยม", "ดีเลย", "ดีค่ะ", "ดีครับ",
    "terima kasih", "bagus", "mantap", "keren", "memuaskan", "puas", "oke baik",
    "salamat", "maganda", "ayos", "galing",
    "ok pls", "ok thanks",   # NEW — casual positive acknowledgements from data
]

NEGATIVE_KWS = [
    "terrible", "worst", "angry", "disappointed", "frustrated", "cheated", "scam",
    "fraud", "fake", "broken", "damaged", "wrong item", "missing", "never received",
    "unacceptable", "horrible", "awful", "complain", "complaint", "refund",
    "ผิดหวัง", "โกรธ", "ไม่พอใจ", "แย่มาก", "แย่", "หลอกลวง", "ของเสีย",
    "ของปลอม", "ช้ามาก", "รอนาน", "สินค้าไม่ตรง", "ไม่ได้รับ", "ชำรุด",
    "tipu", "rusak", "cacat", "mengecewakan", "marah", "kecewa", "buruk", "parah",
    "salah", "tidak diterima", "hilang",
    # NEW — from real data
    "consumer court", "file a complaint", "dah pakai", "mcm dah pakai",
    "kotor", "defected", "sole separation",
]

SUGGESTED_REPLIES = {
    "Refund": (
        "Thank you for reaching out, and we sincerely apologise for the inconvenience. "
        "We have reviewed your request and are pleased to confirm that your refund of [AMOUNT] "
        "has been initiated and will be reflected in your original payment method within 3–5 business days. "
        "Your order reference is [ORDER_ID]. We truly value your trust in us and hope to serve you better next time. "
        "If you have any further questions, please don't hesitate to reach out. 😊\n\n"
        "We'd love to hear your feedback — could you take a moment to rate your experience with us?"
    ),
    "Return": (
        "Thank you for contacting us about your return request. We're sorry to hear the product "
        "didn't meet your expectations. We've initiated the return process for order [ORDER_ID]. "
        "Please use the return label / return portal link we'll send to your registered email within 24 hours. "
        "Once we receive the item, the replacement or refund will be processed within 3–5 business days. "
        "We appreciate your patience and your continued support. 😊\n\n"
        "How would you rate your experience with us today?"
    ),
    "Cancellation": (
        "We've received your cancellation request for order [ORDER_ID]. We're sorry to see you go! "
        "Your order has been successfully cancelled and any payment made will be refunded within 3–5 business days. "
        "If you change your mind or need assistance with a future purchase, we're always here to help. 😊\n\n"
        "We'd appreciate your feedback — how was your experience with our team today?"
    ),
    "Delay": (
        "Thank you for your patience, and we sincerely apologise for the delay with your order [ORDER_ID]. "
        "We've checked with our logistics partner and your package is currently [STATUS]. "
        "Estimated delivery is [DATE]. We understand how frustrating delays can be and we truly appreciate your understanding. "
        "You can track your order in real time here: [TRACKING_LINK]. "
        "Please reach out if the delivery isn't received by [DATE+1] and we'll escalate immediately. 😊\n\n"
        "How was your experience with our support team today?"
    ),
    "Damaged/Wrong Item": (
        "We're truly sorry to hear that you received a damaged / incorrect item for order [ORDER_ID]. "
        "This is not the experience we want for you. To resolve this as quickly as possible, "
        "we've arranged a replacement to be dispatched within 1–2 business days. "
        "You do not need to return the incorrect / damaged item. "
        "We sincerely apologise for the inconvenience caused and will ensure this doesn't happen again. 😊\n\n"
        "Could you spare a moment to rate your support experience today?"
    ),
    "Missing Item": (
        "We're sorry to hear that your order [ORDER_ID] arrived with a missing item. "
        "We've raised an investigation with our fulfilment team and will have an update for you within 24 hours. "
        "In the meantime, we'll arrange a replacement or full refund, whichever you prefer. "
        "We apologise for this experience and truly appreciate your patience. 😊\n\n"
        "We'd love your feedback — how would you rate your experience with us today?"
    ),
    "Payment Issue": (
        "Thank you for flagging this payment concern. We've reviewed your account and order [ORDER_ID]. "
        "Our finance team has been notified and the discrepancy will be resolved within 2–3 business days. "
        "A confirmation will be sent to your registered email once completed. "
        "We apologise for any inconvenience and truly value your trust in us. 😊\n\n"
        "How was your experience with our support team today?"
    ),
    "Product Inquiry": (
        "Thank you for your interest in [PRODUCT_NAME]! "
        "Here are the details you requested: [DETAILS]. "
        "If you have more questions about specifications, sizing, or availability, "
        "please feel free to ask — we're happy to help you find the perfect product. 😊\n\n"
        "How can we assist you further today?"
    ),
    "Promotion Issue": (
        "Thank you for reaching out about the promotion. We're sorry for the confusion. "
        "We've reviewed your order [ORDER_ID] and confirmed that the discount of [AMOUNT] is applicable. "
        "The adjustment will be reflected within 24–48 hours. "
        "If the voucher code didn't apply correctly, please share it with us and we'll verify it right away. 😊\n\n"
        "How was your support experience today?"
    ),
    "Technical Issue": (
        "We apologise for the technical difficulty you're experiencing. "
        "Our team has been notified and is working on a resolution. "
        "In the meantime, please try [TROUBLESHOOTING STEP] and let us know if the issue persists. "
        "We aim to have this fully resolved within [TIMEFRAME]. "
        "Thank you for your patience — we appreciate it greatly. 😊\n\n"
        "How was your experience with our support today?"
    ),
    "Complaint": (
        "Thank you for taking the time to share your feedback, and we sincerely apologise for the experience you had. "
        "This is not the standard of service we strive for. We've escalated your case [CASE_ID] to our senior team "
        "for immediate review, and a dedicated agent will contact you within 4 hours. "
        "We take every concern seriously and are committed to making this right for you. 😊\n\n"
        "Your feedback helps us improve — how would you rate your support experience today?"
    ),
    "Other": (
        "Thank you for reaching out to us! We've reviewed your message and our team is addressing your concern. "
        "We aim to provide a resolution within 24 hours and will keep you updated throughout. "
        "We appreciate your patience and your trust in us. 😊\n\n"
        "How was your experience with our support team today?"
    ),
}

TEAM_START_DATE = pd.Timestamp("2026-03-30")

CONVERSION_KEYWORDS = [
    "i want to buy", "i'd like to buy", "i would like to buy", "how to buy",
    "how to order", "how do i order", "place an order", "can i order",
    "add to cart", "how to purchase", "i want to purchase", "proceed to checkout",
    "ready to buy", "i'll take it", "i want this", "i'll buy", "i want to get",
    "interested to buy", "interested in buying", "want to order",
    "อยากสั่ง", "สั่งซื้อ", "จะซื้อ", "ซื้อ", "สนใจซื้อ", "จะสั่ง",
    "mau beli", "mau order", "mau pesan", "ingin beli", "ingin order", "cara beli",
    "mag-order", "gusto kong bilhin", "bibilhin ko", "paano mag-order",
]

ACTION_STEPS = {
    "Refund": (
        "1. Verify order ID and payment method in Seller Centre.\n"
        "2. Check refund eligibility (within 15 days of purchase).\n"
        "3. Initiate refund via platform refund portal — select 'Approved by Seller'.\n"
        "4. Confirm refund amount matches original payment.\n"
        "5. Notify buyer with expected timeline (3–5 business days).\n"
        "6. Log in DKSH tracker under 'Refund Cases'."
    ),
    "Return": (
        "1. Verify product condition and return reason with buyer.\n"
        "2. Check return window (platform-specific: Lazada 7 days, Shopee 15 days).\n"
        "3. Approve return request in Seller Centre.\n"
        "4. Send return shipping label to buyer via platform chat.\n"
        "5. Once item received, inspect and process refund/replacement.\n"
        "6. Update DKSH tracker under 'Return Cases'."
    ),
    "Cancellation": (
        "1. Check order status — cancellable only before 'Ready to Ship'.\n"
        "2. Approve cancellation in Seller Centre if eligible.\n"
        "3. If already shipped, advise buyer to reject delivery.\n"
        "4. Refund will auto-process within 3–5 business days.\n"
        "5. Log in DKSH tracker under 'Cancellation Cases'."
    ),
    "Delay": (
        "1. Check logistics tracking in Seller Centre → Order Details.\n"
        "2. Contact logistics provider if package stalled > 3 days.\n"
        "3. Share tracking link with buyer immediately.\n"
        "4. If lost in transit, file a claim with logistics partner.\n"
        "5. Offer replacement or refund if delivery fails SLA.\n"
        "6. Escalate to platform CS if logistics provider unresponsive."
    ),
    "Damaged/Wrong Item": (
        "1. Request photo evidence from buyer (damaged/wrong item + packaging).\n"
        "2. Log dispute in Seller Centre under 'Return & Refund'.\n"
        "3. Approve replacement dispatch — do NOT ask buyer to return.\n"
        "4. Arrange courier pickup of damaged item (optional).\n"
        "5. Update DKSH tracker under 'Damaged/Wrong Item'.\n"
        "6. Report to warehouse for quality investigation."
    ),
    "Missing Item": (
        "1. Request unboxing video/photo from buyer as evidence.\n"
        "2. Check packing list vs order items in warehouse system.\n"
        "3. If confirmed missing, dispatch replacement within 24 hours.\n"
        "4. If uncertain, raise internal investigation with warehouse.\n"
        "5. Log in DKSH tracker under 'Missing Item'."
    ),
    "Payment Issue": (
        "1. Verify transaction details in platform payment dashboard.\n"
        "2. Check for double-charge or incorrect deduction.\n"
        "3. Raise dispute ticket with platform finance team.\n"
        "4. Provide buyer with case/ticket reference number.\n"
        "5. Follow up within 2 business days for resolution update."
    ),
    "Product Inquiry": (
        "1. Provide accurate product specs/details from official product sheet.\n"
        "2. If stock inquiry — check live inventory in Seller Centre.\n"
        "3. For sizing — share size guide image or chart.\n"
        "4. For availability — advise on restock ETA if applicable.\n"
        "5. Opportunity to upsell / cross-sell related products."
    ),
    "Promotion Issue": (
        "1. Verify voucher/promo code validity in Seller Centre → Promotions.\n"
        "2. Check eligibility criteria (min. spend, product category, date range).\n"
        "3. If code valid but not applied — advise buyer to re-checkout.\n"
        "4. If code expired — offer alternative discount if authorised.\n"
        "5. Escalate to marketing team for promo setup errors."
    ),
    "Technical Issue": (
        "1. Identify the platform and device buyer is using.\n"
        "2. Advise standard troubleshooting: clear cache, update app, reinstall.\n"
        "3. If platform-side issue — check platform status page.\n"
        "4. Raise support ticket with platform technical team.\n"
        "5. Keep buyer updated with ETA from platform team."
    ),
    "Complaint": (
        "1. Acknowledge and empathise — do NOT be defensive.\n"
        "2. Log complaint details in DKSH escalation tracker.\n"
        "3. Identify root cause (product/logistics/service failure).\n"
        "4. Offer concrete resolution: refund / replacement / discount.\n"
        "5. Escalate to senior manager if buyer threatens churn/review.\n"
        "6. Follow up within 4 hours with resolution update."
    ),
    "Other": (
        "1. Understand buyer's concern fully before responding.\n"
        "2. Route to appropriate team if issue is specialised.\n"
        "3. Aim to resolve within 24 hours.\n"
        "4. Log in DKSH tracker under 'General Enquiries'."
    ),
}


# ─────────────────────────────────────────────────────────────────────────────
# MESSAGE TEXT EXTRACTION  (NEW)
# ─────────────────────────────────────────────────────────────────────────────

def extract_plain_text(raw_msg: str) -> str:
    """
    Extract human-readable text from a platform message blob.

    Messages on Lazada, Shopee, and TikTok are JSON-encoded strings.
    Lazada uses {"txt": "..."} or {"ext": "{...}", "txt": "..."}.
    Shopee uses {"text": "...", "translation": {...}}.
    TikTok uses {"content": "..."} or {"order_id": "..."}.

    UPDATED: widened to also recognise "message", "body", "caption", "desc",
    "description", "msg", "summary", "content_text", "message_text" as text
    fields, and to recurse into nested containers ("ext", "data", "payload",
    "item", "extra", or a list of message blocks) so more real message text
    is captured. Previously, messages using any of these less-common shapes
    came back as an empty string and were then unavoidably classified as
    "Other" downstream — this reduces that failure mode without touching
    the classification logic itself.

    Returns a clean plain-text string (may be empty for genuinely
    media-only messages with no text anywhere in the blob).
    """
    if not isinstance(raw_msg, str) or not raw_msg.strip():
        return ""

    # Fast path: message is already plain text (no JSON braces)
    stripped = raw_msg.strip()
    if not stripped.startswith("{") and not stripped.startswith("["):
        return stripped

    try:
        obj = json.loads(stripped)
    except (json.JSONDecodeError, ValueError):
        # Partial / malformed JSON — try direct text extraction
        return _regex_extract_text(stripped)

    return _extract_from_obj(obj)


# Text-ish field names to look for, in priority order, across platforms.
_TEXT_FIELD_NAMES = (
    "txt", "text", "content", "message", "body", "caption",
    "desc", "description", "msg", "summary", "content_text", "message_text",
)

# Container field names that may hold a nested message structure.
_NESTED_CONTAINER_NAMES = ("ext", "data", "payload", "item", "extra", "detail")


def _extract_from_obj(obj, depth: int = 0, permissive: bool = False) -> str:
    """
    Recursively search a parsed JSON value (dict / list / str) for the first
    usable piece of message text.

    At the top level (permissive=False) only _TEXT_FIELD_NAMES /
    _NESTED_CONTAINER_NAMES are considered, so we never accidentally grab
    unrelated metadata (order_id, user_id, timestamps, etc.) as "message text".

    Once we've unwrapped into something already known to be a text-carrying
    value — e.g. a Lazada `txt` field that is itself a JSON object of
    translations like {"th": "...", "en": "..."}, or a known nested
    container (`ext`/`data`/`payload`/...) — we switch to permissive=True,
    which additionally falls back to the first non-empty string found among
    ANY of that dict's values (matching the previous, simpler behavior for
    these already-scoped nested objects).

    Depth-limited to avoid runaway recursion on pathological/self-referential
    structures.
    """
    if depth > 4:
        return ""

    if isinstance(obj, str):
        return obj.strip()

    if isinstance(obj, list):
        for item in obj:
            found = _extract_from_obj(item, depth + 1, permissive)
            if found:
                return found
        return ""

    if isinstance(obj, dict):
        # ── Direct text fields first ──────────────────────────────────────────
        for key in _TEXT_FIELD_NAMES:
            val = obj.get(key)
            if isinstance(val, str) and val.strip():
                # Value may itself be JSON (e.g. Lazada txt = {"th": "...", "en": "..."})
                if val.strip().startswith("{") or val.strip().startswith("["):
                    try:
                        inner = _extract_from_obj(json.loads(val), depth + 1, permissive=True)
                    except (json.JSONDecodeError, ValueError):
                        inner = ""
                    if inner:
                        return inner
                    continue
                return val.strip()
            if isinstance(val, (dict, list)):
                found = _extract_from_obj(val, depth + 1, permissive)
                if found:
                    return found

        # ── Nested containers (may be stringified JSON or a raw object) ──────
        for key in _NESTED_CONTAINER_NAMES:
            val = obj.get(key)
            if isinstance(val, str) and val.strip().startswith("{"):
                try:
                    found = _extract_from_obj(json.loads(val), depth + 1, permissive=True)
                    if found:
                        return found
                except (json.JSONDecodeError, ValueError):
                    pass
            elif isinstance(val, (dict, list)):
                found = _extract_from_obj(val, depth + 1, permissive=True)
                if found:
                    return found

        # ── Permissive fallback: this dict is already known to be a scoped
        # text-carrying value (e.g. a translation map {"en": ..., "th": ...}),
        # so grab the first non-empty string value regardless of key name. ──
        if permissive:
            for val in obj.values():
                if isinstance(val, str) and val.strip():
                    return val.strip()
                if isinstance(val, (dict, list)):
                    found = _extract_from_obj(val, depth + 1, permissive=True)
                    if found:
                        return found

        return ""

    return ""


def _regex_extract_text(s: str) -> str:
    """Fallback: extract quoted text values from a JSON-like string via regex."""
    # Match values for common text-ish keys (same set as _TEXT_FIELD_NAMES)
    m = re.search(
        r'"(?:txt|text|content|message|body|caption|desc|description|msg|'
        r'summary|content_text|message_text)"\s*:\s*"([^"]{2,})"',
        s, re.IGNORECASE,
    )
    if m:
        return m.group(1).replace("\\n", " ").replace('\\"', '"').strip()
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────

def detect_sentiment(text: str) -> str:
    if not isinstance(text, str) or not text.strip():
        return "Neutral"
    t = text.lower()
    neg = sum(1 for kw in NEGATIVE_KWS if kw in t)
    pos = sum(1 for kw in POSITIVE_KWS if kw in t)
    if neg > pos:
        return "Negative"
    if pos > neg:
        return "Positive"
    return "Neutral"


def detect_issue_type(text: str) -> str:
    """
    Classify buyer message text into an issue type.

    LOGIC:
    1. ISSUE_KEYWORDS is checked; every matching keyword contributes weight=1
       to that issue type's cumulative score.
    2. The issue type with the highest cumulative score wins.
    3. On ties, priority order (High > Medium > Low) breaks the tie.
    4. Returns "Other" if no keywords matched.

    NOTE: ISSUE_KEYWORDS carries a substantially expanded, multilingual
    keyword set (see constants section) covering many more real-world phrasings
    per issue type. This function's matching logic is unchanged — it still does
    a case-insensitive substring match — but with more keywords available,
    far fewer genuine buyer issues will fail to match any keyword and fall
    through to "Other". "Other" is only returned when truly no keyword from
    any issue type is found in the buyer's message.

    Parameters
    ----------
    text : str
        Pre-extracted plain text from buyer messages (not raw JSON).
    """
    if not isinstance(text, str) or not text.strip():
        return "Other"
    t = text.lower()

    scores: dict[str, float] = {}

    # ── Match against ISSUE_KEYWORDS (weight=1 per matching keyword) ─────────
    for issue, kws in ISSUE_KEYWORDS.items():
        for kw in kws:
            if kw.lower() in t:
                scores[issue] = scores.get(issue, 0) + 1

    if not scores:
        return "Other"

    # ── Break ties by priority (High wins) ────────────────────────────────────
    priority_order = {"High": 0, "Medium": 1, "Low": 2}
    max_score = max(scores.values())
    tied = [iss for iss, sc in scores.items() if sc == max_score]
    if len(tied) == 1:
        return tied[0]

    # Among tied, pick the one with highest (most urgent) priority
    def _priority_rank(iss):
        for p, issues in PRIORITY_MAP.items():
            if iss in issues:
                return priority_order.get(p, 3)
        return 3

    return min(tied, key=_priority_rank)


def _has_escalation_keyword(text: str) -> bool:
    """
    NEW: Returns True if the buyer's text contains any High-priority
    escalation keyword (fraud, account hacked, legal/chargeback threats,
    double charge, safety issue, false "delivered" status, etc.).

    This is deliberately independent of Issue Type detection — it only
    looks at raw keyword presence in the buyer's message so it can override
    priority even when the Issue Type itself resolves to something that
    would normally be Medium/Low priority (e.g. "Payment Issue").
    """
    if not isinstance(text, str) or not text.strip():
        return False
    t = text.lower()
    return any(kw in t for kw in HIGH_PRIORITY_ESCALATION_KEYWORDS)


def get_priority(issue_type: str, text: str | None = None) -> str:
    """
    Return the priority ("High" / "Medium" / "Low") for a given issue type.

    UPDATED: now accepts an optional `text` parameter (the buyer's message
    text for the conversation). If any High-priority escalation keyword
    (fraud, account hacked, legal threat/chargeback, double charge, safety
    issue, false "delivered" status, etc. — see HIGH_PRIORITY_ESCALATION_KEYWORDS)
    is present, the conversation is escalated to "High" priority regardless
    of its Issue Type. This lets urgent situations that happen to live inside
    a Medium/Low-priority Issue Type (e.g. a double-charge inside
    "Payment Issue") still get flagged as High priority.

    `text` is optional and defaults to None, so every existing call site that
    calls `get_priority(issue_type)` without a text argument continues to
    work exactly as before (base PRIORITY_MAP lookup only).

    FIX: escalation is intentionally skipped when issue_type == "Other". No
    Issue Type keyword matched in that case, so we can't be confident the
    escalation phrase is meaningfully connected to a real, classifiable
    issue rather than incidental wording. This also prevents "Other" from
    ever splitting into two separate priority buckets (Other/Low and
    Other/High) in breakdown tables — "Other" is always Low.
    """
    if issue_type != "Other" and text is not None and _has_escalation_keyword(text):
        return "High"
    for priority, issues in PRIORITY_MAP.items():
        if issue_type in issues:
            return priority
    return "Low"


def matches_any(text: str, patterns: list) -> bool:
    if not isinstance(text, str):
        return False
    t = text.lower()
    return any(re.search(p, t, re.IGNORECASE) for p in patterns)


def is_auto_reply(text: str) -> bool:
    return matches_any(text, AUTO_REPLY_PATTERNS)


def conversation_is_unresolved(seller_msgs_plain: list) -> bool:
    """
    Determine whether a conversation is unresolved.

    IMPROVED LOGIC:
    - Receives pre-parsed plain-text seller messages (not raw JSON).
    - Auto-replies and bot messages are excluded from both stall and
      resolution detection to avoid false positives.
    - Tracks a running stall_found flag. A resolution pattern resets it
      to False. If stall_found is True at the end → unresolved.
    - Conversations with zero seller messages are considered unresolved
      (buyer has not received any response).

    Parameters
    ----------
    seller_msgs_plain : list[str]
        List of plain-text seller messages in chronological order.
    """
    if not seller_msgs_plain:
        # No seller messages at all — buyer has not been responded to
        return True

    stall_found = False
    human_msg_seen = False  # at least one real (non-bot) seller message

    for msg in seller_msgs_plain:
        if not isinstance(msg, str) or not msg.strip():
            continue

        # Skip bot / auto-reply messages — they don't count for resolution
        if is_auto_reply(msg):
            continue

        human_msg_seen = True

        if matches_any(msg, RESOLUTION_PATTERNS):
            stall_found = False   # resolution clears the stall flag
        elif matches_any(msg, STALLING_PATTERNS):
            stall_found = True    # stall raised (until resolved)

    # If only bot messages exist with no human response → unresolved
    if not human_msg_seen:
        return True

    return stall_found


def compute_csat(sentiment: str, is_resolved: bool) -> float:
    matrix = {
        ("Positive", True):  5.0,
        ("Positive", False): 3.5,
        ("Neutral",  True):  4.0,
        ("Neutral",  False): 3.0,
        ("Negative", True):  2.5,
        ("Negative", False): 1.0,
    }
    return matrix.get((sentiment, is_resolved), 3.0)


def generate_summary(buyer_msgs: list, issue_type: str) -> str:
    if not buyer_msgs:
        return "No buyer messages."
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)])[:400]
    return f"[{issue_type}] Buyer enquiry: {combined[:200]}{'...' if len(combined) > 200 else ''}"


def fmt_mins(mins) -> str:
    if pd.isna(mins) or mins < 0:
        return "—"
    if mins < 60:
        return f"{int(mins)}m"
    h = int(mins // 60)
    m = int(mins % 60)
    return f"{h}h {m}m" if m else f"{h}h"


def get_team_member(store_code: str, country_code: str = "") -> str:
    code = str(store_code).strip().upper()
    if not code:
        return "Others"
    # GED is a shared store code — route by country
    if code == "GED":
        country = str(country_code).strip().upper()
        return GED_COUNTRY_TO_AGENT.get(country, "Others")
    return STORE_TO_AGENT.get(code, "Others")


def detect_conversion(buyer_msgs: list) -> bool:
    combined = " ".join([m for m in buyer_msgs if isinstance(m, str)]).lower()
    return any(kw.lower() in combined for kw in CONVERSION_KEYWORDS)


def get_action_steps(issue_type: str) -> str:
    return ACTION_STEPS.get(issue_type, ACTION_STEPS["Other"])


def filter_stores_by_search(all_stores: list, search_term: str) -> list:
    if not search_term or not search_term.strip():
        return all_stores
    term = search_term.strip().upper()
    return [s for s in all_stores if term in str(s).upper()]


def compute_wow_mom(conv_df: pd.DataFrame) -> tuple:
    df = conv_df.copy()
    df = df[df["LAST_MSG_TIME"].notna()].copy()
    if df.empty:
        return pd.DataFrame(), pd.DataFrame()

    df["WEEK"]  = df["LAST_MSG_TIME"].dt.to_period("W").apply(lambda r: r.start_time)
    df["MONTH"] = df["LAST_MSG_TIME"].dt.to_period("M").apply(lambda r: r.start_time)

    def agg_metrics(df_in, period_col):
        agg = (
            df_in.groupby(period_col)
            .agg(
                Conversations=("CONVERSATION_ID", "count"),
                Resolved=("IS_RESOLVED", "sum"),
                Unresolved=("IS_UNRESOLVED", "sum"),
                Avg_CSAT=("CSAT_PROXY", "mean"),
                Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
                Negative=("SENTIMENT", lambda x: (x == "Negative").sum()),
                Positive=("SENTIMENT", lambda x: (x == "Positive").sum()),
                Conversions=("IS_CONVERSION", "sum"),
                TikTok=("PLATFORM", lambda x: (x == "TikTok").sum()),
                Shopee=("PLATFORM", lambda x: (x == "Shopee").sum()),
                Lazada=("PLATFORM", lambda x: (x == "Lazada").sum()),
            )
            .reset_index()
            .sort_values(period_col)
        )
        agg["CRR_%"] = (agg["Resolved"] / agg["Conversations"] * 100).round(1)
        agg["Avg_CSAT"] = agg["Avg_CSAT"].round(2)
        agg["Avg_CRT_mins"] = agg["Avg_CRT_mins"].round(1)
        for col in ["Conversations", "Avg_CSAT", "CRR_%", "Avg_CRT_mins", "Conversions"]:
            agg[f"Δ {col}"] = agg[col].diff().round(2)
        return agg

    wow = agg_metrics(df, "WEEK")
    mom = agg_metrics(df, "MONTH")
    return wow, mom


def compute_team_performance(conv_df: pd.DataFrame) -> pd.DataFrame:
    df = conv_df.copy()
    df = df[df["LAST_MSG_TIME"] >= TEAM_START_DATE].copy()
    if df.empty or "TEAM_MEMBER" not in df.columns:
        return pd.DataFrame()

    perf = (
        df.groupby("TEAM_MEMBER")
        .agg(
            Conversations=("CONVERSATION_ID", "count"),
            Resolved=("IS_RESOLVED", "sum"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
            Positive_Sent=("SENTIMENT", lambda x: (x == "Positive").sum()),
            Negative_Sent=("SENTIMENT", lambda x: (x == "Negative").sum()),
            Conversions=("IS_CONVERSION", "sum"),
            High_Priority=("PRIORITY", lambda x: (x == "High").sum()),
        )
        .reset_index()
    )
    perf["CRR_%"]        = (perf["Resolved"] / perf["Conversations"] * 100).round(1)
    perf["Avg_CSAT"]     = perf["Avg_CSAT"].round(2)
    perf["Avg_CRT_mins"] = perf["Avg_CRT_mins"].round(1)
    perf["Shift"]        = perf["TEAM_MEMBER"].map(AGENT_SHIFT).fillna("Day")
    perf = perf.sort_values("Conversations", ascending=False).reset_index(drop=True)
    return perf


# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────

def _repair_streaming_xlsx(file_bytes: bytes) -> bytes:
    """
    Repair a streaming XLSX (Google Sheets / some exporters) that is missing
    the End-of-Central-Directory record AND required OOXML metadata files.
    Returns original bytes unchanged if already a valid ZIP.
    """
    import zipfile as _zf, struct as _st, zlib as _zl, re as _re

    try:
        with _zf.ZipFile(io.BytesIO(file_bytes)):
            pass
        return file_bytes
    except _zf.BadZipFile:
        pass

    data = bytearray(file_bytes)
    extracted = {}

    pos = 0
    while pos < len(data) - 30:
        idx = data.find(b"PK\x03\x04", pos)
        if idx == -1:
            break
        try:
            (_, _, flag, method, _, _, _,
             comp_size, uncomp_size,
             fname_len, extra_len) = _st.unpack_from("<4sHHHHHIIIHH", data, idx)
            fname = data[idx+30: idx+30+fname_len].decode("utf-8", errors="replace")
            data_start = idx + 30 + fname_len + extra_len
            raw = bytes(data[data_start:])

            if method == 8:
                d = _zl.decompressobj(-15)
                content = b""
                i = 0
                while i < len(raw):
                    try:
                        content += d.decompress(raw[i:i+65536])
                        i += 65536
                    except _zl.error:
                        break
            elif method == 0:
                content = raw[:uncomp_size] if uncomp_size > 0 else b""
            else:
                pos = idx + 4
                continue

            if fname and content:
                if fname == "xl/sharedStrings.xml" and b"</sst>" not in content:
                    last_si = content.rfind(b"</si>")
                    if last_si != -1:
                        content = content[:last_si + 5] + b"\n</sst>"
                extracted[fname] = content
        except Exception:
            pass
        pos = idx + 4

    if not extracted:
        return file_bytes

    sheet_keys = sorted(
        [k for k in extracted if k.startswith("xl/worksheets/sheet") and k.endswith(".xml")],
        key=lambda x: int(_re.search(r"sheet(\d+)", x).group(1)) if _re.search(r"sheet(\d+)", x) else 0
    )
    n_sheets = len(sheet_keys)

    if "[Content_Types].xml" not in extracted:
        sheet_overrides = "\n".join(
            f'  <Override PartName="/xl/worksheets/sheet{i+1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(n_sheets)
        )
        extracted["[Content_Types].xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
            '  <Default Extension="xml" ContentType="application/xml"/>\n'
            '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
            '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>\n'
            '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>\n'
            + sheet_overrides + "\n"
            "</Types>"
        ).encode("utf-8")

    if "_rels/.rels" not in extracted:
        extracted["_rels/.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n'
            "</Relationships>"
        ).encode("utf-8")

    if "xl/workbook.xml" not in extracted:
        sheet_elems = "\n".join(
            f'    <sheet name="Sheet{i+1}" sheetId="{i+1}" r:id="rId{i+1}"/>'
            for i in range(n_sheets)
        )
        extracted["xl/workbook.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"\n'
            '  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
            '  <sheets>\n'
            + sheet_elems + "\n"
            "  </sheets>\n"
            "</workbook>"
        ).encode("utf-8")

    if "xl/_rels/workbook.xml.rels" not in extracted:
        rels = []
        for i in range(n_sheets):
            rels.append(
                f'  <Relationship Id="rId{i+1}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{i+1}.xml"/>'
            )
        if "xl/sharedStrings.xml" in extracted:
            rels.append(
                f'  <Relationship Id="rId{n_sheets+1}" '
                f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                f'Target="sharedStrings.xml"/>'
            )
        extracted["xl/_rels/workbook.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            + "\n".join(rels) + "\n"
            "</Relationships>"
        ).encode("utf-8")

    if "xl/styles.xml" not in extracted:
        extracted["xl/styles.xml"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
            '  <fonts><font><sz val="11"/><name val="Calibri"/></font></fonts>\n'
            '  <fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>\n'
            '  <borders><border><left/><right/><top/><bottom/><diagonal/></border></borders>\n'
            '  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'
            '  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>\n'
            "</styleSheet>"
        ).encode("utf-8")

    new_zip = io.BytesIO()
    with _zf.ZipFile(new_zip, "w", compression=_zf.ZIP_DEFLATED) as zout:
        for priority in ["[Content_Types].xml", "_rels/.rels",
                         "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
                         "xl/styles.xml", "xl/sharedStrings.xml"]:
            if priority in extracted:
                zout.writestr(priority, extracted[priority])
        for fname, content in extracted.items():
            if fname not in ["[Content_Types].xml", "_rels/.rels",
                              "xl/workbook.xml", "xl/_rels/workbook.xml.rels",
                              "xl/styles.xml", "xl/sharedStrings.xml"]:
                zout.writestr(fname, content)

    new_zip.seek(0)
    return new_zip.read()


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, _file_hash: str = "") -> pd.DataFrame:
    """
    Load and normalise the raw message-level DataFrame from an XLSX file.

    CHANGE: Added MESSAGE_PARSED column population using extract_plain_text()
    so that downstream analysis works on clean text rather than raw JSON.
    """
    file_bytes = _repair_streaming_xlsx(file_bytes)

    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets_found = xl.sheet_names

    dfs = []
    for s in sheets_found:
        name_lower = s.lower()
        if "lazada" in name_lower:
            platform = "Lazada"
        elif "shopee" in name_lower:
            platform = "Shopee"
        elif "tiktok" in name_lower or "tik_tok" in name_lower or "tik tok" in name_lower:
            platform = "TikTok"
        else:
            platform = "Unknown"

        df = xl.parse(s, dtype=str)

        if platform == "Unknown" and "SITE_NICK_NAME_ID" in df.columns:
            sample_sites = df["SITE_NICK_NAME_ID"].dropna().astype(str).str.lower()
            if sample_sites.str.contains("shopee").any():
                platform = "Shopee"
            elif sample_sites.str.contains("lazada").any():
                platform = "Lazada"
            elif sample_sites.str.contains("tiktok|tik_tok").any():
                platform = "TikTok"

        df["PLATFORM"] = platform

        # ── Normalise column name variants ────────────────────────────────────
        if "MESSAGE_TIME" not in df.columns:
            if "MSG_TIME_RAW_TS" in df.columns:
                df = df.rename(columns={"MSG_TIME_RAW_TS": "MESSAGE_TIME"})
            elif "MSG_TIME_RAW" in df.columns:
                df = df.rename(columns={"MSG_TIME_RAW": "MESSAGE_TIME"})

        if "MESSAGE_PARSED" not in df.columns:
            if "MESSAGE" in df.columns:
                df = df.rename(columns={"MESSAGE": "MESSAGE_PARSED"})
            elif "CONTENT" in df.columns:
                df = df.rename(columns={"CONTENT": "MESSAGE_PARSED"})
            elif "MSG_CONTENT" in df.columns:
                df = df.rename(columns={"MSG_CONTENT": "MESSAGE_PARSED"})

        if "CONVERSATION_ID" not in df.columns:
            if "CHAT_ID" in df.columns:
                df = df.rename(columns={"CHAT_ID": "CONVERSATION_ID"})
            elif "SESSION_ID" in df.columns:
                df = df.rename(columns={"SESSION_ID": "CONVERSATION_ID"})

        if "STORE_CODE" not in df.columns:
            if "SHOP_CODE" in df.columns:
                df = df.rename(columns={"SHOP_CODE": "STORE_CODE"})
            elif "SELLER_ID" in df.columns:
                df = df.rename(columns={"SELLER_ID": "STORE_CODE"})

        if "BUYER_NAME" not in df.columns:
            if "USER" in df.columns:
                df = df.rename(columns={"USER": "BUYER_NAME"})
            elif "CUSTOMER_NAME" in df.columns:
                df = df.rename(columns={"CUSTOMER_NAME": "BUYER_NAME"})

        if "SENDER" not in df.columns:
            if "ROLE" in df.columns:
                df = df.rename(columns={"ROLE": "SENDER"})
            elif "FROM" in df.columns:
                df = df.rename(columns={"FROM": "SENDER"})

        if "MESSAGE_TIME" not in df.columns:
            continue

        # ── NEW: Extract plain text from JSON-encoded message blobs ───────────
        if "MESSAGE_PARSED" in df.columns:
            df["MESSAGE_PARSED"] = df["MESSAGE_PARSED"].apply(extract_plain_text)

        dfs.append(df)

    if not dfs:
        st.error("No valid chat sheets found in the uploaded file.")
        return pd.DataFrame()

    combined = pd.concat(dfs, ignore_index=True)

    dedup_cols = [c for c in ["CONVERSATION_ID", "MESSAGE_TIME", "SENDER", "MESSAGE_PARSED"] if c in combined.columns]
    combined = combined.drop_duplicates(subset=dedup_cols, keep="first")

    combined["MESSAGE_TIME"] = pd.to_datetime(combined["MESSAGE_TIME"], errors="coerce")
    combined = combined[combined["MESSAGE_TIME"].notna()].copy()

    for col in ["STORE_CODE", "SITE_NICK_NAME_ID", "CHANNEL_NAME", "COUNTRY_CODE",
                "CONVERSATION_ID", "BUYER_NAME", "MESSAGE_PARSED",
                "MESSAGE_TYPE", "SENDER"]:
        if col in combined.columns:
            combined[col] = combined[col].fillna("").astype(str).str.strip()

    for flag in ["IS_READ", "IS_ANSWERED"]:
        if flag in combined.columns:
            combined[flag] = (
                combined[flag].astype(str).str.strip().str.lower()
                .isin(["true", "1", "yes"])
            )

    return combined


# ─────────────────────────────────────────────────────────────────────────────
# ANALYSIS ENGINE
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_data(show_spinner=False, max_entries=1)
def analyse(df: pd.DataFrame) -> pd.DataFrame:
    """
    Conversation-level analysis engine.

    CHANGE: Passes pre-parsed plain-text messages to detect_issue_type()
    and conversation_is_unresolved() instead of raw JSON blobs.

    NEW: `get_priority()` is now called with the conversation's buyer text
    as well as its Issue Type, so High-priority escalation keywords (fraud,
    account hacked, legal threat/chargeback, double charge, safety issue,
    false "delivered" status, etc.) can bump PRIORITY to "High" even when
    the Issue Type itself would normally resolve to Medium/Low. Issue Type
    detection itself is unchanged.
    """
    df = df.copy()
    df["_sender_lower"] = df["SENDER"].str.lower().fillna("")
    df_sorted = df.sort_values(["CONVERSATION_ID", "MESSAGE_TIME"])

    buyer_mask  = df_sorted["_sender_lower"] == "buyer"
    seller_mask = df_sorted["_sender_lower"] == "seller"

    # MESSAGE_PARSED is now plain text (extracted in load_data)
    buyer_text_per_conv = (
        df_sorted[buyer_mask]
        .groupby("CONVERSATION_ID")["MESSAGE_PARSED"]
        .apply(lambda msgs: " ".join(m for m in msgs if isinstance(m, str) and m.strip()))
    )

    issue_map     = buyer_text_per_conv.apply(detect_issue_type)
    sentiment_map = buyer_text_per_conv.apply(detect_sentiment)

    meta_cols = ["PLATFORM", "STORE_CODE", "SITE_NICK_NAME_ID", "CHANNEL_NAME",
                 "COUNTRY_CODE", "BUYER_NAME", "BUYER_ID", "IS_ANSWERED", "IS_READ"]
    meta_cols = [c for c in meta_cols if c in df_sorted.columns]
    meta_df = df_sorted.groupby("CONVERSATION_ID")[meta_cols].first()

    time_df = df_sorted.groupby("CONVERSATION_ID")["MESSAGE_TIME"].agg(
        FIRST_MSG_TIME="min", LAST_MSG_TIME="max"
    )

    total_msgs        = df_sorted.groupby("CONVERSATION_ID").size().rename("MSG_COUNT")
    buyer_msgs_count  = df_sorted[buyer_mask].groupby("CONVERSATION_ID").size().rename("BUYER_MSG_COUNT")
    seller_msgs_count = df_sorted[seller_mask].groupby("CONVERSATION_ID").size().rename("SELLER_MSG_COUNT")

    # Seller messages as plain text lists (already extracted in load_data)
    seller_msgs_per_conv = (
        df_sorted[seller_mask]
        .groupby("CONVERSATION_ID")["MESSAGE_PARSED"]
        .apply(list)
    )
    buyer_msgs_per_conv = (
        df_sorted[buyer_mask]
        .groupby("CONVERSATION_ID")["MESSAGE_PARSED"]
        .apply(list)
    )

    rows = []
    for conv_id, grp in df_sorted.groupby("CONVERSATION_ID", sort=False):
        issue_type = issue_map.get(conv_id, "Other")
        sentiment  = sentiment_map.get(conv_id, "Neutral")
        b_msgs     = buyer_msgs_per_conv.get(conv_id, [])
        s_msgs     = seller_msgs_per_conv.get(conv_id, [])
        meta       = meta_df.loc[conv_id] if conv_id in meta_df.index else {}

        # CHANGE: seller messages are now plain text — resolution detection works correctly
        is_unresolved = conversation_is_unresolved(s_msgs)
        is_resolved   = not is_unresolved
        buyer_text    = buyer_text_per_conv.get(conv_id, "")
        # NEW: pass buyer_text so High-priority escalation keywords can
        # bump PRIORITY to "High" regardless of the detected Issue Type.
        priority      = get_priority(issue_type, buyer_text)
        csat          = compute_csat(sentiment, is_resolved)

        crt_list = []
        last_buyer_time = None
        for sender, msg_time in zip(grp["_sender_lower"].tolist(), grp["MESSAGE_TIME"].tolist()):
            if sender == "buyer":
                last_buyer_time = msg_time
            elif sender == "seller" and last_buyer_time is not None:
                delta = (msg_time - last_buyer_time).total_seconds() / 60
                if 0 <= delta <= 1440:
                    crt_list.append(delta)
                last_buyer_time = None
        avg_crt = float(np.mean(crt_list)) if crt_list else np.nan

        def _get(field, default=""):
            try:
                return meta[field] if hasattr(meta, "__getitem__") else getattr(meta, field, default)
            except Exception:
                return default

        rows.append({
            "CONVERSATION_ID":   conv_id,
            "PLATFORM":          _get("PLATFORM"),
            "STORE_CODE":        _get("STORE_CODE"),
            "SITE_NICK_NAME_ID": _get("SITE_NICK_NAME_ID"),
            "CHANNEL_NAME":      _get("CHANNEL_NAME"),
            "COUNTRY_CODE":      _get("COUNTRY_CODE"),
            "BUYER_NAME":        _get("BUYER_NAME"),
            "BUYER_ID":          _get("BUYER_ID"),
            "FIRST_MSG_TIME":    time_df.loc[conv_id, "FIRST_MSG_TIME"] if conv_id in time_df.index else pd.NaT,
            "LAST_MSG_TIME":     time_df.loc[conv_id, "LAST_MSG_TIME"]  if conv_id in time_df.index else pd.NaT,
            "MSG_COUNT":         int(total_msgs.get(conv_id, 0)),
            "BUYER_MSG_COUNT":   int(buyer_msgs_count.get(conv_id, 0)),
            "SELLER_MSG_COUNT":  int(seller_msgs_count.get(conv_id, 0)),
            "ISSUE_TYPE":        issue_type,
            "PRIORITY":          priority,
            "SENTIMENT":         sentiment,
            "IS_UNRESOLVED":     is_unresolved,
            "IS_RESOLVED":       is_resolved,
            "CSAT_PROXY":        round(csat, 1),
            "AVG_CRT_MINS":      round(avg_crt, 1) if not np.isnan(avg_crt) else None,
            "BUYER_SUMMARY":     generate_summary(b_msgs, issue_type),
            "IS_CONVERSION":     detect_conversion(b_msgs),
            "TEAM_MEMBER":       get_team_member(_get("STORE_CODE"), _get("COUNTRY_CODE")),
            "IS_ANSWERED":       str(_get("IS_ANSWERED")).lower() == "true",
            "IS_READ":           str(_get("IS_READ")).lower() == "true",
        })

    result = pd.DataFrame(rows)

    for col in ["PLATFORM", "ISSUE_TYPE", "PRIORITY", "SENTIMENT",
                "STORE_CODE", "CHANNEL_NAME", "COUNTRY_CODE", "TEAM_MEMBER", "SITE_NICK_NAME_ID"]:
        if col in result.columns:
            result[col] = result[col].astype("category")

    for col in ["BUYER_SUMMARY"]:
        if col in result.columns:
            result[col] = result[col].str[:300]

    gc.collect()
    return result


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(conv_df: pd.DataFrame, today_str: str) -> bytes:
    df = conv_df.copy()
    df["SUGGESTED_REPLY"] = df["ISSUE_TYPE"].astype(str).map(
        lambda it: SUGGESTED_REPLIES.get(it, SUGGESTED_REPLIES["Other"])
    )
    df["ACTION_STEPS"] = df["ISSUE_TYPE"].astype(str).map(get_action_steps)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        hdr_fmt  = wb.add_format({"bold": True, "bg_color": "#1B2A4A", "font_color": "#FFFFFF",
                                   "border": 1, "font_size": 10, "align": "center", "valign": "vcenter"})
        sub_fmt  = wb.add_format({"bold": True, "bg_color": "#00C4B4", "font_color": "#FFFFFF",
                                   "border": 1, "font_size": 10})
        num_fmt  = wb.add_format({"num_format": "#,##0", "border": 1})
        dec_fmt  = wb.add_format({"num_format": "0.0", "border": 1})
        cell_fmt = wb.add_format({"border": 1, "font_size": 9, "text_wrap": True, "valign": "top"})

        # ── Sheet 1: Summary Dashboard ────────────────────────────────────────
        ws1 = wb.add_worksheet("Summary Dashboard")
        writer.sheets["Summary Dashboard"] = ws1
        ws1.set_column(0, 0, 28)
        ws1.set_column(1, 1, 20)

        total      = len(df)
        resolved   = df["IS_RESOLVED"].sum()
        unresolved = df["IS_UNRESOLVED"].sum()
        crr        = round(resolved / total * 100, 1) if total else 0
        avg_crt    = df["AVG_CRT_MINS"].mean()
        avg_csat   = df["CSAT_PROXY"].mean()
        data_max_ts = df["LAST_MSG_TIME"].dropna().max()
        data_max_date = data_max_ts.normalize() if pd.notna(data_max_ts) else pd.Timestamp(today_str)
        today_df   = df[df["LAST_MSG_TIME"].dt.normalize() == data_max_date]
        hi_today   = len(today_df[today_df["PRIORITY"] == "High"])

        ws1.write(0, 0, f"Chat Analyzer Summary — {today_str}", wb.add_format(
            {"bold": True, "font_size": 14, "font_color": "#1B2A4A"}))
        ws1.write(1, 0, "Generated by Graas.ai Chat Analyzer Dashboard", wb.add_format(
            {"italic": True, "font_color": "#7A8EA8"}))

        summary_data = [
            ["Total Conversations", total],
            ["Today's Conversations", len(today_df)],
            ["Resolved Conversations", int(resolved)],
            ["Unresolved Conversations", int(unresolved)],
            ["Chat Resolution Rate (CRR)", f"{crr}%"],
            ["Avg Chat Response Time (CRT)", fmt_mins(avg_crt)],
            ["Avg CSAT Proxy Score (1–5)", round(avg_csat, 2) if not np.isnan(avg_csat) else "—"],
            ["Today's High Priority Chats", hi_today],
            ["Platforms", ", ".join(sorted(df["PLATFORM"].astype(str).unique().tolist()))],
            ["Shopee Conversations", int((df["PLATFORM"] == "Shopee").sum())],
            ["Lazada Conversations", int((df["PLATFORM"] == "Lazada").sum())],
            ["TikTok Conversations", int((df["PLATFORM"] == "TikTok").sum())],
        ]
        for i, (label, val) in enumerate(summary_data, start=3):
            ws1.write(i, 0, label, sub_fmt)
            ws1.write(i, 1, val, cell_fmt)

        ws1.write(13, 0, "ISSUE TYPE BREAKDOWN", sub_fmt)
        ws1.write(13, 1, "COUNT", hdr_fmt)
        for i, (issue, cnt) in enumerate(df["ISSUE_TYPE"].value_counts().items(), start=14):
            ws1.write(i, 0, issue, cell_fmt)
            ws1.write(i, 1, int(cnt), num_fmt)

        # ── Sheet 2: Today Priority Chats ─────────────────────────────────────
        priority_cols = [c for c in [
            "CONVERSATION_ID", "PLATFORM", "STORE_CODE", "CHANNEL_NAME",
            "SITE_NICK_NAME_ID", "COUNTRY_CODE", "TEAM_MEMBER", "BUYER_NAME",
            "ISSUE_TYPE", "PRIORITY", "SENTIMENT", "IS_UNRESOLVED",
            "CSAT_PROXY", "AVG_CRT_MINS", "IS_CONVERSION", "BUYER_SUMMARY", "SUGGESTED_REPLY",
        ] if c in df.columns]
        today_pri = today_df.sort_values(
            "PRIORITY", key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )[priority_cols]
        today_pri.to_excel(writer, sheet_name="Today Priority Chats", index=False)
        ws2 = writer.sheets["Today Priority Chats"]
        ws2.set_column(0, 0, 40); ws2.set_column(1, 5, 15); ws2.set_column(10, 13, 50)
        for c_idx, col in enumerate(today_pri.columns):
            ws2.write(0, c_idx, col, hdr_fmt)

        # ── Sheet 3: Detailed Chat Analysis ───────────────────────────────────
        detail_cols = [c for c in [
            "CONVERSATION_ID", "PLATFORM", "STORE_CODE", "CHANNEL_NAME",
            "SITE_NICK_NAME_ID", "COUNTRY_CODE", "TEAM_MEMBER", "BUYER_NAME",
            "FIRST_MSG_TIME", "LAST_MSG_TIME",
            "MSG_COUNT", "ISSUE_TYPE", "PRIORITY", "SENTIMENT",
            "IS_RESOLVED", "IS_UNRESOLVED", "CSAT_PROXY", "AVG_CRT_MINS",
            "IS_CONVERSION", "BUYER_SUMMARY", "SUGGESTED_REPLY",
        ] if c in df.columns]
        detail = df[detail_cols].copy()
        detail["FIRST_MSG_TIME"] = detail["FIRST_MSG_TIME"].dt.strftime("%Y-%m-%d %H:%M")
        detail["LAST_MSG_TIME"]  = detail["LAST_MSG_TIME"].dt.strftime("%Y-%m-%d %H:%M")
        detail.to_excel(writer, sheet_name="Detailed Chat Analysis", index=False)
        ws3 = writer.sheets["Detailed Chat Analysis"]
        ws3.set_column(0, 0, 40); ws3.set_column(7, 8, 18); ws3.set_column(17, 19, 60)
        for c_idx, col in enumerate(detail.columns):
            ws3.write(0, c_idx, col, hdr_fmt)

        # ── Sheet 4: Unresolved Chats ─────────────────────────────────────────
        unres = df[df["IS_UNRESOLVED"]][priority_cols].sort_values(
            "PRIORITY", key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        unres.to_excel(writer, sheet_name="Unresolved Chats", index=False)
        ws4 = writer.sheets["Unresolved Chats"]
        ws4.set_column(0, 0, 40); ws4.set_column(10, 13, 50)
        for c_idx, col in enumerate(unres.columns):
            ws4.write(0, c_idx, col, hdr_fmt)

    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# UI COMPONENTS
# ─────────────────────────────────────────────────────────────────────────────

def render_header():
    st.markdown("""
    <div class="graas-header">
        <div class="graas-logo">📊</div>
        <div>
            <h1>Chat Analyzer Dashboard</h1>
            <p>Graas.ai · Shopee, Lazada & TikTok · Sentiment · CSAT · Unresolved Detection · Suggested Replies</p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_metrics(conv_df: pd.DataFrame, today_ts: pd.Timestamp):
    total      = len(conv_df)
    resolved   = int(conv_df["IS_RESOLVED"].sum())
    unresolved = int(conv_df["IS_UNRESOLVED"].sum())
    crr        = round(resolved / total * 100, 1) if total else 0
    avg_crt    = conv_df["AVG_CRT_MINS"].mean()
    avg_csat   = conv_df["CSAT_PROXY"].mean()
    today_conv = conv_df[conv_df["LAST_MSG_TIME"].dt.normalize() == today_ts]
    hi_today   = len(today_conv[today_conv["PRIORITY"] == "High"])
    neg_pct    = round(len(conv_df[conv_df["SENTIMENT"] == "Negative"]) / total * 100, 1) if total else 0

    cols = st.columns(8)
    metrics = [
        (cols[0], "🗣️ Total Convs", f"{total:,}",   "all platforms", ""),
        (cols[1], "📅 Today",       f"{len(today_conv):,}", "conversations", "navy"),
        (cols[2], "✅ Resolved",    f"{resolved:,}", f"CRR {crr}%", "green"),
        (cols[3], "🔴 Unresolved",  f"{unresolved:,}", "need action", "red"),
        (cols[4], "⚡ CRT",         fmt_mins(avg_crt), "avg response time", "orange"),
        (cols[5], "⭐ CSAT",        f"{avg_csat:.1f}/5" if not np.isnan(avg_csat) else "—", "proxy score", ""),
        (cols[6], "😠 Negative",    f"{neg_pct}%",  "sentiment", "red"),
        (cols[7], "🔥 High Pri",    f"{hi_today}",  "today's urgent", "orange"),
    ]
    for col, label, val, sub, cls in metrics:
        with col:
            st.markdown(f"""
            <div class="metric-card {cls}">
                <div class="metric-label">{label}</div>
                <div class="metric-val">{val}</div>
                <div class="metric-sub">{sub}</div>
            </div>
            """, unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────────────────────────────────────

def apply_filters(conv_df: pd.DataFrame, today_ts: pd.Timestamp, data_end=None) -> pd.DataFrame:
    src = conv_df

    st.sidebar.markdown("## 🔍 Filters")
    st.sidebar.markdown("---")

    platforms = ["All"] + sorted(src["PLATFORM"].dropna().unique().tolist())
    sel_platform = st.sidebar.selectbox("🌐 Platform", platforms)

    _ts_min = src["LAST_MSG_TIME"].dropna().min()
    _ts_max = src["LAST_MSG_TIME"].dropna().max()
    min_date = _ts_min.date() if pd.notna(_ts_min) else datetime.today().date()
    max_date = _ts_max.date() if pd.notna(_ts_max) else datetime.today().date()
    default_start = min_date

    date_range = st.sidebar.date_input(
        "📅 Date Range",
        value=(default_start, max_date),
        min_value=min_date,
        max_value=max_date,
        help="Defaults to full data range in the file. Narrow to focus on a specific period.",
    )

    sel_prio  = st.sidebar.selectbox("🔴 Priority", ["All", "High", "Medium", "Low"])
    sel_sent  = st.sidebar.selectbox("😊 Sentiment", ["All", "Positive", "Neutral", "Negative"])
    sel_res   = st.sidebar.selectbox("✅ Resolution Status", ["All", "Resolved", "Unresolved"])
    issue_opts = ["All"] + sorted(src["ISSUE_TYPE"].dropna().unique().tolist())
    sel_issue = st.sidebar.selectbox("🏷️ Issue Type", issue_opts)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔎 Search & Filter")

    if "TEAM_MEMBER" in src.columns:
        all_agents = sorted(src["TEAM_MEMBER"].dropna().unique().tolist())
        sel_agents = st.sidebar.multiselect("👤 Team Member", all_agents)
    else:
        sel_agents = []

    all_stores = sorted(src["STORE_CODE"].dropna().unique().tolist())
    store_search = st.sidebar.text_input(
        "🔍 Search Store Code",
        placeholder="Type GED, IGZ, EWG…",
        help="Partial match — type any part of the store code.",
    )
    filtered_store_opts = filter_stores_by_search(all_stores, store_search)

    if store_search and filtered_store_opts:
        st.sidebar.markdown(
            f'<div class="store-search-info">🔍 {len(filtered_store_opts)} store(s) match "<b>{store_search}</b>"</div>',
            unsafe_allow_html=True,
        )
    elif store_search and not filtered_store_opts:
        st.sidebar.warning(f'No stores match "{store_search}"')

    sel_stores = st.sidebar.multiselect(
        "🏪 Store Code (select from results)",
        options=filtered_store_opts,
        default=filtered_store_opts if store_search and filtered_store_opts else [],
    )

    all_countries = sorted(src["COUNTRY_CODE"].dropna().unique().tolist())
    sel_countries = st.sidebar.multiselect("🌍 Country", all_countries)

    if "CHANNEL_NAME" in src.columns:
        all_channels = sorted(src["CHANNEL_NAME"].dropna().replace("", pd.NA).dropna().unique().tolist())
        sel_channels = st.sidebar.multiselect("📡 Channel Name", all_channels)
    else:
        sel_channels = []

    buyer_search = st.sidebar.text_input("🔍 Buyer Name")
    conv_search  = st.sidebar.text_input("🔍 Conversation ID")

    # ── Apply filters ─────────────────────────────────────────────────────────
    result = src.copy()

    if sel_platform != "All":
        result = result[result["PLATFORM"] == sel_platform]

    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_ts = pd.Timestamp(date_range[0])
        end_ts   = pd.Timestamp(date_range[1]) + pd.Timedelta(hours=23, minutes=59, seconds=59)
        result = result[
            (result["LAST_MSG_TIME"] >= start_ts) &
            (result["LAST_MSG_TIME"] <= end_ts)
        ]

    if sel_prio != "All":
        result = result[result["PRIORITY"] == sel_prio]

    if sel_sent != "All":
        result = result[result["SENTIMENT"] == sel_sent]

    if sel_res == "Resolved":
        result = result[result["IS_RESOLVED"]]
    elif sel_res == "Unresolved":
        result = result[result["IS_UNRESOLVED"]]

    if sel_issue != "All":
        result = result[result["ISSUE_TYPE"] == sel_issue]

    if sel_agents:
        result = result[result["TEAM_MEMBER"].isin(sel_agents)]

    if sel_stores:
        result = result[result["STORE_CODE"].isin(sel_stores)]

    if sel_countries:
        result = result[result["COUNTRY_CODE"].isin(sel_countries)]

    if sel_channels and "CHANNEL_NAME" in result.columns:
        result = result[result["CHANNEL_NAME"].isin(sel_channels)]

    if buyer_search:
        result = result[result["BUYER_NAME"].str.contains(buyer_search, case=False, na=False)]

    if conv_search:
        result = result[result["CONVERSATION_ID"].str.contains(conv_search, case=False, na=False)]

    st.sidebar.markdown("---")
    total = len(result)
    st.sidebar.markdown(f"**{total:,}** of **{len(src):,}** conversations")
    if total == 0:
        st.sidebar.warning("No results — try widening the date range or clearing filters.")
    return result


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────

def main():
    render_header()

    st.markdown('<div class="section-title">📂 Upload Chat Data</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Upload Excel file with sheets: lazada_chat_enquiries, shopee_chat_enquiries & tiktok_chat_enquiries",
        type=["xlsx"],
        help="Single Excel file containing Lazada, Shopee, and/or TikTok chat sheets.",
    )

    if not uploaded:
        st.info("👆 Upload your chat enquiries Excel file to get started.")
        st.markdown("""
        **Expected Excel format:**
        - Sheet 1: `lazada_chat_enquiries`
        - Sheet 2: `shopee_chat_enquiries`
        - Sheet 3: `tiktok_chat_enquiries` *(optional)*
        - Columns: `STORE_CODE`, `SITE_NICK_NAME_ID`, `COUNTRY_CODE`, `CONVERSATION_ID`,
          `IS_READ`, `IS_ANSWERED`, `MESSAGE_TIME`, `BUYER_NAME`, `MESSAGE_PARSED` (or `MESSAGE`),
          `MESSAGE_TYPE`, `MESSAGE_ID`, `SENDER`, `BUYER_ID`
        """)
        return

    import hashlib
    file_bytes = uploaded.read()
    file_hash  = hashlib.md5(file_bytes).hexdigest()

    with st.spinner("⏳ Loading chat data…"):
        raw_df = load_data(file_bytes, file_hash)

    _max_ts = raw_df["MESSAGE_TIME"].dropna().max()
    _min_ts = raw_df["MESSAGE_TIME"].dropna().min()

    today_date  = datetime.today().date()
    today_ts    = pd.Timestamp(today_date)
    today_str   = today_date.strftime("%Y-%m-%d")
    data_end    = _max_ts.date() if pd.notna(_max_ts) else today_date
    data_start  = _min_ts.date() if pd.notna(_min_ts) else today_date

    _platforms_found = sorted(raw_df["PLATFORM"].unique().tolist())
    _plat_str = " · ".join(_platforms_found)
    st.success(
        f"✅ Loaded **{len(raw_df):,}** messages · "
        f"**{raw_df['CONVERSATION_ID'].nunique():,}** conversations · "
        f"Platforms: **{_plat_str}** · "
        f"Data range: **{data_start}** → **{data_end}**"
    )

    with st.spinner("🔍 Analysing conversations — this runs once and is cached…"):
        conv_df = analyse(raw_df)
    del raw_df; gc.collect()

    conv_filtered = apply_filters(conv_df, today_ts, data_end)

    if conv_filtered.empty:
        st.warning("No conversations match the current filters.")
        return

    st.markdown('<div class="section-title">📈 Key Metrics</div>', unsafe_allow_html=True)
    render_metrics(conv_filtered, today_ts)

    st.markdown('<div class="section-title">📊 Analytics</div>', unsafe_allow_html=True)

    # ── Platform summary cards ──────────────────────────────────────────────
    platform_order  = ["Shopee", "Lazada", "TikTok"]
    all_platforms   = [p for p in platform_order if p in conv_filtered["PLATFORM"].values]
    other_platforms = [p for p in conv_filtered["PLATFORM"].unique() if p not in platform_order]
    all_platforms  += other_platforms

    if len(all_platforms) > 1:
        pcols = st.columns(len(all_platforms))
        plat_colors = {"Shopee": "#EE4D2D", "Lazada": "#0F146D", "TikTok": "#010101"}
        for i, plat in enumerate(all_platforms):
            plat_df   = conv_filtered[conv_filtered["PLATFORM"] == plat]
            plat_crr  = round(plat_df["IS_RESOLVED"].sum() / len(plat_df) * 100, 1) if len(plat_df) else 0
            plat_csat_raw = plat_df["CSAT_PROXY"].mean()
            plat_csat_str = f"{plat_csat_raw:.1f}" if not pd.isna(plat_csat_raw) else "—"
            bg_color  = plat_colors.get(plat, "#1B2A4A")
            with pcols[i]:
                st.markdown(f"""
                <div style="background:{bg_color};border-radius:10px;padding:12px 14px;color:white;margin-bottom:8px;">
                  <div style="font-size:13px;font-weight:700;letter-spacing:0.5px;opacity:0.85;">{plat.upper()}</div>
                  <div style="font-size:22px;font-weight:800;margin:4px 0;">{len(plat_df):,}</div>
                  <div style="font-size:11px;opacity:0.75;">conversations</div>
                  <div style="display:flex;gap:12px;margin-top:8px;font-size:12px;">
                    <span>CRR {plat_crr}%</span>
                    <span>CSAT {plat_csat_str}</span>
                  </div>
                </div>
                """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        issue_counts = conv_filtered["ISSUE_TYPE"].value_counts().reset_index()
        issue_counts.columns = ["Issue Type", "Count"]
        st.markdown("**Issue Type Distribution**")
        st.bar_chart(issue_counts.set_index("Issue Type")["Count"], color="#00C4B4")

    with c2:
        sent_counts = conv_filtered["SENTIMENT"].value_counts().reset_index()
        sent_counts.columns = ["Sentiment", "Count"]
        st.markdown("**Sentiment Breakdown**")
        st.bar_chart(sent_counts.set_index("Sentiment")["Count"])

    with c3:
        daily = (
            conv_filtered
            .assign(DATE=conv_filtered["LAST_MSG_TIME"].dt.normalize())
            .groupby("DATE")
            .size()
            .reset_index(name="Conversations")
        )
        st.markdown("**Daily Conversation Volume**")
        st.line_chart(daily.set_index("DATE")["Conversations"], color="#FF6B35")

    with c4:
        if len(all_platforms) > 1:
            plat_daily = (
                conv_filtered
                .assign(DATE=conv_filtered["LAST_MSG_TIME"].dt.normalize())
                .groupby(["DATE", "PLATFORM"])
                .size()
                .reset_index(name="Count")
                .pivot(index="DATE", columns="PLATFORM", values="Count")
                .fillna(0)
            )
            st.markdown("**Volume by Platform**")
            st.line_chart(plat_daily)
        else:
            plat_counts = conv_filtered["PLATFORM"].value_counts().reset_index()
            plat_counts.columns = ["Platform", "Count"]
            st.markdown("**Platform Breakdown**")
            st.bar_chart(plat_counts.set_index("Platform")["Count"], color="#1B2A4A")

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🔥 Today's Priority Chats",
        "📋 All Conversations",
        "🔴 Unresolved Chats",
        "💬 Suggested Replies",
        "📈 WoW / MoM Performance",
        "👥 Team Performance",
    ])

    display_cols = [c for c in [
        "CONVERSATION_ID", "PLATFORM", "STORE_CODE", "CHANNEL_NAME",
        "SITE_NICK_NAME_ID", "COUNTRY_CODE", "BUYER_NAME",
        "ISSUE_TYPE", "PRIORITY", "SENTIMENT", "IS_UNRESOLVED",
        "CSAT_PROXY", "AVG_CRT_MINS", "BUYER_SUMMARY",
    ] if c in conv_filtered.columns]

    col_config_base = {
        "CSAT_PROXY":    st.column_config.NumberColumn("CSAT (1-5)", format="%.1f"),
        "AVG_CRT_MINS":  st.column_config.NumberColumn("CRT (mins)", format="%.0f"),
        "IS_UNRESOLVED": st.column_config.CheckboxColumn("Unresolved?"),
        "BUYER_SUMMARY": st.column_config.TextColumn("Summary", width="large"),
    }

    with tab1:
        latest_date = pd.Timestamp(data_end)
        today_df = conv_filtered[conv_filtered["LAST_MSG_TIME"].dt.normalize() == latest_date]
        today_sorted = today_df.sort_values(
            "PRIORITY", key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        st.markdown(f"**{len(today_sorted)} conversations today** — sorted by priority")
        if today_sorted.empty:
            st.info("No conversations found for today.")
        else:
            st.dataframe(today_sorted[display_cols].reset_index(drop=True),
                         use_container_width=True, height=450, column_config=col_config_base)

    with tab2:
        all_sorted = conv_filtered.sort_values("LAST_MSG_TIME", ascending=False)
        st.markdown(f"**{len(all_sorted)} conversations** in filtered view")
        st.dataframe(all_sorted[display_cols].reset_index(drop=True),
                     use_container_width=True, height=500, column_config=col_config_base)

    with tab3:
        unres_df = conv_filtered[conv_filtered["IS_UNRESOLVED"]].sort_values(
            "PRIORITY", key=lambda s: s.map({"High": 0, "Medium": 1, "Low": 2}).fillna(3)
        )
        st.markdown(f"**{len(unres_df)} unresolved conversations** — stalling phrases without resolution")
        if unres_df.empty:
            st.success("🎉 No unresolved conversations found!")
        else:
            st.dataframe(unres_df[display_cols].reset_index(drop=True),
                         use_container_width=True, height=450, column_config=col_config_base)

    with tab4:
        st.markdown("### 💬 Suggested Reply Templates by Issue Type")
        st.caption("Empathetic, resolution-oriented replies — replace [PLACEHOLDERS] before sending.")
        for issue_type, reply_text in SUGGESTED_REPLIES.items():
            if issue_type == "Other":
                continue
            priority = get_priority(issue_type)
            badge_color = {"High": "🔴", "Medium": "🟡", "Low": "🔵"}.get(priority, "⚪")
            with st.expander(f"{badge_color} {issue_type}  ({priority} Priority)"):
                st.markdown(f"""
                <div class="reply-label">Suggested Reply</div>
                <div class="reply-box">{reply_text}</div>
                """, unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### 🔍 Look Up Reply for a Specific Conversation")
        conv_ids = conv_filtered["CONVERSATION_ID"].tolist()
        if conv_ids:
            sel_conv = st.selectbox("Select Conversation ID", conv_ids[:500])
            row = conv_filtered[conv_filtered["CONVERSATION_ID"] == sel_conv].iloc[0]
            st.markdown(f"""
            **Issue Type:** {row['ISSUE_TYPE']}  |
            **Priority:** {row['PRIORITY']}  |
            **Sentiment:** {row['SENTIMENT']}  |
            **CSAT Proxy:** {row['CSAT_PROXY']}
            """)
            st.markdown(f"""
            <div class="reply-label">Buyer Summary</div>
            <div class="reply-box">{row['BUYER_SUMMARY']}</div>
            """, unsafe_allow_html=True)
            suggested = SUGGESTED_REPLIES.get(str(row['ISSUE_TYPE']), SUGGESTED_REPLIES["Other"])
            st.markdown(f"""
            <div class="reply-label">Suggested Reply</div>
            <div class="reply-box">{suggested}</div>
            """, unsafe_allow_html=True)

    with tab5:
        st.markdown("### 📈 Week-on-Week & Month-on-Month Performance")
        wow_df, mom_df = compute_wow_mom(conv_filtered)
        wow_tab, mom_tab = st.tabs(["📅 Week-on-Week", "🗓️ Month-on-Month"])

        with wow_tab:
            if wow_df.empty:
                st.info("Not enough data for weekly comparison.")
            else:
                st.markdown("**Weekly Conversation Trend**")
                st.bar_chart(wow_df.set_index("WEEK")[["Conversations"]], color="#00C4B4")
                disp_wow = wow_df.copy()
                disp_wow["WEEK"] = disp_wow["WEEK"].dt.strftime("%d %b %Y")
                disp_wow["Avg_CRT_mins"] = disp_wow["Avg_CRT_mins"].apply(
                    lambda x: fmt_mins(x) if pd.notna(x) else "—")
                st.dataframe(
                    disp_wow[[c for c in ["WEEK","Conversations","Shopee","Lazada","TikTok",
                               "CRR_%","Avg_CSAT","Avg_CRT_mins","Conversions",
                               "Δ Conversations","Δ CRR_%","Δ Avg_CSAT"] if c in disp_wow.columns]].reset_index(drop=True),
                    use_container_width=True,
                    column_config={
                        "WEEK":            st.column_config.TextColumn("Week Starting"),
                        "Shopee":          st.column_config.NumberColumn("Shopee", format="%d"),
                        "Lazada":          st.column_config.NumberColumn("Lazada", format="%d"),
                        "TikTok":          st.column_config.NumberColumn("TikTok", format="%d"),
                        "CRR_%":           st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                        "Avg_CSAT":        st.column_config.NumberColumn("CSAT", format="%.2f"),
                        "Conversions":     st.column_config.NumberColumn("Conversions"),
                        "Δ Conversations": st.column_config.NumberColumn("Δ Conv", format="%+.0f"),
                        "Δ CRR_%":         st.column_config.NumberColumn("Δ CRR%", format="%+.1f"),
                        "Δ Avg_CSAT":      st.column_config.NumberColumn("Δ CSAT", format="%+.2f"),
                    },
                )

        with mom_tab:
            if mom_df.empty:
                st.info("Not enough data for monthly comparison.")
            else:
                st.markdown("**Monthly Conversation Trend**")
                st.bar_chart(mom_df.set_index("MONTH")[["Conversations"]], color="#FF6B35")
                disp_mom = mom_df.copy()
                disp_mom["MONTH"] = disp_mom["MONTH"].dt.strftime("%b %Y")
                disp_mom["Avg_CRT_mins"] = disp_mom["Avg_CRT_mins"].apply(
                    lambda x: fmt_mins(x) if pd.notna(x) else "—")
                st.dataframe(
                    disp_mom[[c for c in ["MONTH","Conversations","Shopee","Lazada","TikTok",
                               "CRR_%","Avg_CSAT","Avg_CRT_mins","Conversions",
                               "Δ Conversations","Δ CRR_%","Δ Avg_CSAT"] if c in disp_mom.columns]].reset_index(drop=True),
                    use_container_width=True,
                    column_config={
                        "MONTH":           st.column_config.TextColumn("Month"),
                        "Shopee":          st.column_config.NumberColumn("Shopee", format="%d"),
                        "Lazada":          st.column_config.NumberColumn("Lazada", format="%d"),
                        "TikTok":          st.column_config.NumberColumn("TikTok", format="%d"),
                        "CRR_%":           st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                        "Avg_CSAT":        st.column_config.NumberColumn("CSAT", format="%.2f"),
                        "Conversions":     st.column_config.NumberColumn("Conversions"),
                        "Δ Conversations": st.column_config.NumberColumn("Δ Conv", format="%+.0f"),
                        "Δ CRR_%":         st.column_config.NumberColumn("Δ CRR%", format="%+.1f"),
                        "Δ Avg_CSAT":      st.column_config.NumberColumn("Δ CSAT", format="%+.2f"),
                    },
                )

    with tab6:
        st.markdown("### 👥 Team Member Performance")
        st.caption(
            f"Data from **{TEAM_START_DATE.strftime('%d %b %Y')}** onwards · "
            f"Store → Agent mapping as configured in constants"
        )
        team_perf = compute_team_performance(conv_filtered)

        if team_perf.empty:
            st.info(
                "No team performance data available. "
                f"Check that conversations exist from {TEAM_START_DATE.strftime('%d %b %Y')} "
                "and that store codes match assignments."
            )
        else:
            agents = team_perf["TEAM_MEMBER"].tolist()
            for i in range(0, len(agents), 3):
                cols = st.columns(3)
                for j, agent in enumerate(agents[i:i+3]):
                    row_a = team_perf[team_perf["TEAM_MEMBER"] == agent].iloc[0]
                    with cols[j]:
                        st.markdown(f"""
                        <div style="background:#1B2A4A;border-radius:10px;padding:14px;color:white;margin-bottom:8px;">
                          <div style="font-size:16px;font-weight:700;color:#00C4B4;">👤 {agent}</div>
                          <div style="font-size:11px;color:#aaa;margin-bottom:8px;">{row_a['Shift']}</div>
                          <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;">
                            <div><span style="font-size:20px;font-weight:700">{int(row_a['Conversations'])}</span><br><span style="font-size:11px;color:#ccc;">Conversations</span></div>
                            <div><span style="font-size:20px;font-weight:700">{row_a['CRR_%']:.1f}%</span><br><span style="font-size:11px;color:#ccc;">CRR</span></div>
                            <div><span style="font-size:20px;font-weight:700">{row_a['Avg_CSAT']:.2f}</span><br><span style="font-size:11px;color:#ccc;">CSAT</span></div>
                            <div><span style="font-size:20px;font-weight:700">{int(row_a['Avg_CRT_mins']) if pd.notna(row_a['Avg_CRT_mins']) else '—'}m</span><br><span style="font-size:11px;color:#ccc;">Avg CRT</span></div>
                            <div><span style="font-size:20px;font-weight:700;color:#FF6B35">{int(row_a['Conversions'])}</span><br><span style="font-size:11px;color:#ccc;">Conversions</span></div>
                            <div><span style="font-size:20px;font-weight:700;color:#f87171">{int(row_a['High_Priority'])}</span><br><span style="font-size:11px;color:#ccc;">High Priority</span></div>
                          </div>
                        </div>
                        """, unsafe_allow_html=True)

            st.markdown("---")
            st.markdown("**Team Summary Table**")
            summary_cols = [
                "TEAM_MEMBER", "Shift", "Conversations", "Resolved", "Unresolved",
                "CRR_%", "Avg_CSAT", "Avg_CRT_mins", "Positive_Sent",
                "Negative_Sent", "Conversions", "High_Priority",
            ]
            st.dataframe(
                team_perf[summary_cols].reset_index(drop=True),
                use_container_width=True,
                column_config={
                    "TEAM_MEMBER":   st.column_config.TextColumn("Agent"),
                    "Shift":         st.column_config.TextColumn("Shift / Market"),
                    "Conversations": st.column_config.NumberColumn("Conv"),
                    "Resolved":      st.column_config.NumberColumn("Resolved"),
                    "Unresolved":    st.column_config.NumberColumn("Unresolved"),
                    "CRR_%":         st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                    "Avg_CSAT":      st.column_config.NumberColumn("CSAT", format="%.2f"),
                    "Avg_CRT_mins":  st.column_config.NumberColumn("CRT (min)", format="%.0f"),
                    "Positive_Sent": st.column_config.NumberColumn("Positive"),
                    "Negative_Sent": st.column_config.NumberColumn("Negative"),
                    "Conversions":   st.column_config.NumberColumn("Conversions"),
                    "High_Priority": st.column_config.NumberColumn("High Pri."),
                },
            )

            st.markdown("---")
            st.markdown("**Drill Down by Agent**")
            agent_sel = st.selectbox("Select Agent", ["(All)"] + agents)
            if agent_sel == "(All)":
                drilldown_df = conv_filtered[conv_filtered["LAST_MSG_TIME"] >= TEAM_START_DATE]
            else:
                drilldown_df = conv_filtered[
                    (conv_filtered["TEAM_MEMBER"] == agent_sel) &
                    (conv_filtered["LAST_MSG_TIME"] >= TEAM_START_DATE)
                ]

            if agent_sel == "Others" and not drilldown_df.empty:
                st.markdown("**Others — Store Code Breakdown**")
                others_summary = (
                    drilldown_df.groupby("STORE_CODE")
                    .agg(
                        Conversations=("CONVERSATION_ID", "count"),
                        Unresolved=("IS_UNRESOLVED", "sum"),
                        Avg_CSAT=("CSAT_PROXY", "mean"),
                        Platform=("PLATFORM", lambda x: x.mode().iloc[0] if not x.empty else "—"),
                        Country=("COUNTRY_CODE", lambda x: x.mode().iloc[0] if not x.empty else "—"),
                    )
                    .reset_index()
                    .sort_values("Conversations", ascending=False)
                )
                others_summary["Avg_CSAT"] = others_summary["Avg_CSAT"].round(1)
                others_summary["CRR%"] = (
                    (others_summary["Conversations"] - others_summary["Unresolved"])
                    / others_summary["Conversations"] * 100
                ).round(1)
                st.dataframe(others_summary, use_container_width=True, hide_index=True,
                    column_config={
                        "Conversations": st.column_config.NumberColumn("Conv"),
                        "Unresolved":    st.column_config.NumberColumn("Unresolved"),
                        "Avg_CSAT":      st.column_config.NumberColumn("CSAT", format="%.1f"),
                        "CRR%":          st.column_config.NumberColumn("CRR %", format="%.1f%%"),
                    })
                st.markdown("**All Conversations — Others**")

            drill_cols = [c for c in [
                "CONVERSATION_ID", "STORE_CODE", "CHANNEL_NAME", "SITE_NICK_NAME_ID",
                "COUNTRY_CODE", "BUYER_NAME", "LAST_MSG_TIME", "ISSUE_TYPE", "PRIORITY",
                "SENTIMENT", "IS_RESOLVED", "CSAT_PROXY", "AVG_CRT_MINS",
                "IS_CONVERSION", "TEAM_MEMBER",
            ] if c in drilldown_df.columns]
            st.dataframe(
                drilldown_df[drill_cols].sort_values("LAST_MSG_TIME", ascending=False).reset_index(drop=True),
                use_container_width=True, height=400,
                column_config={
                    "CSAT_PROXY":    st.column_config.NumberColumn("CSAT", format="%.1f"),
                    "AVG_CRT_MINS":  st.column_config.NumberColumn("CRT(m)", format="%.0f"),
                    "IS_RESOLVED":   st.column_config.CheckboxColumn("Resolved?"),
                    "IS_CONVERSION": st.column_config.CheckboxColumn("Conversion?"),
                },
            )

            st.markdown("---")
            st.markdown("**🔍 Unassigned / Other Stores in This Data**")
            st.caption("Based on current sidebar filters.")
            all_known = set(STORE_TO_AGENT.keys())
            if "STORE_CODE" in conv_filtered.columns:
                others_stores = sorted(
                    s for s in conv_filtered["STORE_CODE"].dropna().unique()
                    if str(s).strip().upper() not in all_known and str(s).strip()
                )
                if others_stores:
                    others_rows = []
                    for sc in others_stores:
                        sc_df = conv_filtered[conv_filtered["STORE_CODE"] == sc]
                        site = sc_df["SITE_NICK_NAME_ID"].mode().iloc[0] if "SITE_NICK_NAME_ID" in sc_df.columns and not sc_df["SITE_NICK_NAME_ID"].dropna().empty else "—"
                        channel = sc_df["CHANNEL_NAME"].mode().iloc[0] if "CHANNEL_NAME" in sc_df.columns and not sc_df["CHANNEL_NAME"].replace("", pd.NA).dropna().empty else "—"
                        country = sc_df["COUNTRY_CODE"].mode().iloc[0] if "COUNTRY_CODE" in sc_df.columns and not sc_df["COUNTRY_CODE"].dropna().empty else "—"
                        platform = sc_df["PLATFORM"].mode().iloc[0] if "PLATFORM" in sc_df.columns and not sc_df["PLATFORM"].dropna().empty else "—"
                        others_rows.append({
                            "Store Code":    sc,
                            "Channel Name":  channel,
                            "Site Nickname": site,
                            "Platform":      platform,
                            "Country":       country,
                            "Conversations": len(sc_df),
                            "Unresolved":    int(sc_df["IS_UNRESOLVED"].sum()) if "IS_UNRESOLVED" in sc_df.columns else 0,
                            "Avg CSAT":      round(sc_df["CSAT_PROXY"].mean(), 1) if "CSAT_PROXY" in sc_df.columns else "—",
                            "Assign To":     "⚠️ Not assigned",
                        })
                    others_df = pd.DataFrame(others_rows).sort_values("Conversations", ascending=False)
                    st.dataframe(others_df, use_container_width=True, hide_index=True,
                        column_config={
                            "Conversations": st.column_config.NumberColumn("Conv", format="%d"),
                            "Unresolved":    st.column_config.NumberColumn("Unresolved", format="%d"),
                            "Avg CSAT":      st.column_config.NumberColumn("CSAT", format="%.1f"),
                        })
                    st.warning(f"⚠️ **{len(others_stores)} store(s)** found with no team member assigned.")
                else:
                    st.success("✅ All stores in this dataset are assigned to team members.")

            with st.expander("📋 Store → Agent Assignment Reference"):
                assign_rows = [
                    {"Agent": agent_name, "Shift": AGENT_SHIFT.get(agent_name, "Day"),
                     "Assigned Stores": ", ".join(stores)}
                    for agent_name, stores in TEAM_ASSIGNMENTS.items()
                ]
                st.dataframe(pd.DataFrame(assign_rows), use_container_width=True, hide_index=True)

    # ── Issue Breakdown Table ─────────────────────────────────────────────────
    st.markdown('<div class="section-title">📂 Issue Type Breakdown</div>', unsafe_allow_html=True)

    _plat_opts = ["All Platforms"] + sorted(conv_filtered["PLATFORM"].dropna().unique().tolist())
    _sel_plat_ib = st.selectbox("Filter by Platform", _plat_opts, key="ib_platform")
    _ib_df = conv_filtered if _sel_plat_ib == "All Platforms" else conv_filtered[conv_filtered["PLATFORM"] == _sel_plat_ib]

    ib = (
        _ib_df
        .groupby(["ISSUE_TYPE", "PRIORITY"])
        .agg(
            Count=("CONVERSATION_ID", "count"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
        )
        .reset_index()
        .sort_values("Count", ascending=False)
    )
    ib["Avg_CSAT"]     = ib["Avg_CSAT"].round(1)
    ib["Avg_CRT_mins"] = ib["Avg_CRT_mins"].round(0).fillna(0).astype(int)
    ib["Unresolved"]   = ib["Unresolved"].astype(int)
    st.dataframe(ib, use_container_width=True, height=300,
        column_config={
            "Count":         st.column_config.NumberColumn("Count"),
            "Unresolved":    st.column_config.NumberColumn("Unresolved"),
            "Avg_CSAT":      st.column_config.NumberColumn("CSAT", format="%.1f"),
            "Avg_CRT_mins":  st.column_config.NumberColumn("CRT (min)"),
        }
    )

    # ── Store Performance ─────────────────────────────────────────────────────
    st.markdown('<div class="section-title">🏪 Store Performance</div>', unsafe_allow_html=True)
    sp = (
        conv_filtered
        .groupby(["STORE_CODE", "PLATFORM", "COUNTRY_CODE"])
        .agg(
            Conversations=("CONVERSATION_ID", "count"),
            Unresolved=("IS_UNRESOLVED", "sum"),
            Avg_CSAT=("CSAT_PROXY", "mean"),
            Avg_CRT_mins=("AVG_CRT_MINS", "mean"),
            Negative_Sent=("SENTIMENT", lambda x: (x == "Negative").sum()),
        )
        .reset_index()
        .sort_values("Conversations", ascending=False)
    )
    sp["Avg_CSAT"]     = sp["Avg_CSAT"].round(1)
    sp["Avg_CRT_mins"] = sp["Avg_CRT_mins"].round(0).fillna(0).astype(int)
    sp["Unresolved"]   = sp["Unresolved"].astype(int)
    sp["CRR%"]         = ((sp["Conversations"] - sp["Unresolved"]) / sp["Conversations"] * 100).round(1)
    sp = sp[["STORE_CODE", "PLATFORM", "COUNTRY_CODE", "Conversations",
             "Unresolved", "CRR%", "Avg_CSAT", "Avg_CRT_mins", "Negative_Sent"]]
    st.dataframe(
        sp,
        use_container_width=True,
        height=350,
        column_config={
            "STORE_CODE":   st.column_config.TextColumn("Store Code"),
            "PLATFORM":     st.column_config.TextColumn("Platform"),
            "COUNTRY_CODE": st.column_config.TextColumn("Country"),
            "Conversations":st.column_config.NumberColumn("Conversations"),
            "Unresolved":   st.column_config.NumberColumn("Unresolved"),
            "CRR%":         st.column_config.NumberColumn("CRR %", format="%.1f%%"),
            "Avg_CSAT":     st.column_config.NumberColumn("CSAT", format="%.1f"),
            "Avg_CRT_mins": st.column_config.NumberColumn("CRT (min)", format="%d"),
            "Negative_Sent":st.column_config.NumberColumn("Negative Sent"),
        },
    )

    # ── Excel Download ────────────────────────────────────────────────────────
    st.markdown('<div class="section-title">⬇️ Download Report</div>', unsafe_allow_html=True)
    cutoff_7d  = pd.Timestamp(data_end) - pd.Timedelta(days=6)
    conv_7day  = conv_df[conv_df["LAST_MSG_TIME"] >= cutoff_7d].copy()

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        if st.button("📊 Generate Last 7 Days Report", use_container_width=True):
            with st.spinner("Building report…"):
                excel_7day = build_excel(conv_7day, today_str)
            st.download_button(
                label=f"📥 Download ({cutoff_7d.date()} → {data_end})",
                data=excel_7day,
                file_name=f"Chat_Analysis_Last7Days_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with dl_col2:
        if st.button("📊 Generate Filtered View Report", use_container_width=True):
            with st.spinner("Building report…"):
                excel_filtered = build_excel(conv_filtered, today_str)
            st.download_button(
                label="📥 Download Filtered View",
                data=excel_filtered,
                file_name=f"Chat_Analysis_Filtered_{today_str}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    st.caption(
        "Click **Generate** first, then **Download**. "
        "**Last 7 Days** = default daily export · "
        "**Filtered View** = matches current sidebar selection."
    )


if __name__ == "__main__":
    main()
