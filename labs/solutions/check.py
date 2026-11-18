# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Run the reference solution of every lab against the course files and check it.

    uv run labs/solutions/check.py            # all labs
    uv run labs/solutions/check.py 3.2 4.3    # some labs

Every lab of this course is done in a spreadsheet, a diagram or a written letter; none calls a
model. This script does what the reference solution does on the files (rule lookups, class sizes,
linkage counts, due dates, shares, counts in logs) and checks the answer keys in
solutions/reference/ against the files, so a key the files do not support fails here. It prints
the numbers the debrief slides quote and PASS per lab. Nothing is written to disk.
"""
from __future__ import annotations

import csv
import json
import sys
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

import openpyxl

HERE = Path(__file__).resolve().parent
FILES = HERE.parent / "files"
REF = HERE / "reference"
LABS: dict = {}


def lab(lab_id):
    def wrap(fn):
        LABS[lab_id] = fn
        return fn
    return wrap


def ref(name):
    return json.loads((REF / name).read_text(encoding="utf-8"))


def text(rel):
    return (FILES / rel).read_text(encoding="utf-8")


def rows(path, sheet):
    ws = openpyxl.load_workbook(path, read_only=True)[sheet]
    it = ws.iter_rows(values_only=True)
    head = next(it)
    return [dict(zip(head, r)) for r in it if any(v is not None for v in r)]


def band(year: int) -> int:
    return ((2026 - year) // 10) * 10


@lab("1.1")
def lab_1_1():
    key = ref("lab_1_1.json")["key"]
    acts = rows(FILES / "basis/activities_12.xlsx", "الأنشطة")
    codes = {r["الرمز"] for r in rows(FILES / "basis/activities_12.xlsx", "الأسس")}
    assert len(acts) == 12 and set(key) == {r["المعرّف"] for r in acts}
    assert set(key.values()) <= codes
    # rules the key must respect: no legitimate interest on health data; explicit consent where sensitive or credit
    for r in acts:
        k = key[r["المعرّف"]]
        sensitive = "صحية" in r["البيانات"] or "ائتماني" in r["البيانات"]
        if sensitive:
            assert k != "B6", (r["المعرّف"], "legitimate interest on sensitive data")
            assert k in ("B2", "B4", "B7"), (r["المعرّف"], k)
    dist = Counter(key.values())
    print("  basis distribution:", dict(sorted(dist.items())))
    # the twelve activities must not be activities of the day-2 register
    reg = [r["النشاط"] for r in rows(FILES / "flows/systems.xlsx", "الأنشطة العشرة")]
    for r in acts:
        assert r["النشاط"] not in reg


@lab("1.2")
def lab_1_2():
    key = ref("lab_1_2.json")
    ev = rows(FILES / "obligations/events_10.xlsx", "الأحداث")
    assert [r["المعرّف"] for r in ev] == list(key)
    assert key["E01"]["when"].startswith("خلال 30") and "72" in key["E03"]["when"] and "90" in key["E08"]["when"]
    print("  events:", len(ev), "| deadlines with a number:", sum(1 for v in key.values() if any(ch.isdigit() for ch in v["when"])))


@lab("1.3")
def lab_1_3():
    key = ref("lab_1_3.json")
    book = FILES / "officer/officer_pack.xlsx"
    assert len(rows(book, "حالات التعيين")) == 4
    ind = rows(book, "المؤشرات")
    assert len(ind) == 14 and len(key["indicators_good"]) == 8 and len(key["indicators_vanity"]) == 6
    assert sum(1 for v in key["appoint"].values() if v == "نعم") == 3
    assert [k for k, v in key["options"].items() if v == "نعم"] == ["O3"]
    print("  must appoint: 3 of 4 | acceptable officer option: O3 | good indicators: 8 of 14, choose 6")


@lab("1.4")
def lab_1_4():
    key = ref("lab_1_4.json")
    cards = rows(FILES / "status/status_cards.xlsx", "البطاقات")
    assert len(cards) == 12
    dist = Counter(key["cards"].values())
    assert dist == {"نافذ": 5, "دليل": 3, "مسودة": 2, "غير موجود": 2}, dist
    assert text("status/scenarios_practice.txt").count("\nP") >= 7
    print("  cards:", dict(dist))


@lab("2.1")
def lab_2_1():
    key = ref("lab_2_1.json")
    sysrows = rows(FILES / "flows/systems.xlsx", "الأنظمة")
    assert len(sysrows) == key["systems"] == 11
    abroad = [r["المعرّف"] for r in sysrows if "خارج المملكة" in (r["جهات المعالجة الفرعية والنسخ"] or "")]
    assert abroad == ["SY07"], abroad
    assert "خارج" in text("flows/interviews/I3_support.txt")
    assert "سبب الزيارة" in text("flows/interviews/I4_contracts.txt")
    assert "تفريغ نصي" in text("flows/interviews/I1_cto.txt")
    assert "صندوق بريد مشترك" in text("flows/interviews/I2_medical.txt")
    print("  systems: 11 | planted flows: F1 to F4 found in the files")


@lab("2.2")
def lab_2_2():
    key = ref("lab_2_2.json")
    reg = rows(HERE / "reference/register_reference.xlsx", "السجل")
    assert len(reg) == 10
    blanks = sum(1 for r in reg for c in key["columns"] if not r.get(c))
    assert blanks == 0, blanks
    by = {r["النشاط"].split()[0]: r for r in reg}
    assert "خارج المملكة" in by["A7"]["النقل خارج المملكة ومسوغه"]
    assert "لا أسماء" in by["A8"]["فئات البيانات"]
    assert "90 يومًا" in by["A2"]["مدة الاحتفاظ"]
    sensitive = sum(1 for r in reg if r["بيانات حساسة؟"] == "نعم")
    dpia = sum(1 for r in reg if r["يحتاج تقويم أثر؟"] == "نعم")
    # Art 25 (1) (a): every activity with sensitive data needs an impact assessment
    assert all(r["يحتاج تقويم أثر؟"] == "نعم" for r in reg if r["بيانات حساسة؟"] == "نعم")
    assert (sensitive, dpia) == (7, 8), (sensitive, dpia)
    template = rows(FILES / "register/register_template.xlsx", "السجل")
    assert len([r for r in template if r.get("النشاط")]) >= 10
    print(f"  reference register: 10 activities x {len(key['columns'])} fields, 0 blank | sensitive: {sensitive} | need an assessment: {dpia}")


@lab("2.3")
def lab_2_3():
    key = ref("lab_2_3.json")
    notice = text("notices/privacy_notice.txt")
    assert "مسؤول حماية" not in notice and "نحتفظ" not in notice and "خارج المملكة" not in notice
    assert "سحب الموافقة" not in notice and "إلزامي" not in notice and "اختياري" not in notice
    dpa = text("notices/video_dpa.txt")
    assert "تسرب" not in dpa and "فرعية" not in dpa and "دول أخرى" not in dpa
    print(f"  notice gaps: {len(key['notice_gaps'])} | agreement gaps: {len(key['dpa_gaps'])}")


@lab("2.4")
def lab_2_4():
    key = ref("lab_2_4.json")
    fields = rows(FILES / "minimise/intake_fields.xlsx", "حقول النموذج")
    assert len(fields) == 14 == len(key["fields"])
    for r in fields:
        if r["الاستخدام المذكور"] == "لا استخدام مذكور":
            assert key["fields"][r["الحقل"]] == "يُحذف"
    print("  fields:", key["counts"], "| retention changes:", sum(1 for k in ("RT2", "RT3", "RT5") if k in key["retention"]))


@lab("3.1")
def lab_3_1():
    key = ref("lab_3_1.json")
    ini = rows(FILES / "screening/initiatives_8.xlsx", "المبادرات")
    assert len(ini) == 8
    yes = [k for k, v in key.items() if v["dpia"] == "نعم"]
    tra = [k for k, v in key.items() if v["tra"].startswith("نعم")]
    print(f"  assessment needed: {len(yes)} of 8 ({', '.join(yes)}) | transfer risk assessment: {', '.join(tra)}")
    assert len(yes) == 5


@lab("3.2")
def lab_3_2():
    key = ref("lab_3_2.json")
    ext = rows(FILES / "deid/analytics_extract.xlsx", "المستخلص")
    assert len(ext) == key["rows"]
    for r in ext:
        assert r["رقم المريض"].startswith("QRB-TEST-") and "TEST" in r["الجوال"]

    def share(keyfn):
        ks = [keyfn(r) for r in ext]
        c = Counter(ks)
        small = sum(1 for k in ks if c[k] < 5)
        return {"records_k_lt_5": small, "share_k_lt_5": round(small / len(ext), 4), "unique": sum(1 for k in ks if c[k] == 1)}

    by = lambda r: r["تاريخ الميلاد"].year
    raw = share(lambda r: (r["تاريخ الميلاد"], r["الجنس"], r["الحي"]))
    g1 = share(lambda r: (band(by(r)), r["الجنس"], r["المدينة"]))
    g1e = share(lambda r: (band(by(r)), r["الجنس"], r["المدينة"], r["جهة العمل"] or ""))
    assert raw == key["raw_qi"] and g1 == key["generalised"] and g1e == key["generalised_with_employer"], (raw, g1, g1e)
    d = rows(FILES / "deid/linkage_list.xlsx", "دليل الموظفين")
    assert len(d) == key["directory"]
    e1 = [r for r in ext if r["جهة العمل"] == key["employer_e1"]]
    c_year = Counter((r["سنة الميلاد"], r["الجنس"], r["المدينة"]) for r in d)
    c_band = Counter((band(r["سنة الميلاد"]), r["الجنس"], r["المدينة"]) for r in d)
    linked_year = sum(1 for r in e1 if c_year[(by(r), r["الجنس"], r["المدينة"])] == 1)
    linked_band = sum(1 for r in e1 if c_band[(band(by(r)), r["الجنس"], r["المدينة"])] == 1)
    assert linked_year == key["linked_unique_birthyear"] and linked_band == key["linked_unique_band"]
    print(f"  raw (birth date, gender, district): k<5 for {raw['records_k_lt_5']} of 2000, unique {raw['unique']}")
    print(f"  age band + gender + city: k<5 for {g1['records_k_lt_5']} | with employer kept: k<5 for {g1e['records_k_lt_5']} ({g1e['share_k_lt_5']:.1%}), unique {g1e['unique']}")
    print(f"  linkage to the {len(d)}-name directory: {len(e1)} extract rows of that employer; linked to one name by birth year {linked_year}, by age band {linked_band}")


@lab("3.3")
def lab_3_3():
    key = ref("lab_3_3_3_4.json")["risks"]
    ref_rows = rows(FILES / "hads/risk_register_reference.xlsx", "السجل المرجعي")
    for r in ref_rows:
        k = key[r["المعرّف"]]
        assert r["الجسامة"] * r["الاحتمال"] == k["score"] == r["الدرجة"]
        assert r["النطاق"] == k["band"]
    bands = Counter(v["band"] for v in key.values())
    print("  inherent bands:", dict(bands), "| highest:", max(key.items(), key=lambda kv: kv[1]["score"])[0])


@lab("3.4")
def lab_3_4():
    full = ref("lab_3_3_3_4.json")
    key = full["risks"]
    catalogue = {r["الرمز"] for r in rows(FILES / "hads/risk_register_reference.xlsx", "دليل الضوابط")}
    for v in key.values():
        assert v["control"] in catalogue and v["control"] not in full["ineffective_controls"]
        assert v["residual"] < v["score"]
    worst = max(v["residual"] for v in key.values())
    bands = Counter(v["residual_band"] for v in key.values())
    print(f"  residual bands: {dict(bands)} | worst residual score: {worst}")
    assert worst <= 6


@lab("4.1")
def lab_4_1():
    key = ref("lab_4_1_4_2.json")
    today = date.fromisoformat(key["today"])
    order = []
    for rid, v in key["requests"].items():
        t = text(f"requests/{rid}.txt")
        rec = date.fromisoformat(t.split("تاريخ الاستلام: ")[1].split()[0])
        due = rec + timedelta(days=30)
        assert due.isoformat() == v["due"] and (due - today).days == v["days_left"]
        order.append(((due - today).days, rid))
    order.sort()
    print("  days left at the start of the lab:", ", ".join(f"{rid} {d}" for d, rid in order))
    assert order[0][1] == "R2"


@lab("4.2")
def lab_4_2():
    key = ref("lab_4_1_4_2.json")
    draft = text("letters/assistant_draft_R3.txt")
    assert "المادة 40" in draft and "60 يومًا" in draft
    print("  planted errors in the stand-in draft:", len(key["draft_errors"]))


@lab("4.3")
def lab_4_3():
    key = ref("lab_4_3.json")
    t = rows(FILES / "transfer/tickets_sample.xlsx", "التذاكر")
    health = sum(1 for r in t if r["في النص بيانات صحية؟"] == "نعم")
    lab_att = sum(1 for r in t if r["نوع المرفق"] == "نتيجة تحليل")
    sens = sum(1 for r in t if r["في النص بيانات صحية؟"] == "نعم" or r["نوع المرفق"] == "نتيجة تحليل")
    assert (len(t), health, lab_att, sens) == (key["tickets"], key["health_text"], key["lab_attachments"], key["sensitive_any"])
    print(f"  tickets: {len(t)} | health text: {health} | lab-result attachments: {lab_att} | any sensitive: {sens} ({sens / len(t):.1%})")


@lab("4.4")
def lab_4_4():
    key = ref("lab_4_4.json")
    with open(FILES / "drill/login_log.csv", encoding="utf-8-sig") as f:
        log = list(csv.DictReader(f))
    ok = {r["account"] for r in log if r["result"] == "نجاح"}
    rx = {r["account"] for r in log if r["viewed"] == "المواعيد والوصفات"}
    assert (len(log), len(ok), len(rx)) == (key["attempts"], key["accounts_breached"], key["viewed_prescriptions"])
    print(f"  attempts {len(log)} | accounts entered {len(ok)} | saw prescriptions {len(rx)} | notice due at hour {key['authority_deadline_hour']}, {key['hours_left_when_dpo_learns']} hours after the officer learns")


@lab("capstone")
def lab_capstone():
    key = ref("capstone.json")
    ex = rows(FILES / "capstone/misdirected_export.xlsx", "التصدير")
    minors = sum(1 for r in ex if r["العمر"] < 18)
    special = sum(1 for r in ex if r["التحليل المطلوب"] in ("فحص ما قبل الزواج", "فحص الحمل", "الأمراض المنقولة جنسيًا"))
    assert (len(ex), minors, special) == (key["rows"], key["minors"], key["sensitive_tests"])
    assert key["authority_deadline_hour"] == key["aware_hour"] + 72
    assert "خمسة أيام" in text("capstone/assistant_draft_notice.txt")
    print(f"  export rows {len(ex)} | minors {minors} | premarital, pregnancy or STI tests {special} | notice due at simulation hour {key['authority_deadline_hour']} (end of phase 2)")


def main(argv):
    sys.stdout.reconfigure(encoding="utf-8")
    wanted = argv or list(LABS)
    failed = 0
    for lab_id in wanted:
        print(f"lab {lab_id}")
        try:
            LABS[lab_id]()
            print("  PASS")
        except Exception as e:  # report and continue
            failed += 1
            print("  FAIL:", repr(e))
    print(f"{len(wanted) - failed} of {len(wanted)} labs pass")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
