"""
Mapping Sheet Builder  (ShelfWatch / DAP)
-----------------------------------------
Label -> Group sheet  +  CGC (class_name, group_name, image link)
        =>  Mapping sheet (Parent group | Parent Class | Child Group | Child Class)
        +   Error sheet (jo group CGC me add nahi hue, image missing / mismatch, etc.)

Run:  streamlit run app.py
"""
import csv
import difflib
import html
import io
import re
import unicodedata
from collections import OrderedDict, defaultdict

import pandas as pd
import streamlit as st

OTHERS = "others"
IMG_HINT = re.compile(r"image|img|link|url|path|photo")


# ----------------------------------------------------------------------------
# Normalisation helpers
# ----------------------------------------------------------------------------
def strict(s) -> str:
    """lower-case, accents removed, extra spaces removed."""
    if s is None:
        return ""
    s = html.unescape(str(s))
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip().lower()


def loose(s) -> str:
    """strict + trailing _1/_2 removed + punctuation ignored."""
    s = strict(s)
    s = re.sub(r"_\d+$", "", s)
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def is_others(s) -> bool:
    return strict(s) in ("other", "others")


def cell(row, j) -> str:
    if j is None or j < 0 or j >= len(row):
        return ""
    v = str(row[j]).replace("\xa0", " ").strip()
    return "" if v.lower() in ("nan", "none") else v


def norm_url(u: str) -> str:
    return u.strip().lower()


# ----------------------------------------------------------------------------
# File reading
# ----------------------------------------------------------------------------
def read_raw(name: str, data: bytes) -> "dict[str, list[list[str]]]":
    """Return {sheet_name: rows}. Works for csv / xlsx / xls."""
    low = name.lower()
    if low.endswith((".xlsx", ".xlsm", ".xls")):
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=str)
        return {k: v.fillna("").astype(str).values.tolist() for k, v in sheets.items()}
    text = None
    for enc in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = data.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    try:
        delim = csv.Sniffer().sniff(text[:4096], delimiters=",;\t").delimiter
    except Exception:
        delim = ","
    rows = list(csv.reader(io.StringIO(text), delimiter=delim))
    return {"CSV": rows}


# ----------------------------------------------------------------------------
# Label -> Group sheet parsing
# ----------------------------------------------------------------------------
def parse_label_sheet(rows):
    """
    Header row = row jisme 'PARENT ...' likha ho. Us column ke baad ke saare
    non-empty cells children hain.  Returns list of dict.
    """
    hdr_i, parent_j = -1, 0
    for i, r in enumerate(rows[:15]):
        for j, c in enumerate(r):
            if strict(c).startswith("parent"):
                hdr_i, parent_j = i, j
                break
        if hdr_i >= 0:
            break
    header = rows[hdr_i] if hdr_i >= 0 else []
    date_j = next((j for j, c in enumerate(header) if strict(c) == "date"), None)
    name_j = next((j for j, c in enumerate(header) if strict(c) in ("name", "labeler", "user")), None)

    recs = []
    for ri in range(hdr_i + 1, len(rows)):
        r = rows[ri]
        parent = cell(r, parent_j)
        kids = [cell(r, j) for j in range(parent_j + 1, len(r))]
        kids = [k for k in kids if k]
        if not parent and not kids:
            continue
        recs.append(
            dict(
                sheet_row=ri + 1,
                date=cell(r, date_j),
                labeler=cell(r, name_j),
                parent=parent,
                children=kids,
            )
        )
    return recs


def parse_mapping_sheet(rows):
    """Existing mapping file (4 columns) -> pairs with given classes."""
    hdr_i = 0
    for i, r in enumerate(rows[:10]):
        if any(strict(c) == "parent group" for c in r):
            hdr_i = i
            break
    header = rows[hdr_i]

    def col(*names):
        for j, c in enumerate(header):
            if strict(c) in names:
                return j
        return None

    pg, pc = col("parent group"), col("parent class")
    cg, cc = col("child group"), col("child class")
    if None in (pg, pc, cg, cc):
        pg, pc, cg, cc = 0, 1, 2, 3
    pairs = []
    for ri in range(hdr_i + 1, len(rows)):
        r = rows[ri]
        if not any(cell(r, j) for j in (pg, pc, cg, cc)):
            continue
        pairs.append(
            dict(
                sheet_row=ri + 1,
                date="",
                labeler="",
                parent=cell(r, pg),
                child=cell(r, cg),
                parent_class_given=cell(r, pc),
                child_class_given=cell(r, cc),
            )
        )
    return pairs


def labels_to_pairs(recs):
    """Unpivot parent -> children. Returns (pairs, pre_issues)."""
    pairs, issues = [], []

    def add(kind, sev, detail, r, child=""):
        issues.append(
            {
                "Issue": kind,
                "Severity": sev,
                "Details": detail,
                "Parent group": r["parent"],
                "Child group": child,
                "Labeler": r["labeler"],
                "Date": r["date"],
                "Sheet row": r["sheet_row"],
                "Suggestion": "",
            }
        )

    for r in recs:
        if not r["parent"]:
            for k in r["children"]:
                add("Parent blank", "Error", "Child diya hai lekin Parent group khali hai", r, k)
            continue
        if not r["children"]:
            add("No child", "Error", "Parent group hai lekin koi child group nahi", r)
            continue
        seen = set()
        for k in r["children"]:
            ks = strict(k)
            if ks == strict(r["parent"]) or (is_others(r["parent"]) and is_others(k)):
                # "other | other" jaisi rows: child parent ke barabar -> skip
                if not is_others(r["parent"]):
                    add("Child = Parent", "Warning", "Child ka naam parent ke same hai, skip kiya", r, k)
                continue
            if ks in seen:
                add("Duplicate in same row", "Warning", "Ek hi row me child 2 baar likha hai, ek hi liya", r, k)
                continue
            seen.add(ks)
            pairs.append(
                dict(
                    sheet_row=r["sheet_row"],
                    date=r["date"],
                    labeler=r["labeler"],
                    parent=r["parent"],
                    child=k,
                    parent_class_given="",
                    child_class_given="",
                )
            )
    return pairs, issues


# ----------------------------------------------------------------------------
# CGC (class + group + image) index
# ----------------------------------------------------------------------------
def guess_cgc_columns(rows):
    hdr_i = 0
    for i, r in enumerate(rows[:10]):
        s = " ".join(strict(c) for c in r)
        if "class" in s or "group" in s:
            hdr_i = i
            break
    header = [cell(rows[hdr_i], j) for j in range(len(rows[hdr_i]))]
    hs = [strict(h) for h in header]

    def pick(exact, contains):
        for j, h in enumerate(hs):
            if re.fullmatch(exact, h):
                return j
        for j, h in enumerate(hs):
            if contains in h and not IMG_HINT.search(h) and "id" not in h.split():
                return j
        return None

    class_j = pick(r"class[ _]?name", "class")
    group_j = pick(r"group[ _]?name", "group")
    img_js = [j for j, h in enumerate(hs) if IMG_HINT.search(h)]
    return hdr_i, header, class_j, group_j, img_js


class CGC:
    def __init__(self):
        self.group = OrderedDict()   # strict(group) -> entry
        self.klass = OrderedDict()   # strict(class) -> entry
        self.has_images = False
        self.n_rows = 0

    @staticmethod
    def _entry(name):
        return {"name": name, "classes": OrderedDict(), "groups": OrderedDict(), "images": OrderedDict()}

    def add(self, cls, grp, imgs):
        self.n_rows += 1
        grp = grp or cls
        cls = cls or grp
        if not grp and not cls:
            return
        for idx, key, nm in ((self.group, strict(grp), grp), (self.klass, strict(cls), cls)):
            e = idx.setdefault(key, self._entry(nm))
            e["classes"].setdefault(strict(cls), cls)
            e["groups"].setdefault(strict(grp), grp)
            for u in imgs:
                e["images"].setdefault(norm_url(u), u)

    def finalize(self):
        self.loose_group, self.loose_class = defaultdict(set), defaultdict(set)
        for k in self.group:
            self.loose_group[loose(k)].add(k)
        for k in self.klass:
            self.loose_class[loose(k)].add(k)
        self.cands = {}
        for k, e in self.group.items():
            self.cands[k] = e["name"]
        for k, e in self.klass.items():
            self.cands.setdefault(k, e["name"])
        self.cand_keys = list(self.cands.keys())

    def lookup(self, name):
        """Returns dict(found, group, cls, classes, images, match)."""
        ks = strict(name)
        res = dict(found=False, group=name, cls="", classes=[], images=[], match="")
        entry, how, is_group = None, "", False
        if ks in self.group:
            entry, how, is_group = self.group[ks], "exact", True
        elif ks in self.klass:
            entry, how = self.klass[ks], "exact"
        else:
            lk = loose(name)
            g = self.loose_group.get(lk, set())
            c = self.loose_class.get(lk, set())
            if len(g) == 1:
                entry, how, is_group = self.group[next(iter(g))], "loose", True
            elif not g and len(c) == 1:
                entry, how = self.klass[next(iter(c))], "loose"
        if entry is None:
            return res
        classes = list(entry["classes"].values())
        cls = next((v for k, v in entry["classes"].items() if k == strict(entry["name"]) or k == ks), classes[0])
        res.update(
            found=True,
            group=entry["name"] if is_group else name.strip(),
            cls=cls,
            classes=classes,
            images=list(entry["images"].values()),
            match=how,
        )
        return res

    def suggest(self, name, n=2):
        ks = strict(name)
        if not ks:
            return []
        hits = difflib.get_close_matches(ks, self.cand_keys, n=n, cutoff=0.72)
        return [self.cands[h] for h in hits]


def build_cgc(rows, hdr_i, class_j, group_j, img_js) -> CGC:
    cg = CGC()
    cg.has_images = bool(img_js)
    for r in rows[hdr_i + 1:]:
        cls, grp = cell(r, class_j), cell(r, group_j)
        if not cls and not grp:
            continue
        imgs = []
        for j in img_js:
            v = cell(r, j)
            if not v:
                continue
            found = re.findall(r"https?://[^\s,|;\"']+", v)
            imgs.extend(found if found else [v])
        cg.add(cls, grp, imgs)
    cg.finalize()
    return cg


# ----------------------------------------------------------------------------
# Core analysis
# ----------------------------------------------------------------------------
def analyze(pairs, cgc: CGC, pre_issues=None, check_classes=False, flag_split_pending=True):
    issues = list(pre_issues or [])
    mapping, review = [], []
    notin = {}
    logged_missing = set()
    sug_cache = {}

    def sug(name):
        if name not in sug_cache:
            sug_cache[name] = cgc.suggest(name)
        return sug_cache[name]

    def add(kind, sev, detail, p, child="", suggestion=""):
        issues.append(
            {
                "Issue": kind,
                "Severity": sev,
                "Details": detail,
                "Parent group": p["parent"],
                "Child group": child,
                "Labeler": p["labeler"],
                "Date": p["date"],
                "Sheet row": p["sheet_row"],
                "Suggestion": suggestion,
            }
        )

    def note_missing(p, name, role):
        k = strict(name)
        e = notin.setdefault(
            k,
            {"Group (jaisa file me likha)": name, "Role": set(), "Rows": 0, "Labeler": set(), "Date": set(),
             "Closest CGC match": ", ".join(sug(name))},
        )
        e["Role"].add(role)
        e["Rows"] += 1
        if p["labeler"]:
            e["Labeler"].add(p["labeler"])
        if p["date"]:
            e["Date"].add(p["date"])

    child_parents = defaultdict(set)
    for p in pairs:
        child_parents[strict(p["child"])].add(OTHERS if is_others(p["parent"]) else strict(p["parent"]))

    seen_pairs = set()
    img_to_children = defaultdict(OrderedDict)  # image -> {child key: child name}
    parents_used = set()
    referenced = {}  # strict group -> (role, name, images)

    for p in pairs:
        parent_other = is_others(p["parent"])
        pk = OTHERS if parent_other else strict(p["parent"])
        ck = strict(p["child"])
        if (pk, ck) in seen_pairs:
            add("Duplicate row", "Warning", "Same parent + child dobara aaya, sirf pehla liya", p, p["child"])
            continue
        seen_pairs.add((pk, ck))

        errors = 0
        child = cgc.lookup(p["child"])
        parent = None if parent_other else cgc.lookup(p["parent"])
        if not parent_other:
            parents_used.add(pk)

        # ---- parent checks
        if not parent_other and not parent["found"]:
            errors += 1
            note_missing(p, p["parent"], "Parent")
            if pk not in logged_missing:
                logged_missing.add(pk)
                add("Parent group CGC me nahi mila", "Error",
                    "Parent group CGC me add nahi hai (name ya group dobara check karo)",
                    p, "", ", ".join(sug(p["parent"])))
        # ---- child checks
        if not child["found"]:
            errors += 1
            note_missing(p, p["child"], "Child")
            add("Child group CGC me nahi mila", "Error",
                "Child group abhi CGC me add nahi hua", p, p["child"], ", ".join(sug(p["child"])))

        if child["found"] and not parent_other and parent["found"] and check_classes:
            pass  # class check below (needs both)

        # ---- loose-match warnings
        for who, res, nm in (("Parent", parent, p["parent"]), ("Child", child, p["child"])):
            if res and res["found"] and res["match"] == "loose":
                add("Loose match", "Warning",
                    f"{who} '{nm}' ka exact naam nahi mila, '{res['group']}' se match kiya (punctuation/_1 ignore)",
                    p, p["child"])

        # ---- image checks
        if cgc.has_images:
            if child["found"] and not child["images"]:
                errors += 1
                add("Child image missing", "Error", "Child group CGC me hai lekin class image nahi mili", p, p["child"])
            if parent and parent["found"] and not parent["images"]:
                errors += 1
                add("Parent image missing", "Error", "Parent group CGC me hai lekin class image nahi mili", p, p["child"])
            if child["found"] and len(child["images"]) > 1:
                add("Split pending (child multi-image)", "Warning",
                    f"Child group me {len(child['images'])} images hain -> alag alag group banane padenge",
                    p, p["child"])
            if child["found"] and parent and parent["found"]:
                same = {norm_url(u) for u in child["images"]} & {norm_url(u) for u in parent["images"]}
                if same:
                    add("Same image in parent & child", "Warning",
                        "Parent aur child dono me same image hai (image child me move nahi hui?)",
                        p, p["child"])
            if child["found"]:
                for u in child["images"]:
                    img_to_children[norm_url(u)].setdefault(ck, p["child"])

        # ---- multi parent
        if len(child_parents[ck]) > 1:
            add("Child multiple parents me", "Warning",
                "Ye child group 1 se zyada parent ke neeche likha hai", p, p["child"])

        # ---- class check (validate-existing-mapping mode)
        if check_classes:
            for who, res, given in (
                ("Parent", parent, p["parent_class_given"]),
                ("Child", child, p["child_class_given"]),
            ):
                if res is None:
                    if strict(given) not in ("", OTHERS, "other"):
                        add("Class mismatch", "Error", "Parent others hai lekin class 'others' nahi", p, p["child"])
                        errors += 1
                    continue
                if res["found"] and strict(given) not in {strict(c) for c in res["classes"]}:
                    errors += 1
                    add("Class mismatch", "Error",
                        f"{who} class '{given}' CGC se match nahi karti. CGC class: {' | '.join(res['classes'])}",
                        p, p["child"], " | ".join(res["classes"]))

        for res in (parent, child):
            if res and res["found"]:
                referenced[strict(res["group"])] = res

        row = {
            "Parent group": "others" if parent_other else (parent["group"] if parent["found"] else p["parent"]),
            "Parent Class": "others" if parent_other else (parent["cls"] if parent["found"] else ""),
            "Child Group": child["group"] if child["found"] else p["child"],
            "Child Class": child["cls"] if child["found"] else "",
        }
        rev = dict(row)
        rev.update(
            {
                "Parent images": "" if parent_other else "\n".join(parent["images"]) if parent["found"] else "",
                "Child images": "\n".join(child["images"]) if child["found"] else "",
                "Child image count": len(child["images"]) if child["found"] else 0,
                "Labeler": p["labeler"],
                "Date": p["date"],
                "Sheet row": p["sheet_row"],
                "Status": "ERROR" if errors else "OK",
            }
        )
        review.append(rev)
        if not errors:
            mapping.append(row)

    # same image in different children
    for u, kids in img_to_children.items():
        if len(kids) > 1:
            issues.append(
                {
                    "Issue": "Same image in multiple child groups",
                    "Severity": "Warning",
                    "Details": "Ek hi image in child groups me hai: " + " | ".join(kids.values()),
                    "Parent group": "", "Child group": " | ".join(kids.values()),
                    "Labeler": "", "Date": "", "Sheet row": "", "Suggestion": u,
                }
            )

    mapping_df = pd.DataFrame(mapping, columns=["Parent group", "Parent Class", "Child Group", "Child Class"])
    review_df = pd.DataFrame(review)
    issues_df = pd.DataFrame(
        issues,
        columns=["Issue", "Severity", "Details", "Parent group", "Child group", "Labeler", "Date", "Sheet row", "Suggestion"],
    )
    if not issues_df.empty:
        issues_df["_o"] = issues_df["Severity"].map({"Error": 0, "Warning": 1})
        issues_df = issues_df.sort_values(["_o", "Issue"]).drop(columns="_o").reset_index(drop=True)

    notin_rows = []
    for e in notin.values():
        notin_rows.append(
            {
                "Group (jaisa file me likha)": e["Group (jaisa file me likha)"],
                "Role": " + ".join(sorted(e["Role"])),
                "Rows affected": e["Rows"],
                "Labeler": ", ".join(sorted(e["Labeler"])),
                "Date": ", ".join(sorted(e["Date"])),
                "Closest CGC match": e["Closest CGC match"],
            }
        )
    notin_df = pd.DataFrame(
        notin_rows,
        columns=["Group (jaisa file me likha)", "Role", "Rows affected", "Labeler", "Date", "Closest CGC match"],
    )

    multi_rows = []
    for res in referenced.values():
        if len(res["images"]) > 1:
            multi_rows.append({"Group": res["group"], "Class": res["cls"], "Image count": len(res["images"]),
                               "Images": "\n".join(res["images"])})
    multi_df = pd.DataFrame(multi_rows, columns=["Group", "Class", "Image count", "Images"])

    pending_rows = []
    if flag_split_pending and cgc.has_images:
        for k, e in cgc.group.items():
            if len(e["images"]) > 1 and k not in parents_used:
                pending_rows.append({"Group": e["name"], "Class": next(iter(e["classes"].values())),
                                     "Image count": len(e["images"]), "Images": "\n".join(e["images"].values())})
    pending_df = pd.DataFrame(pending_rows, columns=["Group", "Class", "Image count", "Images"])

    return dict(mapping=mapping_df, review=review_df, issues=issues_df, notin=notin_df,
                multi=multi_df, pending=pending_df)


# ----------------------------------------------------------------------------
# Excel export
# ----------------------------------------------------------------------------
def build_excel(res) -> bytes:
    out = io.BytesIO()
    errs = int((res["issues"]["Severity"] == "Error").sum()) if not res["issues"].empty else 0
    warns = int((res["issues"]["Severity"] == "Warning").sum()) if not res["issues"].empty else 0
    summary = pd.DataFrame(
        [
            ("Mapping rows (clean)", len(res["mapping"])),
            ("Rows with errors (mapping me nahi gayi)", int((res["review"]["Status"] == "ERROR").sum()) if not res["review"].empty else 0),
            ("Errors", errs),
            ("Warnings", warns),
            ("Groups CGC me nahi mile", len(res["notin"])),
            ("Multi-image groups (used in mapping)", len(res["multi"])),
            ("Multi-image groups jinka split pending", len(res["pending"])),
        ],
        columns=["Check", "Count"],
    )
    sheets = [
        ("Summary", summary),
        ("Mapping", res["mapping"]),
        ("Errors", res["issues"]),
        ("Not in CGC", res["notin"]),
        ("Multi-image groups", res["multi"]),
        ("Split pending", res["pending"]),
        ("Mapping review", res["review"]),
    ]
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        wb = xw.book
        head = wb.add_format({"bold": True, "bg_color": "#D9EAD3", "border": 1, "valign": "top"})
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        for name, df in sheets:
            df.to_excel(xw, sheet_name=name, index=False)
            ws = xw.sheets[name]
            for j, col in enumerate(df.columns):
                ws.write(0, j, col, head)
                longest = max([len(str(col))] + [len(str(v).split("\n")[0]) for v in df[col].astype(str).head(500)])
                ws.set_column(j, j, min(max(12, longest + 2), 60), wrap)
            ws.freeze_panes(1, 0)
            if len(df):
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)
    return out.getvalue()


# ----------------------------------------------------------------------------
# Streamlit UI
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _read(name, data):
    return read_raw(name, data)


def main():
    st.set_page_config(page_title="Mapping Sheet Builder", page_icon="🗂️", layout="wide")
    st.title("🗂️ Mapping Sheet Builder")
    st.caption("Label → Group sheet + CGC (class image ke saath) se Mapping sheet banao aur errors check karo.")

    with st.sidebar:
        st.header("Settings")
        mode = st.radio(
            "Kya karna hai?",
            ["Mapping sheet banao (label → group file se)", "Existing mapping sheet validate karo"],
        )
        split_pending = st.checkbox("CGC ke multi-image groups (jinka split baaki hai) bhi list karo", value=True)
        st.markdown(
            "**Rules**\n"
            "- Parent = group jisse naya group bana\n"
            "- Parent/Child Class = CGC ke class_name se match\n"
            "- Parent `other/others` → class bhi `others`\n"
            "- Error wali rows Mapping me nahi jaati, **Errors** sheet me aati hain"
        )

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1) " + ("Label → Group file" if mode.startswith("Mapping") else "Mapping file"))
        f_main = st.file_uploader("csv / xlsx", type=["csv", "xlsx", "xls"], key="main")
    with col2:
        st.subheader("2) CGC file (class + group + image)")
        f_cgc = st.file_uploader("xlsx / csv", type=["csv", "xlsx", "xls"], key="cgc")

    if not f_main or not f_cgc:
        st.info("Dono files upload karo.")
        return

    main_sheets = _read(f_main.name, f_main.getvalue())
    cgc_sheets = _read(f_cgc.name, f_cgc.getvalue())

    c1, c2 = st.columns(2)
    main_sheet = c1.selectbox("File 1 ki sheet", list(main_sheets), key="ms")
    cgc_sheet = c2.selectbox("CGC ki sheet", list(cgc_sheets), key="cs")
    cgc_rows = cgc_sheets[cgc_sheet]

    hdr_i, header, class_j, group_j, img_js = guess_cgc_columns(cgc_rows)
    with st.expander("CGC columns (auto-detect hue hain, galat ho to badlo)", expanded=class_j is None or not img_js):
        opts = [f"{j}: {h or '(blank)'}" for j, h in enumerate(header)]
        none = "— none —"
        sel_class = st.selectbox("class_name column", [none] + opts, index=(class_j + 1) if class_j is not None else 0)
        sel_group = st.selectbox("group_name column", [none] + opts, index=(group_j + 1) if group_j is not None else 0)
        sel_imgs = st.multiselect("image link column(s)", opts, default=[opts[j] for j in img_js])

    def idx(s):
        return None if s == none else int(s.split(":")[0])

    class_j, group_j = idx(sel_class), idx(sel_group)
    img_js = [int(s.split(":")[0]) for s in sel_imgs]
    if class_j is None and group_j is None:
        st.error("CGC me class_name ya group_name column chuno.")
        return
    if not img_js:
        st.warning("Image column nahi chuna — image checks skip honge.")

    if st.button("▶ Generate", type="primary"):
        with st.spinner("Matching chal raha hai…"):
            cgc = build_cgc(cgc_rows, hdr_i, class_j, group_j, img_js)
            if mode.startswith("Mapping"):
                recs = parse_label_sheet(main_sheets[main_sheet])
                pairs, pre = labels_to_pairs(recs)
                res = analyze(pairs, cgc, pre, False, split_pending)
                res["n_parents"] = len(recs)
            else:
                pairs = parse_mapping_sheet(main_sheets[main_sheet])
                res = analyze(pairs, cgc, [], True, split_pending)
                res["n_parents"] = len({strict(p["parent"]) for p in pairs})
            res["cgc_rows"] = cgc.n_rows
            res["cgc_groups"] = len(cgc.group)
            res["excel"] = build_excel(res)
        st.session_state["res"] = res

    res = st.session_state.get("res")
    if not res:
        return

    errs = int((res["issues"]["Severity"] == "Error").sum()) if not res["issues"].empty else 0
    warns = int((res["issues"]["Severity"] == "Warning").sum()) if not res["issues"].empty else 0
    m = st.columns(5)
    m[0].metric("Mapping rows (clean)", len(res["mapping"]))
    m[1].metric("Errors", errs)
    m[2].metric("Warnings", warns)
    m[3].metric("CGC me nahi mile", len(res["notin"]))
    m[4].metric("Split pending (multi-image)", len(res["pending"]))
    st.caption(f"CGC: {res['cgc_rows']} rows, {res['cgc_groups']} unique groups.")

    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇ Full Excel (Mapping + Errors + extra sheets)",
        res["excel"],
        file_name="mapping_sheet_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    d2.download_button(
        "⬇ Sirf Mapping (CSV)",
        res["mapping"].to_csv(index=False).encode("utf-8-sig"),
        file_name="mapping.csv",
        mime="text/csv",
    )

    tabs = st.tabs(["Mapping", "Errors", "Not in CGC", "Multi-image", "Split pending", "Image review"])
    with tabs[0]:
        st.dataframe(res["mapping"], width="stretch", hide_index=True)
    with tabs[1]:
        if res["issues"].empty:
            st.success("Koi error/warning nahi 🎉")
        else:
            kind = st.multiselect("Filter", sorted(res["issues"]["Issue"].unique()))
            df = res["issues"] if not kind else res["issues"][res["issues"]["Issue"].isin(kind)]
            st.dataframe(df, width="stretch", hide_index=True)
    with tabs[2]:
        st.caption("Ye groups label sheet / mapping me hain lekin CGC me add nahi hue.")
        st.dataframe(res["notin"], width="stretch", hide_index=True)
    with tabs[3]:
        st.caption("Mapping me use hue groups jinme 1 se zyada image hai.")
        st.dataframe(res["multi"], width="stretch", hide_index=True)
    with tabs[4]:
        st.caption("CGC ke multi-image groups jo label sheet me parent ke roop me aaye hi nahi (alag group banana baaki).")
        st.dataframe(res["pending"], width="stretch", hide_index=True)
    with tabs[5]:
        rv = res["review"]
        if rv.empty:
            st.info("Kuch review karne ko nahi.")
        else:
            labels = [f"{i}: {r['Parent group']} → {r['Child Group']} [{r['Status']}]" for i, r in rv.iterrows()]
            pick = st.selectbox("Row chuno", range(len(rv)), format_func=lambda i: labels[i])
            r = rv.iloc[pick]
            a, b = st.columns(2)
            for col, title, key in ((a, "Parent images", "Parent images"), (b, "Child images", "Child images")):
                col.markdown(f"**{title}**")
                urls = [u for u in str(r[key]).split("\n") if u.startswith("http")]
                if not urls:
                    col.caption("— image nahi —")
                for u in urls[:6]:
                    col.image(u, width=180)


if __name__ == "__main__":
    main()
