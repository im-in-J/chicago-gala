"""자료 폴더(C:\\Users\\user\\cicago) → data.json + assets + index.html 내장 데이터.

    python tools/import_materials.py <자료 폴더> [pymupdf 경로]

큐시트(xlsx)나 동선(pdf)이 바뀌면 다시 실행한다.
"""
import json, os, re, shutil, sys, zipfile
import xml.etree.ElementTree as E

SRC = sys.argv[1]
if len(sys.argv) > 2:
    sys.path.insert(0, sys.argv[2])
import pymupdf

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SONGS = [
    ("all-that-jazz", "All That Jazz", "ATJ", "Jazz", "#A8262F", "#F0787E", "All That Jazz MR2.mp3",
     "젠 - 핀마이크, 듀이·아라레 - 스탠딩마이크"),
    ("cell-block-tango", "Cell Block Tango", "CBT", "Cell", "#9A6F12", "#E3B341", "Cell Block Tango MR2.mp3",
     "캐시 - 핀마이크, 랄다 - 핸드마이크, 에디 - 핸드마이크 or 스탠딩 마이크"),
    ("reached-for-the-gun", "We Both Reached for the Gun", "Gun", "Gun", "#3E7C8F", "#6FB3C6", None,
     "랄다·에디·디나·동주 - 스탠딩마이크, 대니·젠 - 핀마이크"),
    ("roxie", "Roxie", "락시", "락시", "#7A4FA3", "#B79AE0", "Roxie MR.mp3",
     "여림 - 스탠딩마이크, 솜 - 핀마이크, 슈 - 스탠딩 핸드마이크"),
]
CUE_SONG = {"All That Jazz": 0, "Cell Block Tango": 1, "We Both Reached for the Gun": 2, "락시": 3}
SESSIONS = [
    ("2026-10-16", "19:00", "마곡 (버즈 X)", ["7:00~ Gun 안무, 동선 (30분)", "7:40~ CBT 안무, 동선 (80분)"]),
    ("2026-10-23", "19:00", "신논현", ["7:00~ ATJ 안무, 동선 (80분)", "8:30~ 락시 안무, 동선 (30분)"]),
    ("2026-10-30", "18:00", "마곡", ["6:00~ Gun / 버즈 슈 에디 랄다", "6:30~ CBT / 캐시 에디 랄다 슈", "7:00~ Gun 씬", "8:00~ CBT 씬"]),
    ("2026-11-06", "18:00", "신논현 (+1시간)", ["6:00~ 락시 / 솜 여림 슈", "6:30~ ATJ / 젠 듀이 아라레", "7:00~ ATJ 씬", "8:00~ 락시 씬", "9:00~ (추가)"]),
    ("2026-11-13", "18:30", "마곡 (삭제)", ["6:30~ 락시 / 에디 캐시 아라레 랄다 솜 슈", "7:00~ Gun, ATJ", "8:00~ CBT, 락시"]),
    ("2026-11-20", "", "마곡 (TBD)", []),
    ("2026-11-27", "", "신논현 (TBD)", []),
    ("2026-12-04", "", "마곡 (TBD) +1시간", []),
]
CAST = ["젠", "솜", "아라레", "랄다", "슈", "듀이", "캐시", "여림", "동주", "대니", "에디", "디나"]


def hexcolor(rgb):
    return "#%02X%02X%02X" % tuple(round(c * 255) for c in rgb)


def read_blocking(path):
    doc = pymupdf.open(path)
    members, scenes, by_color = [], [], {}
    for pn, page in enumerate(doc, 1):
        words = page.get_text("words")
        drawings = page.get_drawings()
        stage = next(d["rect"] for d in drawings if d.get("fill") and d["rect"].width > 700)
        dots = [d for d in drawings if d.get("fill") and d.get("fill_opacity") == 1.0 and 25 < d["rect"].width < 32]
        assert len(dots) == 13, (pn, len(dots))
        if not by_color:  # 1쪽의 원 안 숫자로 색↔번호를 정하고, 이후 쪽은 색으로 찾는다 (글자가 붙어 나오는 쪽이 있음)
            for d in dots:
                c = d["rect"]
                nums = [w[4] for w in words if w[4].isdigit() and c.x0 <= (w[0] + w[2]) / 2 <= c.x1 and c.y0 <= (w[1] + w[3]) / 2 <= c.y1]
                assert len(nums) == 1, (pn, nums)
                by_color[hexcolor(d["fill"])] = int(nums[0])
            members = [{"id": "m%d" % no, "no": no, "name": CAST[no - 1], "role": "",
                        "color": next(k for k, v in by_color.items() if v == no)} for no in range(1, 13)]
        spots = {}
        for d in dots:
            c = d["rect"]
            spots[by_color[hexcolor(d["fill"])]] = {
                "x": round(((c.x0 + c.x1) / 2 - stage.x0) / stage.width * 100, 1),
                "y": round(((c.y0 + c.y1) / 2 - stage.y0) / stage.height * 100, 1),
            }
        assert sorted(spots) == list(range(1, 14)), pn
        title = " ".join(w[4] for w in sorted(words, key=lambda w: w[0]) if w[3] < stage.y0 - 15)
        scenes.append({
            "page": pn, "title": title,
            "positions": {m["id"]: [spots[m["no"]]["x"], spots[m["no"]]["y"]] for m in members},
            "mic": [spots[13]["x"], spots[13]["y"]],
        })
    return members, scenes


def read_cues(path):
    M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    z = zipfile.ZipFile(path)
    shared = ["".join(t.text or "" for t in si.iter(M + "t")) for si in E.fromstring(z.read("xl/sharedStrings.xml"))]
    rows = []
    for r in E.fromstring(z.read("xl/worksheets/sheet1.xml")).iter(M + "row"):
        o = {}
        for c in r:
            v, t = c.find(M + "v"), c.get("t")
            if t == "inlineStr":
                x = "".join(q.text or "" for q in c.iter(M + "t"))
            elif v is None:
                continue
            else:
                x = shared[int(v.text)] if t == "s" else v.text
            o[re.sub(r"\d+", "", c.get("r"))] = x.strip()
        rows.append(o)
    cues, last = [], None
    for o in rows[1:]:
        if not o:
            continue
        idx = CUE_SONG.get(o.get("A"), last)  # 곡명이 빠진 메모 행은 앞 곡에 붙인다
        last = idx
        cues.append({
            "songId": SONGS[idx][0], "part": re.sub(r"\s*\n\s*", " ", o.get("B", "")), "type": o.get("C", ""),
            "mic": o.get("D", ""), "text": o.get("E", ""), "memo": o.get("F", ""), "setting": o.get("G", ""),
        })
    return cues


def lyrics_for(song_id, cues):
    out, prev = [], None
    for c in cues:
        if c["songId"] != song_id or not c["text"]:
            continue
        part = c["part"].strip("[]") or c["type"]
        if part != prev:
            if out:
                out.append("")
            out.append("[%s]" % part)
            prev = part
        out.append("(%s)" % c["text"] if c["type"] == "진행 큐" else c["text"])
    return "\n".join(out)


def size(p):
    n = os.path.getsize(p)
    return "%.1f MB" % (n / 1048576) if n > 1048576 else "%d KB" % (n / 1024)


def main():
    members, pages = read_blocking(os.path.join(SRC, "Chicago 동선.pdf"))
    cues = read_cues(os.path.join(SRC, "Chicago_musical_gala_cue_sheet.xlsx"))
    assets = os.path.join(ROOT, "assets")

    def put(name, rel):
        dst = os.path.join(assets, rel)
        os.makedirs(os.path.dirname(dst), exist_ok=True)
        shutil.copyfile(os.path.join(SRC, name), dst)
        return "assets/" + rel.replace(os.sep, "/"), size(dst)

    songs, scenes, files = [], [], []
    for sid, title, short, prefix, c, cd, mr, mic in SONGS:
        songs.append({"id": sid, "title": title, "short": short, "slug": sid, "color": c, "colorDark": cd,
                      "micNote": mic, "lyrics": lyrics_for(sid, cues)})
        for p in pages:
            if p["title"].split()[0] != prefix:
                continue
            scenes.append({
                "id": "c%d" % p["page"], "songId": sid, "label": p["title"].split(" ", 1)[1],
                "note": "동선 PDF %d쪽" % p["page"], "positions": p["positions"],
                "props": [{"id": "p%d" % p["page"], "type": "mic", "x": p["mic"][0], "y": p["mic"][1]}],
            })
        if mr:
            path, sz = put(mr, "%s/mr.mp3" % sid)
            files.append({"name": mr.rsplit(".", 1)[0], "path": path, "size": sz, "kind": "audio", "songId": sid, "group": "audio"})
    assert len(scenes) == len(pages) == 34, (len(scenes), [p["title"] for p in pages])
    pdf, pdf_sz = put("Chicago 동선.pdf", "chicago-blocking.pdf")
    xlsx, xlsx_sz = put("Chicago_musical_gala_cue_sheet.xlsx", "chicago-cue-sheet.xlsx")
    files.append({"name": "Chicago 동선 (PDF)", "path": pdf, "size": pdf_sz, "kind": "pdf", "group": "score"})
    files.append({"name": "큐시트 원본 (XLSX)", "path": xlsx, "size": xlsx_sz, "kind": "doc", "group": "etc"})
    data = {
        "show": {"eyebrow": "CHICAGO GALA · STAGE BINDER V2", "title": "시카고 갈라반", "dateISO": "", "venue": ""},
        "members": members, "songs": songs, "scenes": scenes, "cues": cues,
        "scoreDocs": [{"id": "d1", "title": "Chicago 동선", "path": pdf}], "files": files,
        "sessions": [{"id": "t%d" % (i + 1), "dateISO": d, "start": st, "end": "", "concept": con, "plan": "\n".join(pl), "off": False}
                     for i, (d, st, con, pl) in enumerate(SESSIONS)],
        "people": [{"id": "p%d" % m["no"], "name": m["name"], "color": m["color"]} for m in members],
    }
    with open(os.path.join(ROOT, "data.json"), "w", encoding="utf8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    html_path = os.path.join(ROOT, "index.html")
    html = open(html_path, encoding="utf8", newline="").read()
    blob = json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")
    html, n = re.subn(r'(<script type="application/json" id="app-state">).*?(</script>)',
                      lambda m: m.group(1) + blob + m.group(2), html, count=1, flags=re.S)
    assert n == 1
    open(html_path, "w", encoding="utf8", newline="").write(html)
    sys.stdout.reconfigure(encoding="utf8")
    print("members", len(members), "scenes", len(scenes), "cues", len(cues), "files", len(files), "sessions", len(data["sessions"]))
    for s in songs:
        print(s["short"], "scenes", sum(1 for c in scenes if c["songId"] == s["id"]),
              "cues", sum(1 for c in cues if c["songId"] == s["id"]), "lyric lines", s["lyrics"].count("\n") + 1)
    print({m["name"]: m["color"] for m in members})
    print([p["title"] for p in pages])


main()
