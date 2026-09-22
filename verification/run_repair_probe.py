#!/usr/bin/env python3
"""表の作り直し（app/fusion.py の repair_tables）を、実 API で測る。

    python run_repair_probe.py --yes --repeats 2 --max-cost 0.4

- **両方の LLM 呼び出しを本物で行う**。①本番の統合分析（Claude 最上位・画像つき）→ ②その応答の `table` だけを
  途中で壊し（形が壊れた場合の再現）→ ③作り直しの呼び出し（本物・文章だけ）→ ④複数シートの Excel。
  壊すのは「AI が表の形を誤った」状況の再現で、呼び出し自体は実 API。
- 確かめること: 実際のモデルが `repair_tables` を呼ぶか／候補は何通り返るか／個人情報の検査を通るか／
  Excel が複数シートになるか／作り直しの追加費用と時間。
- キーは ORCA_API_KEY（表示しない）。素材は verification/assets/（README 参照）。
"""
from __future__ import annotations

import argparse, datetime as dt, io, json, os, pathlib, sys, tempfile, time
from types import SimpleNamespace

HERE = pathlib.Path(__file__).resolve().parent
ROOT = HERE.parent
os.environ.setdefault("MIRUCON_LLM_PROFILE", "orca")
sys.path.insert(0, str(HERE)); sys.path.insert(0, str(ROOT))
from PIL import Image  # noqa: E402
import run_fusion_probe as fp  # noqa: E402
from app import auth, authz, cards, db, fusion, jobs, llm, objects  # noqa: E402
from app.tests.fakes import client_dynamic  # noqa: E402

C = {"reason": "r", "evidence": ""}
OPEN = (940, 660, 596, 364)   # site.jpg（1536×1024 を想定）の右下だけを開ける


def breaking_factory(real_factory, broken: dict):
    """本物のクライアントに、①本番の応答の `table` だけを壊す ②作り直しはそのまま通す、という細工をかぶせる。"""
    def factory(provider):
        client = real_factory(provider)

        def create(**kw):
            raw = client.messages.with_raw_response.create(**kw)
            names = {t["name"] for t in kw.get("tools", [])}
            if "fuse_and_draft" not in names:
                return raw                            # 作り直し・音声は、手を加えない
            parsed = raw.parse()
            for b in parsed.content:
                if getattr(b, "type", "") == "tool_use" and b.name == "fuse_and_draft":
                    b.input = {**b.input, "table": broken}   # 表の形だけを壊す（AI の誤りの再現）
            return SimpleNamespace(headers=raw.headers, parse=lambda: parsed)

        return SimpleNamespace(messages=SimpleNamespace(with_raw_response=SimpleNamespace(create=create)))
    return factory


def main() -> None:
    for s in (sys.stdout, sys.stderr):
        s.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    ap.add_argument("--yes", action="store_true"); ap.add_argument("--repeats", type=int, default=2)
    ap.add_argument("--max-cost", type=float, default=0.4)
    ap.add_argument("--broken", choices=["columns_str", "columns_obj", "rows_empty"], default="columns_str")
    a = ap.parse_args()
    broken = {"columns_str": {"title": "点検項目", "columns": "項目、状態", "rows": [["バケット", "要確認"]]},
              "columns_obj": {"title": "点検項目", "columns": {"1": "項目"}, "rows": []},
              "rows_empty": {"title": "点検項目", "columns": ["項目", "状態"], "rows": []}}[a.broken]
    print(f"壊し方: {a.broken} / 繰り返し {a.repeats} / 上限 ${a.max_cost}。dry-run: {not a.yes}")
    if not a.yes:
        return

    wav = fp.wav_of(fp.VOICE)
    img = Image.open(fp.PHOTO).convert("RGB")
    buf = io.BytesIO(); fp.mcs.reveal_filter(img, [OPEN]).save(buf, "JPEG", quality=88); filtered = buf.getvalue()
    ratio = OPEN[2] * OPEN[3] / (img.width * img.height)

    cfg = llm.load_config()
    cfg["providers"] = [p for p in cfg["providers"] if p in ("orca", "anthropic") and os.environ.get(p.upper() + "_API_KEY")]
    tmpd = pathlib.Path(tempfile.mkdtemp(prefix="repair_")); conn = db.connect(tmpd / "r.db"); db.init(conn)
    db.run(conn, "INSERT INTO org VALUES('org_r','検証','info@example.test',?)", (db.now(),))
    db.run(conn, "INSERT INTO org_setting(org_id, external_llm, talk_llm, talk_audio, vision_llm, fusion_demo) VALUES('org_r',1,1,1,1,1)")
    mid = auth.add_member(conn, "org_r", "maker@example.test", "member"); conn.commit()
    actor = authz.Actor(kind="member", org_id="org_r", member_id=mid, role="member")
    obj = objects.register_object(conn, actor, "検証の現場")[0]
    fake_card = client_dynamic(lambda k, i: [("select_card_type", {**C, "type_id": "maintenance"})] if "select_card_type" in {t["name"] for t in k["tools"]}
                               else [("no_action", C)] if "propose_extra_mask" in {t["name"] for t in k["tools"]}
                               else [("write_card_text", {"title": "t", "changes": ["c"], "description": "d"})], echo_model=True)

    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S"); outd = HERE / "output" / f"repair_{ts}"; outd.mkdir(parents=True, exist_ok=True)
    real = llm.call; tally = {"cost": 0.0, "calls": 0, "repair_calls": 0, "repair_cost": 0.0}

    def counting(conn_, org, kind, system, tools, messages, cf, cfg_):
        names = {t["name"] for t in tools}
        t0 = time.time(); r = real(conn_, org, kind, system, tools, messages, cf, cfg_); el = time.time() - t0
        tally["cost"] += r.cost_usd or 0; tally["calls"] += 1
        if "repair_tables" in names:
            tally["repair_calls"] += 1; tally["repair_cost"] += r.cost_usd or 0
            got = next((t for t in r.tool_uses if t["name"] == "repair_tables"), None)
            n = len(got["input"].get("tables", [])) if got and isinstance(got["input"], dict) else 0
            print(f"    作り直しの呼び出し: {r.model}／{el:.1f}秒／${r.cost_usd or 0:.4f}／AI が返した候補 {n} 件"
                  f"／画像の再送 {'あり' if any(b.get('type') == 'image' for m in messages for b in m['content']) else 'なし'}")
        return r

    llm.call = counting
    rows = []
    try:
        for k in range(a.repeats):
            card = cards.create_card(conn, actor, obj["obj_id"], before_desc="現場の確認", after_desc="打合せ中",
                                     images_in=[("before", filtered, None, "call_screen", ratio)], mask_confirmed=True, client_factory=fake_card)["card"]
            image = db.one(conn, "SELECT * FROM image WHERE card_id=?", (card["card_id"],))
            f = breaking_factory(lambda p: llm.default_factory(p, fusion.FUSION_TIMEOUT), broken)
            c0 = tally["cost"]; t0 = time.time()
            fid = fusion.start(conn, actor, card["card_id"], image["image_id"], wav, "wav", consent=True, jobs=jobs.Inline(), client_factory=f, config=cfg)
            st = fusion.get(conn, actor, fid)
            if st["status"] != "transcribed":
                print(f"[{k + 1}] 文字にできなかった: {st['status']} {st['why']}"); rows.append({"run": k + 1, "audio": st["status"]}); continue
            fusion.confirm_and_analyze(conn, actor, fid, None, jobs=jobs.Inline(), client_factory=f, config=cfg)
            st = fusion.get(conn, actor, fid); el = time.time() - t0
            note = (st["result"] or {}).get("table_note", "")
            print(f"[{k + 1}] {st['status']}／{el:.1f}秒／${tally['cost'] - c0:.4f}／why: {st['why'][:80]}")
            sheets = []
            if st["result"]:
                print(f"    採用された表: {st['result']['table']['title']}／列 {st['result']['table']['columns']}／{len(st['result']['table']['rows'])}行")
                if note:
                    print(f"    {note}")
                for name in st["files"]:
                    (outd / f"run{k + 1}_{name}").write_bytes(fusion.download(conn, actor, fid, name)[1])
                try:
                    import openpyxl, warnings
                    warnings.simplefilter("ignore")
                    sheets = openpyxl.load_workbook(outd / f"run{k + 1}_table.xlsx").sheetnames
                    print(f"    Excel のシート: {sheets}")
                except Exception as e:  # noqa: BLE001
                    print(f"    Excel を開けない: {type(e).__name__}")
            rows.append({"run": k + 1, "status": st["status"], "why": st["why"], "table_note": note, "sheets": sheets,
                         "result": st["result"], "sec": round(el, 1), "cost": round(tally["cost"] - c0, 4)})
            if tally["cost"] > a.max_cost:
                print("費用の上限で中断"); break
    finally:
        llm.call = real
    (outd / "raw.json").write_text(json.dumps({"broken": a.broken, "tally": tally, "runs": rows}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\n実費 ${tally['cost']:.4f}／LLM {tally['calls']} 呼び出し（うち作り直し {tally['repair_calls']} 回・${tally['repair_cost']:.4f}）／保存: {outd}")


if __name__ == "__main__":
    main()
