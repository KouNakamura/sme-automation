# -*- coding: utf-8 -*-
"""
サイボウズの【出張】イベントをホテル名なしタイトルに修正
（ANA/Hilton不要・既知予約データを直接指定）
"""
import importlib.util, sys

# check_travel.pyをモジュールとして読み込む
spec = importlib.util.spec_from_file_location("ct", r"C:\Users\nakko\Scripts\check_travel.py")
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

# 既知の予約データ（直接指定）
from datetime import date
KNOWN_HOTELS = [
    {'hotel': 'Hilton Tokyo',  'checkin': date(2026,5,23), 'checkout': date(2026,5,24), 'confirmation': '3451900389'},
    {'hotel': 'Conrad Tokyo',  'checkin': date(2026,6,12), 'checkout': date(2026,6,14), 'confirmation': '3452878855'},
]

print("=== サイボウズ【出張】タイトル修正 ===")
print(f"対象: {len(KNOWN_HOTELS)}件")
for h in KNOWN_HOTELS:
    print(f"  {h['checkin']}〜{h['checkout']} {h['hotel']}")

synced, skipped = mod.sync_cybozu_schedule([], KNOWN_HOTELS)

print(f"\n結果: 登録 {len(synced)}件")
for s in synced:
    print(f"  ✓ {s}")
