# -*- coding: utf-8 -*-
"""
spec_intelligence.py - 規格歷史智能分析 (v3.7)

【目的】
傳統做法:抓最近一筆歷史 → 直接帶入 → Sandy 可能漏掉一些「歷史核心要求」
本模組:全面分析同料號歷史,做 3 件事:
  1. 模糊比對料號(ATP3 Rev G == ATP3-Rev-G)
  2. 識別每行的「身份」(核心 / 偶發 / 最新加 / 舊版有最新沒)
  3. 推薦一個「智能合併版」,以最新為基準補回核心要求

【設計重點】
- 純 Python,不依賴 LLM API
- 對 PCB 注意事項這種高度結構化的文字,規則演算法已經很準
- Sandy 看得到「為什麼建議這個」(透明可驗證)

【主要函式】
- normalize_part_no(pn): 標準化料號做模糊比對
- find_similar_part_numbers(target, all_pns, threshold): 找相似料號
- analyze_spec_history(part_number, history_list): 完整分析
- build_smart_spec(analysis): 產生智能推薦版本
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field


# ════════════════════════════════════════════════════════
# 1. 料號模糊比對
# ════════════════════════════════════════════════════════

def normalize_part_no(pn: str) -> str:
    """
    把料號標準化做模糊比對。
    
    範例:
      "ATP3 Rev G"      → "atp3revg"
      "ATP3-Rev-G"      → "atp3revg"   ← 同一個!
      "atp3_rev_g"      → "atp3revg"   ← 同一個!
      "ATP3 Rev G (0800ATP3G-Red)" → "atp3revg0800atp3gred"
      "00006713v1.1 REV 2" → "00006713v11rev2"
    """
    if not pn:
        return ""
    s = str(pn).lower()
    # 拿掉所有空白、連字符、底線、點、括號
    s = re.sub(r"[\s\-_\.\(\)\,\;\:\/\\]+", "", s)
    return s


def is_similar_part_no(pn1: str, pn2: str, strict: bool = True) -> bool:
    """
    判斷兩個料號是否相似(可能是同一個但寫法不同)。
    
    Args:
        strict: True 時只接受「正規化後完全相同」
                False 時接受「一個是另一個的子字串」(更寬鬆)
    """
    n1 = normalize_part_no(pn1)
    n2 = normalize_part_no(pn2)
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    if not strict:
        # 寬鬆:其中一個是另一個的開頭
        # 例如 "ATP3 Rev G" 跟 "ATP3 Rev G (0800ATP3G-Red)"
        shorter, longer = (n1, n2) if len(n1) < len(n2) else (n2, n1)
        if len(shorter) >= 6 and longer.startswith(shorter):
            return True
    return False


def find_similar_part_numbers(target: str, all_pns: list[str], strict: bool = True) -> list[str]:
    """
    在 all_pns 中找出跟 target 相似的所有料號。
    
    回傳:list of 原始料號字串(沒標準化的)
    """
    results = []
    target_norm = normalize_part_no(target)
    if not target_norm:
        return []
    
    for pn in all_pns:
        if pn == target:
            continue  # 跳過自己
        if is_similar_part_no(target, pn, strict=strict):
            results.append(pn)
    return results


# ════════════════════════════════════════════════════════
# 2. 行內容標準化(用於頻率比對)
# ════════════════════════════════════════════════════════

def normalize_line(line: str) -> str:
    """
    把一行文字標準化做比對。
    
    處理:
    - 拿掉前後空白
    - 拿掉開頭的編號(1. 2. 3. 等)
    - 拿掉重複空白
    - 全形/半形混用統一
    - 大小寫不變(很多 PCB 詞彙大小寫有意義)
    """
    if not line:
        return ""
    s = line.strip()
    # 拿掉開頭編號
    s = re.sub(r"^\d+[\.\s]+", "", s)
    s = re.sub(r"^[•\-\*]\s*", "", s)
    # 拿掉開頭的 tab 跟雜訊
    s = re.sub(r"^[\t\s]+", "", s)
    # 多重空白變單一
    s = re.sub(r"\s+", " ", s)
    # 統一全形括號
    s = s.replace("(", "(").replace(")", ")")
    # 統一全形分號逗號
    s = s.replace(";", ";").replace(",", ",")
    return s.strip()


def lines_are_similar(line1: str, line2: str, threshold: float = 0.85) -> bool:
    """判斷兩行文字內容是否實質相同(允許小差異)"""
    n1 = normalize_line(line1)
    n2 = normalize_line(line2)
    if not n1 or not n2:
        return False
    if n1 == n2:
        return True
    # SequenceMatcher 處理小差異
    from difflib import SequenceMatcher
    ratio = SequenceMatcher(None, n1, n2).ratio()
    return ratio >= threshold


# ════════════════════════════════════════════════════════
# 3. 歷史分析核心
# ════════════════════════════════════════════════════════

@dataclass
class AnnotatedLine:
    """單行規格 + 識別標籤"""
    text: str               # 原始行(保留格式)
    normalized: str         # 標準化後(用於比對)
    occurrences: int        # 在歷史中出現幾次
    total_history: int      # 該料號歷史總筆數
    in_latest: bool         # 是否在最新一筆
    in_history_pos: list[int] = field(default_factory=list)  # 在歷史第幾筆出現 (0=最新)
    
    @property
    def category(self) -> str:
        """
        分類:
          'CORE'    = 每筆都有 (核心要求,絕不能漏)
          'COMMON'  = 大部分有 (>=50%)
          'NEW'     = 只在最新有 (新加的要求)
          'DROPPED' = 舊版有但最新沒 (可能過時,要警告)
          'OCCASIONAL' = 偶爾才有 (<50% 且不在最新)
        """
        if self.total_history == 0:
            return "UNKNOWN"
        ratio = self.occurrences / self.total_history
        if self.in_latest:
            if ratio == 1.0:
                return "CORE"
            elif ratio >= 0.5:
                return "COMMON"
            else:
                return "NEW"  # 在最新但歷史少見
        else:
            # 不在最新
            if ratio >= 0.5:
                return "DROPPED"  # 舊版多數有,最新沒 → 警告
            else:
                return "OCCASIONAL"  # 偶爾才出現,不重要
    
    @property
    def color_code(self) -> str:
        """UI 用的顏色代碼"""
        return {
            "CORE": "🟢",
            "COMMON": "⚪",
            "NEW": "🔵",
            "DROPPED": "🔴",
            "OCCASIONAL": "🟡",
            "UNKNOWN": "⚫",
        }.get(self.category, "⚫")
    
    @property
    def explanation(self) -> str:
        """給 Sandy 看的解釋"""
        ratio_pct = int(self.occurrences / self.total_history * 100) if self.total_history else 0
        if self.category == "CORE":
            return f"核心要求 (歷史 {self.occurrences}/{self.total_history} 都有)"
        elif self.category == "COMMON":
            return f"常見要求 ({ratio_pct}% 訂單有)"
        elif self.category == "NEW":
            return f"最新加的 (只在最新一筆出現)"
        elif self.category == "DROPPED":
            return f"⚠️ 舊版有但最新沒 ({self.occurrences}/{self.total_history} 訂單有,可能過時)"
        elif self.category == "OCCASIONAL":
            return f"偶發 ({self.occurrences} 筆出現過)"
        return ""


@dataclass
class SpecAnalysis:
    """完整分析結果"""
    part_number: str
    total_history: int
    history_records: list[dict]  # 排序後的歷史(最新在前)
    latest_record: dict
    annotated_lines: list[AnnotatedLine]  # 所有分析過的行(來自最新 + 舊版獨有)
    similar_part_numbers: list[str]  # 模糊比對找到的相似料號
    
    @property
    def has_warnings(self) -> bool:
        """是否有需要 Sandy 注意的疑點"""
        return any(l.category == "DROPPED" for l in self.annotated_lines)
    
    @property
    def core_lines(self) -> list[AnnotatedLine]:
        return [l for l in self.annotated_lines if l.category == "CORE"]
    
    @property
    def new_lines(self) -> list[AnnotatedLine]:
        return [l for l in self.annotated_lines if l.category == "NEW"]
    
    @property
    def dropped_lines(self) -> list[AnnotatedLine]:
        return [l for l in self.annotated_lines if l.category == "DROPPED"]


def analyze_spec_history(
    part_number: str,
    history_records: list[dict],
) -> SpecAnalysis:
    """
    分析同料號的所有歷史規格,產出完整分析結果。
    
    Args:
        part_number: 料號(顯示用)
        history_records: list of {
            "po": str,
            "date": str (YYYY-MM-DD),
            "factory": str,
            "spec_text": str,
        }
    
    回傳:SpecAnalysis 物件
    """
    if not history_records:
        return SpecAnalysis(
            part_number=part_number,
            total_history=0,
            history_records=[],
            latest_record={},
            annotated_lines=[],
            similar_part_numbers=[],
        )
    
    # 1. 排序:最新在前
    sorted_history = sorted(
        history_records,
        key=lambda h: h.get("date", ""),
        reverse=True,
    )
    latest = sorted_history[0]
    total = len(sorted_history)
    
    # 2. 蒐集所有 (歷史筆數 idx, 行) 對
    # 為了「行匹配」,我們合併相似的行(避免 1.\tS/M 跟 1. S/M 被當不同)
    # 用 normalized 字串當 key
    
    # 先建立每筆歷史的 normalized lines
    history_norm_lines = []
    for h in sorted_history:
        spec = h.get("spec_text", "")
        lines = [l for l in spec.split("\n") if l.strip()]
        norms = [normalize_line(l) for l in lines]
        history_norm_lines.append((lines, norms))
    
    # 3. 統計每個 normalized line 出現幾次,以及在哪幾筆
    # 「相似」的行視為同一個 key (但這裡先用嚴格相同,效率優先)
    line_index = {}  # norm → {original_text, count, positions[]}
    
    for hist_idx, (raw_lines, norms) in enumerate(history_norm_lines):
        seen_in_this_hist = set()  # 避免同一張單某行重複算
        for raw, norm in zip(raw_lines, norms):
            if not norm or norm in seen_in_this_hist:
                continue
            seen_in_this_hist.add(norm)
            
            # 找有沒有相似的已存在 key
            matched_key = None
            for existing_norm in line_index.keys():
                if lines_are_similar(norm, existing_norm, threshold=0.85):
                    matched_key = existing_norm
                    break
            
            if matched_key is None:
                line_index[norm] = {
                    "raw": raw,
                    "count": 1,
                    "positions": [hist_idx],
                    "in_latest": (hist_idx == 0),
                }
            else:
                line_index[matched_key]["count"] += 1
                line_index[matched_key]["positions"].append(hist_idx)
                if hist_idx == 0:
                    line_index[matched_key]["in_latest"] = True
                    line_index[matched_key]["raw"] = raw  # 用最新的版本當 raw
    
    # 4. 產出 AnnotatedLine list
    # 先放最新一筆有的行(順序按最新一筆),再加 latest 沒有但歷史多筆有的(DROPPED)
    annotated = []
    latest_norms = set(history_norm_lines[0][1])  # 最新筆的 norm 集合
    
    for raw, norm in zip(*history_norm_lines[0]):  # 按最新一筆順序
        if not norm:
            continue
        # 找對應的 line_index entry
        for k, v in line_index.items():
            if lines_are_similar(norm, k, threshold=0.85):
                annotated.append(AnnotatedLine(
                    text=raw,
                    normalized=norm,
                    occurrences=v["count"],
                    total_history=total,
                    in_latest=True,
                    in_history_pos=v["positions"],
                ))
                break
    
    # 加入 latest 沒有但歷史多筆有的(DROPPED)
    annotated_norms = {a.normalized for a in annotated}
    for k, v in line_index.items():
        # 已經放過的跳過
        if any(lines_are_similar(k, an, threshold=0.85) for an in annotated_norms):
            continue
        # 不在最新 + 至少 2 筆有 → DROPPED 警告
        if not v["in_latest"] and v["count"] >= 2:
            annotated.append(AnnotatedLine(
                text=v["raw"],
                normalized=k,
                occurrences=v["count"],
                total_history=total,
                in_latest=False,
                in_history_pos=v["positions"],
            ))
    
    return SpecAnalysis(
        part_number=part_number,
        total_history=total,
        history_records=sorted_history,
        latest_record=latest,
        annotated_lines=annotated,
        similar_part_numbers=[],  # 由呼叫者填(需要全部 spec_history)
    )


# ════════════════════════════════════════════════════════
# 4. 智能合併版本生成器
# ════════════════════════════════════════════════════════

def build_smart_spec(analysis: SpecAnalysis) -> str:
    """
    根據分析結果產生「智能合併版本」。
    
    策略(根據用戶 Q3 選擇:最近一筆最重要):
      1. 以最新一筆當骨架(順序、措辭都用最新的)
      2. 把 DROPPED 的條目「以最新的格式」加回去 — 但加問號標記
      3. CORE / COMMON / NEW 全部保留
    """
    if not analysis.history_records:
        return ""
    
    # 直接用最新一筆 spec_text 當基底
    latest_spec = analysis.latest_record.get("spec_text", "")
    
    # 看有沒有 DROPPED(舊版有但最新沒)
    dropped = analysis.dropped_lines
    if not dropped:
        return latest_spec  # 沒疑點,直接用最新
    
    # 有 DROPPED → 在最新版本後面加上「⚠️ 舊版有但最新沒」區塊
    smart_spec = latest_spec.rstrip()
    smart_spec += "\n\n--- 以下為舊版本曾有的條目,請確認是否要保留 ---\n"
    for d in dropped:
        smart_spec += f"? {d.text.strip()}   [{d.explanation}]\n"
    
    return smart_spec


def get_history_summary_text(analysis: SpecAnalysis) -> str:
    """產生 UI 顯示用的歷史分析摘要(純文字版)"""
    if analysis.total_history == 0:
        return ""
    
    lines = [f"📊 找到 {analysis.total_history} 筆同料號歷史"]
    
    if analysis.core_lines:
        lines.append("")
        lines.append(f"🟢 核心要求 (每筆都有,共 {len(analysis.core_lines)} 條):")
        for al in analysis.core_lines:
            lines.append(f"  • {al.text.strip()[:60]}")
    
    if analysis.new_lines:
        lines.append("")
        lines.append(f"🔵 最新加的 (只在最新一筆,共 {len(analysis.new_lines)} 條):")
        for al in analysis.new_lines:
            lines.append(f"  • {al.text.strip()[:60]}")
    
    if analysis.dropped_lines:
        lines.append("")
        lines.append(f"🔴 ⚠️ 舊版有但最新沒 (共 {len(analysis.dropped_lines)} 條,請確認):")
        for al in analysis.dropped_lines:
            lines.append(f"  • {al.text.strip()[:60]}  ({al.occurrences}/{analysis.total_history} 筆有)")
    
    return "\n".join(lines)
