#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
從學術 PDF（中/英）中擷取「摘要 / Abstract」區段，並輸出：

- 每個 PDF 一個純文字檔（含 題目/年分/作者/摘要 四段），存放於輸出資料夾下的 txt 子資料夾
- 另產生彙整 CSV（論文整理.csv），欄位：題目、年分、作者、摘要；位於輸出資料夾根目錄

使用方式（Usage）:
	python pdf_abstract.py <input_path> [--output-dir OUTPUT_DIR] [--recursive] <-v>

- <input_path> 可為單一 PDF 檔，或包含多個 PDF 的資料夾；若為資料夾，加入 --recursive 會遞迴掃描子資料夾。
- 產出：每個 PDF 會對應一個 .txt 檔案（在 <output-dir>/txt/ 內），另有彙整 CSV 在 <output-dir>/。

擷取啟發式（Heuristics）：
- 起始偵測：支援「摘要」、「中文摘要」、「Abstract」、「ABSTRACT」，允許後方有中英文冒號（：或:）。
- 內嵌形式：支援一行同時包含『標題＋內容』的情況，例如「摘要：本研究…」。
- 結束偵測：遇到「關鍵字/关键词/Keywords」、「References/參考文獻/参考文献」、
	或常見第一章標題（Introduction/引言/緒論/編號章節）即視為摘要結束。
- 支援跨頁摘要：為避免誤擷取，最多延伸到數頁（可調整）。
- 摘要文字會以單一段落輸出（去除換行），便於閱讀與彙整。

其他處理：
- 題目優先由檔名推斷，並移除如「基礎1_ / 標準10_ / 查詢3_」等目錄型前綴與尾端日期。
- 若 CSV 檔案正在被開啟而無法覆寫，程式會提出警告並略過寫入；請關閉後重新執行以更新 `論文整理.csv`。

相依套件：PyMuPDF（pymupdf）
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
import csv
import io
import time
from pathlib import Path
from typing import Iterable, List, Optional, Sequence, Tuple

try:
	import fitz  # PyMuPDF
except Exception as e:  # pragma: no cover - import error surfaced at runtime
	fitz = None  # type: ignore


START_PATTERNS = [
	# 起始樣式 1：單獨一行的 摘要/Abstract 標題
	r"^\s*(?:中文)?摘要\s*[：:]?\s*$",  # 摘要 header on its own line
	# 起始樣式 2：同一行內含「摘要：內容…」的情況（內嵌摘要第一句）
	r"^\s*摘要\s*[：:]\s*(?P<inline>.+)$",  # 摘要： inline
	r"^\s*中文摘要\s*[：:]\s*(?P<inline>.+)$",
	r"^\s*ABSTRACT\s*[：:]?\s*$",
	r"^\s*Abstract\s*[：:]?\s*$",
	r"^\s*Abstract\s*[：:]\s*(?P<inline>.+)$",
]

# Common section markers indicating the end of abstract
END_PATTERNS = [
	# 常見摘要結束：關鍵字 / 关键词 / Keywords
	r"^\s*(?:關鍵[詞字]|关键词)\s*[：:]?\b",  # 關鍵字：/关键词：/關鍵字/关键词
	r"^\s*(?:Keywords?|KEYWORDS)\b.*",  # Keywords/KEYWORDS
	r"^\s*Index\s+Terms\b.*",  # IEEE 常見關鍵詞標題
	# 其他常見會結束摘要的區塊
	r"^\s*(?:目錄|目次)\s*$",
	r"^\s*(?:參考文獻|参考文献|References)\b",
	r"^\s*(?:致謝|誌謝)\b",
	r"^\s*(?:引言|緒論)\b",
	r"^\s*Introduction\b",
	# Numbered sections like 1. Introduction, 2. Method, etc.
	r"^\s*\d+\s*[\.|、]\s*\S+",
	# Chinese numbered sections like 一、二、三、
	r"^\s*[一二三四五六七八九十]+、",
]

# Safety limit for maximum number of pages to include in abstract extraction
MAX_ABSTRACT_PAGES = 3


@dataclass
class ExtractionResult:
	# 擷取結果模型：包含摘要文字與相關中繼資料（起訖頁、觸發的標記與原因）
	abstract: Optional[str]
	start_page: Optional[int]
	end_page: Optional[int]
	start_marker: Optional[str]
	end_marker: Optional[str]
	reason: Optional[str] = None


def iter_pdf_lines(pdf_path: Path) -> Iterable[Tuple[int, str]]:
	"""以 PyMuPDF 讀取 PDF，每行回傳 (page_index, line)。

	- page_index 為 0 起算。
	- 盡量保留原始換行，利於後續以「行」為單位做模式偵測。
	"""
	if fitz is None:
		raise RuntimeError(
			"PyMuPDF (pymupdf) is not installed. Please install it: pip install pymupdf"
		)
	doc = fitz.open(pdf_path)
	try:
		for i in range(len(doc)):
			page = doc[i]
			text = page.get_text("text")  # 以純文字方式取回（包含換行）
			# Normalize Windows and Mac linebreaks just in case
			text = text.replace("\r\n", "\n").replace("\r", "\n")
			for line in text.split("\n"):
				# 移除行尾空白；保留行首縮排以利偵測標題格式
				yield i, line.rstrip()
	finally:
		doc.close()


def _get_pdf_metadata(pdf_path: Path) -> dict:
	"""讀取 PDF 中的 metadata（title/author/creation date 等）。"""
	if fitz is None:
		raise RuntimeError(
			"PyMuPDF (pymupdf) is not installed. Please install it: pip install pymupdf"
		)
	doc = fitz.open(pdf_path)
	try:
		return dict(doc.metadata or {})
	finally:
		doc.close()


def _sanitize_meta_value(value: Optional[str]) -> Optional[str]:
	if not value:
		return None
	v = value.strip()
	if not v:
		return None
	lower = v.lower()
	if lower in {"unknown", "untitled", "title", "author", "null", "none"}:
		return None
	return v


def _parse_year_from_string(s: str) -> Optional[str]:
	"""從字串找出 1900-2099 的西元年。回傳四位數字字串。"""
	m = re.search(r"\b(19|20)(\d{2})\b", s)
	if m:
		return m.group(1) + m.group(2)
	return None


def _parse_year_from_filename(stem: str) -> Optional[str]:
	# 常見格式：_YYYYMMDD 或 (YYYY) 或 [YYYY] 或 _YYYY
	patterns = [
		r"(?:_|\()((?:19|20)\d{2})(?:\d{4})?\)?$",
		r"\[((?:19|20)\d{2})\]$",
		r"((?:19|20)\d{2})$",
	]
	for p in patterns:
		m = re.search(p, stem)
		if m:
			return m.group(1)
	# 任意位置的第一個四位年份
	return _parse_year_from_string(stem)


def _parse_title_from_filename(stem: str) -> str:
	# 去除前綴如 基礎1_ / 標準10_ / 查詢3_ 等目錄型標籤
	s = re.sub(r"^(?:基礎|標準|查詢)\d+_", "", stem)
	# 去除尾端日期/年份如 _20220219 或 _2022
	s = re.sub(r"_(?:19|20)\d{2}(?:\d{4})?$", "", s)
	# 將底線移除（檔名分隔符），避免影響中文字間距
	s = s.replace("_", "")
	return s.strip()


def _normalize_title(title: str) -> str:
	"""標題正規化：
	- 移除檔名常見的末尾標記：最終版/定稿/final/版本號等
	- 移除出現在中英文邊界、僅作為分隔符的連字號（保留英文詞內部的連字號，如 RBF-UNet）
	- 收斂多重空白
	"""
	t = title.strip()
	# 將內部換行轉為空白，避免題目被分成多行
	t = re.sub(r"[\r\n]+", " ", t)
	# 去除尾綴：最終版/定稿/final 及版本號
	t = re.sub(r"[\s\-_/·]*\(?(?:最終版|最終|定稿|定版|final|FINAL|Final|修訂版|修正版|v\d{1,3})\)?\s*$", "", t)
	# 去除 CJK 與 ASCII 之間僅作分隔的連字號（不影響英文詞內的連字號）
	t = re.sub(fr"([{_CJK}])\s*-\s*", r"\1", t)
	t = re.sub(fr"\s*-\s*([{_CJK}])", r"\1", t)
	# 在 CJK 與 ASCII/數字 之間插入單一空白，提升可讀性
	t = _normalize_cjk_ascii_spacing(t, mode="insert")
	# 收斂多重空白
	t = re.sub(r"\s{2,}", " ", t)
	return t


def _first_n_page_lines(pdf_path: Path, n_pages: int = 3, max_lines: int = 300) -> List[str]:
	"""擷取前 n_pages 的文字行（合併），最多回傳 max_lines 行。"""
	out: List[str] = []
	for pg, line in iter_pdf_lines(pdf_path):
		if pg >= n_pages:
			break
		out.append(line)
		if len(out) >= max_lines:
			break
	return out


def _has_cjk(text: str) -> bool:
	"""檢查字串是否包含中日韓文字。"""
	return re.search(f"[{_CJK}]", text) is not None


def _find_abstract_header_index(lines: List[str]) -> Optional[int]:
	"""在行列中尋找 'Abstract' 標題所在的索引。大小寫不敏感。"""
	for i, ln in enumerate(lines):
		s = ln.strip()
		if re.match(r"^(?:Abstract|ABSTRACT)\b\s*:?.*$", s, re.IGNORECASE):
			return i
	return None


def _looks_like_affiliation(line: str) -> bool:
	"""粗略判斷是否為作者單位/聯絡資訊行，以排除在英文題目擷取中誤納。"""
	s = line.strip()
	if not s:
		return True  # 空行視為分隔
	patterns = [
		r"@", r"Department", r"School", r"University", r"College", r"Faculty",
		r"Institute", r"Laborator(?:y|ies)", r"Centre|Center", r"Affiliation",
		r"Corresponding\s+author", r"Article\s+history", r"Received", r"Accepted",
		r"Published", r"ORCID", r"Keywords?", r"KEYWORDS", r"Index\s+Terms",
	]
	if re.search("|".join(patterns), s, re.IGNORECASE):
		return True
	# 多個逗號的小短行，常見於作者列表或單位地址
	if s.count(",") >= 2 and len(s) <= 120:
		return True
	# 疑似頁首頁尾
	if len(s) <= 6 and s.isupper():
		return True
	return False


def _extract_english_title_from_lines(lines: List[str]) -> Optional[str]:
	"""嘗試於英文論文中，在 'Abstract' 之前擷取多行標題並串接為單行。

	策略：自文件開頭至 Abstract 標題之前，挑選前段連續且不像作者/單位的行，
	取 1~4 行合併為題目；合併時以空白連接並清理多餘空白與連字號換行。
	"""
	# 限制於前兩頁的內容
	head = lines[:200]
	abs_idx = _find_abstract_header_index(head)
	if abs_idx is None:
		window = head
	else:
		window = head[:abs_idx]

	# 移除開頭空白行
	i = 0
	while i < len(window) and not window[i].strip():
		i += 1
	window = window[i:]
	if not window:
		return None

	# 收集連續的 1~4 行作為候選題目，排除疑似單位/作者資訊
	candidate_lines: List[str] = []
	for ln in window:
		if _looks_like_affiliation(ln):
			if candidate_lines:
				break
			else:
				continue
		# 實際內容行
		s = ln.strip()
		candidate_lines.append(s)
		if len(candidate_lines) >= 4:
			break

	if not candidate_lines:
		return None

	# 合併為單行題目，簡單處理英文換行
	merged: List[str] = []
	for ln in candidate_lines:
		if merged and merged[-1].endswith('-') and ln[:1].islower():
			merged[-1] = merged[-1][:-1] + ln.lstrip()
		else:
			merged.append(ln)
	title = " ".join(merged)
	# 清理多重空白
	title = re.sub(r"\s{2,}", " ", title).strip()
	# 題目基本合理性檢查
	if len(title) < 15 or len(title.split()) < 4:
		return None
	return title


def _extract_chinese_title_from_lines(lines: List[str]) -> Optional[str]:
	"""在『摘要/Abstract』之前嘗試擷取中文題目。

	策略：
	- 只看前兩頁（~200 行）且在『摘要/Abstract』標題之前的區塊
	- 跳過疑似學校/系所/指導/致謝/目錄 等標籤行
	- 額外排除敘述性句子（含大量逗號/句號/括號/數字、或以『本研究/本文/本論文』等起首）
	- 選擇最像中文題目的那一行（包含 CJK、長度 8~40、少標點、末尾非句號/冒號）
	"""
	head = lines[:200]
	abs_idx = _find_abstract_header_index(head)
	window = head if abs_idx is None else head[:abs_idx]
	# 進一步限制在文件最前面的一小段（多數題目位於開頭區域）
	window = window[:120]

	if not window:
		return None

	# 過濾非題目行
	deny = re.compile(r"(大學|學校|學院|系|所|研究所|學位|指導|導師|教授|學號|目錄|目次|致謝|誌謝|謝辭|關鍵字|关键词)")
	# 排除明顯的敘述性或非題目特徵
	bad_leading = re.compile(r"^(本研究|本文|本論文|因此|然而|在本研究中|在本論文中)")
	bad_punct = re.compile(r"[，。、；：！？()（）\[\]【】·]")
	candidates: List[str] = []
	for ln in window:
		s = ln.strip()
		if not s:
			continue
		# 必須含有 CJK
		if not _has_cjk(s):
			continue
		# 過濾明顯非題目標籤
		if deny.search(s):
			continue
		# 排除以敘述性開頭
		if bad_leading.search(s):
			continue
		# 排除包含明顯句內標點、或末尾為句號/冒號的行（較可能是段落句）
		if bad_punct.search(s) or s.endswith("。") or s.endswith(":") or s.endswith("："):
			continue
		# 過於複雜（含數字過多）則排除
		if sum(ch.isdigit() for ch in s) >= 2:
			continue
		# 長度範圍（避免過短/過長）
		if not (8 <= len(s) <= 40):
			continue
		candidates.append(s)

	if not candidates:
		return None
	# 選擇最長者做為題目（中文題目常為單行）
	return max(candidates, key=len)


def _expand_title_by_superstring(filename_title: str, lines: List[str]) -> Optional[str]:
	"""在『摘要/Abstract』之前，尋找包含檔名題目且更長的中文行，作為高可信度擴充題目。

	規則：
	- 僅在檔名題目含 CJK 時啟用（避免純英文檔名誤用）
	- 視窗：文件開頭至 Abstract 標題前，最多 ~200 行
	- 候選需包含檔名題目的「去除空白與標點後」的連續子字串
	- 允許少量標點，但末尾不可是句號/冒號；排除常見頁面標籤（學校/系所/致謝/關鍵字…）
	- 需嚴格「更長於」檔名題目（至少多 2 字），長度上限約 80 字
	"""
	if not filename_title:
		return None
	if not _has_cjk(filename_title):
		return None

	head = lines[:200]
	abs_idx = _find_abstract_header_index(head)
	window = head if abs_idx is None else head[:abs_idx]

	def _canon(s: str) -> str:
		# 去除空白與常見標點，便於子字串比對
		t = re.sub(r"[\s\u3000\t\r\n]", "", s)
		t = re.sub(r"[，。、；：！？,.;:!()（）\[\]【】·、\-—_/|]+", "", t)
		return t

	base = filename_title.strip()
	base_c = _canon(base)
	if len(base_c) < 4:
		return None

	deny = re.compile(r"(大學|學校|學院|系|所|研究所|學位|指導|導師|教授|學號|目錄|目次|致謝|誌謝|謝辭|關鍵字|关键词)")
	punct_pat = re.compile(r"[，。、；：！？()（）\[\]【】·,.;:!]")

	candidates: List[str] = []
	for ln in window:
		s = ln.strip()
		if not s:
			continue
		if not _has_cjk(s):
			continue
		if deny.search(s):
			continue
		# 末尾不可為句號/冒號，避免描述句
		if s.endswith("。") or s.endswith(":") or s.endswith("："):
			continue
		# 標點數量過多則排除（但允許少量標點）
		if len(punct_pat.findall(s)) > 3:
			continue
		s_c = _canon(s)
		if len(s_c) <= len(base_c) + 1:
			continue
		if base_c not in s_c:
			continue
		# 限制合理長度
		if len(s) > 80:
			continue
		candidates.append(s)

	if candidates:
		return max(candidates, key=len)

	# 次要路徑：若在 Abstract 前未找到，嘗試在文件開頭區域（前 ~120 行）搜尋一次
	# 排除明顯摘要/關鍵字標題行，避免誤取段落句
	window2 = head[:120]
	start_hdr = re.compile(r"^\s*(?:摘要|Abstract|ABSTRACT)\b")
	for ln in window2:
		s = ln.strip()
		if not s or not _has_cjk(s):
			continue
		if start_hdr.search(s):
			continue
		if deny.search(s):
			continue
		if s.endswith("。") or s.endswith(":") or s.endswith("："):
			continue
		if len(punct_pat.findall(s)) > 3:
			continue
		s_c = _canon(s)
		if len(s_c) <= len(base_c) + 1:
			continue
		if base_c not in s_c:
			continue
		if len(s) > 80:
			continue
		candidates.append(s)

	if not candidates:
		return None
	return max(candidates, key=len)


def _expand_title_by_right_context(filename_title: str, lines: List[str]) -> Optional[str]:
	"""當檔名題目略短時，嘗試在內文中以「右側語境」延伸題目。

	策略：
	- 僅對含 CJK 的檔名題目啟用，避免英文誤用。
	- 建立『去空白/去常見標點』的字元序列，用以定位檔名題目在原始文字中的位置。
	- 自匹配結尾往右在原文掃描，直到遇到明顯句讀或換行（。；！？：: 或換行）為止，
	  同時限制延伸長度（最多額外 30 個字元）。
	- 僅在延伸至少 2 個 CJK 字且結果看起來像標題（末尾不為句號/冒號、長度不超過 80）時回傳。
	"""
	if not filename_title or not _has_cjk(filename_title):
		return None

	# 僅取文件前段（2~3頁）作為搜尋範圍
	head = lines[:300]
	abs_idx = _find_abstract_header_index(head)
	window = head if abs_idx is None else head[:abs_idx]
	original = "\n".join(window)

	# 建立去除空白與弱標點的序列，同時回傳到原文索引的對應表
	def canon_with_map(s: str) -> Tuple[str, List[int]]:
		out_chars: List[str] = []
		mapping: List[int] = []
		for i, ch in enumerate(s):
			if ch in {" ", "\t", "\r", "\n", "\u3000"}:
				continue
			# 去除弱連接符與常見非終止標點，保留內容性字元以利匹配
			if re.match(r"[，,、()（）\[\]【】·\-—_/|]", ch):
				continue
			out_chars.append(ch)
			mapping.append(i)
		return "".join(out_chars), mapping

	base_c, _ = canon_with_map(filename_title)
	if len(base_c) < 4:
		return None

	canon, mapping = canon_with_map(original)
	pos = canon.find(base_c)
	if pos < 0:
		return None

	# 以原文索引為主，向右延伸直到終止符或上限
	start_orig = mapping[pos]
	end_canon = pos + len(base_c)
	end_orig = mapping[end_canon - 1] if end_canon - 1 < len(mapping) else len(original) - 1

	# 定義終止條件：句讀或換行
	TERMINATORS = set("。；！？：:")
	MAX_EXTRA = 60  # 原文額外延伸的最大字元數（含空白/連接符）

	cur = end_orig + 1
	steps = 0
	while cur < len(original) and steps < MAX_EXTRA:
		ch = original[cur]
		if ch in TERMINATORS:
			break
		# 遇到換行時，若下一行像是標題延續（多為純 CJK 且無句讀），則視為可延伸
		if ch == "\n":
			# 查看下一行內容
			nl_end = original.find("\n", cur + 1)
			seg_end = nl_end if nl_end != -1 else len(original)
			next_line = original[cur + 1:seg_end].strip()
			if next_line and _has_cjk(next_line):
				# 排除明顯標籤與句尾符號
				if not re.search(r"(關鍵字|关键词|摘要|Abstract|ABSTRACT|目錄|目次)", next_line) and not next_line.endswith("。") and not next_line.endswith(":") and not next_line.endswith("："):
					cur += 1
					steps += 1
					continue
			# 否則停止延伸
			break
		cur += 1
		steps += 1

	candidate = original[start_orig:cur].replace("\n", "").strip()
	# 基本檢核：需比檔名題目更長，且末尾不可是句號/冒號
	if len(candidate) <= len(filename_title) + 1:
		return None
	if candidate.endswith("。") or candidate.endswith(":") or candidate.endswith("："):
		return None
	# 限制合理長度
	if len(candidate) > 120:
		candidate = candidate[:120].rstrip()

	# 至少新增 2 個 CJK 字
	extra_cjk = re.findall(f"[{_CJK}]", candidate[len(filename_title):])
	if len(extra_cjk) < 2:
		return None

	# 過濾常見頁面標籤
	deny = re.compile(r"(大學|學校|學院|系|所|研究所|學位|指導|導師|教授|學號|目錄|目次|致謝|誌謝|謝辭|關鍵字|关键词)")
	if deny.search(candidate):
		return None

	return candidate


AUTHOR_REGEXES = [
	# 明確標籤行：『作者：』『作者』『研究生：』『學生：』『姓名：』『論文作者：』『作者姓名：』『畢業生：』等
	re.compile(r"^\s*作者\s*[：:]\s*(?P<name>.+?)\s*$"),
	re.compile(r"^\s*作者\s*[：:]?\s*$"),  # 下一行可能是姓名
	re.compile(r"^\s*(?:研究生|研 究 生|學生|学生)\s*[：:]?\s*(?P<name>[\u4e00-\u9fffA-Za-z .・．。-]{2,})\s*$"),
	re.compile(r"^\s*姓\s*名\s*[：:]?\s*(?P<name>[\u4e00-\u9fffA-Za-z .・．。-]{2,})\s*$"),
	re.compile(r"^\s*(?:論文作者|作者姓名|畢業生)\s*[：:]?\s*(?P<name>[\u4e00-\u9fffA-Za-z .・．。-]{2,})\s*$"),
	# 英文標籤
	re.compile(r"^\s*Authors?\s*[：:]\s*(?P<name>.+?)\s*$", re.IGNORECASE),
	re.compile(r"^\s*Author\b\s*[：:]\s*(?P<name>.+?)\s*$", re.IGNORECASE),
	re.compile(r"^\s*Student\b\s*[：:]\s*(?P<name>.+?)\s*$", re.IGNORECASE),
]


def _extract_author_from_lines(lines: List[str], verbose: bool = False) -> Optional[str]:
	# 先看含有『作者』/『Author』關鍵詞的行（加大搜尋範圍以涵蓋封面/口試審查頁面）
	for i, line in enumerate(lines[:200]):
		for rgx in AUTHOR_REGEXES:
			m = rgx.search(line)
			if m:
				name = m.group("name").strip()
				# 過濾過長或看似非人名的內容
				if len(name) > 100:
					continue
				if verbose:
					print(f"[AUTHOR] Label match by /{rgx.pattern}/ at line {i+1}: '{name}'")
				return name
		# 形式：上一行為『作者』，下一行為姓名
		if re.match(r"^\s*作者\s*[：:]?\s*$", line):
			if i + 1 < len(lines):
				nxt = lines[i + 1].strip()
				if 1 < len(nxt) <= 30:
					if verbose:
						print(f"[AUTHOR] Next-line after '作者' header at line {i+2}: '{nxt}'")
					return nxt

	# 結構性猜測：靠近『指導/教授/導師/Advisor/Supervisor』等關鍵上下文
	name_like = re.compile(r"^[\s　]*([\u4e00-\u9fff]{1,2})\s?([\u4e00-\u9fff]{1,2})[\s　]*$")
	ctx = re.compile(r"(指導|導師|教授|Advisor|Supervisor)", re.IGNORECASE)
	stop = {"碩士", "博士", "論文", "學位", "學校", "大學", "學院", "系", "所", "致謝", "誌謝", "謝辭"}
	for i, line in enumerate(lines[:200]):
		if not ctx.search(line):
			continue
		# 在鄰近範圍尋找姓名樣式
		for j in range(max(0, i-5), min(len(lines), i+6)):
			if j == i:
				continue
			s = lines[j].strip()
			m = name_like.match(s)
			if m:
				cand = (m.group(1) or "") + (m.group(2) or "")
				if any(w in cand for w in stop):
					continue
				if verbose:
					print(f"[AUTHOR] Structural fallback near advisor context (lines {i-3}..{i+3}): '{cand}'")
				return cand

	return None


def _guess_author_from_lines(lines: List[str], title: Optional[str] = None, verbose: bool = False) -> Optional[str]:
	"""在未找到帶有標籤的作者行時，嘗試從前幾十行中猜測姓名（2~4 個中文字）。

	改良：
	- 排除常見非姓名詞。
	- 若候選為題目片段則捨棄，避免將標題拆行的一部分誤判為作者。
	- 優先挑選鄰近學校/系所/指導教授等上下文的候選。
	"""
	# 僅保留通用的論文/章節/角色用語，避免領域詞造成過擬合；不使用單字功能詞以免誤殺姓名
	stopwords = {
		# 章節與結構
		"摘要", "中文摘要", "英文摘要", "Abstract", "ABSTRACT",
		"關鍵字", "关键词", "關鍵詞", "Keywords", "Index Terms",
		"目錄", "目次", "附錄",
		"引言", "緒論", "前言", "導論",
		"參考文獻", "参考文献", "References",
		"方法", "結果", "討論", "結論",
		# 身分與標籤
		"作者", "作者姓名", "論文作者", "姓名",
		"研究生", "學生", "學生姓名",
		"指導", "導師", "指導教授", "教授",
		# 校系與學位
		"大學", "學校", "學院", "系", "所", "研究所", "學位", "學程", "學號",
		# 其他常見頁面/欄位
		"致謝", "誌謝", "謝辭", "封面", "審定書", "口試委員", "委員"
	}
	name_like = re.compile(r"^[\s　]*([\u4e00-\u9fff]{1,2})\s?([\u4e00-\u9fff]{1,2})[\s　]*$")
	ctx_hint = re.compile(r"(大學|學校|學院|系|所|研究所|學位|論文|指導|教授|學號)")
	title_str = title or ""

	near_ctx: List[str] = []
	others: List[str] = []
	window = min(150, len(lines))
	for i in range(window):
		s = lines[i].strip()
		if not s:
			continue
		m = name_like.match(s)
		if not m:
			continue
		cand = (m.group(1) or "") + (m.group(2) or "")
		if any(w in cand for w in stopwords):
			continue
		if title_str and cand in title_str:
			continue
		neighborhood = "\n".join(lines[max(0, i-3): i+4])
		if ctx_hint.search(neighborhood):
			near_ctx.append(cand)
		else:
			others.append(cand)

	if near_ctx:
		if verbose:
			print(f"[AUTHOR] Guessed near academic context: '{near_ctx[0]}'")
		return near_ctx[0]
	if others:
		if verbose:
			print(f"[AUTHOR] Guessed by name-shape: '{others[0]}'")
		return others[0]
	return None


def extract_title_year_author(pdf_path: Path, verbose: bool = False, very_verbose: bool = False) -> Tuple[Optional[str], Optional[str], Optional[str]]:
	"""擷取題目、年分、作者（盡力而為的啟發式）。"""
	meta = _get_pdf_metadata(pdf_path)
	meta_title = _sanitize_meta_value(meta.get("title")) if meta else None
	meta_author = _sanitize_meta_value(meta.get("author")) if meta else None
	meta_creation = _sanitize_meta_value(meta.get("creationDate")) if meta else None

	# 題目：優先使用較可靠來源
	# 1) 若 metadata title 合理（非空、非 placeholder 且字數>=10 且含空白），優先使用
	meta_title_ok = (
		meta_title is not None and len(meta_title) >= 10 and (' ' in meta_title or '-' in meta_title)
	)

	filename_title = _parse_title_from_filename(pdf_path.stem)
	# 可疑的檔名題目（例如期刊代碼/DOI 片段），偏好以內文推斷英文題目
	fn_suspicious = (
		not _has_cjk(filename_title)
		and (re.fullmatch(r"[A-Za-z0-9_.-]{1,30}", pdf_path.stem) is not None)
	)

	title: Optional[str]
	if meta_title_ok:
		title = meta_title
	else:
		# 嘗試從前幾頁在 Abstract 之前擷取英文/中文題目，以及以檔名題目為錨點的擴充題目
		lines = _first_n_page_lines(pdf_path, n_pages=5)
		eng_title = _extract_english_title_from_lines(lines)
		cn_title = _extract_chinese_title_from_lines(lines)
		super_cn = _expand_title_by_superstring(filename_title, lines) if filename_title else None
		# 若未找到更長的上位字串，嘗試以右側語境延伸
		right_cn = _expand_title_by_right_context(filename_title, lines) if (filename_title and not super_cn) else None
		if eng_title and (fn_suspicious or not filename_title):
			title = eng_title
		else:
			# 若以檔名題目為錨點找到更完整的中文題目，優先採用；否則偏好檔名題目以避免誤取敘述句
			if super_cn or right_cn:
				title = super_cn or right_cn
			else:
				# 次順位：檔名題目，其次中文題目（僅在檔名題目不存在時使用）
				title = filename_title or cn_title or eng_title or meta_title

	if title:
		title = _normalize_title(title)

	# 年分：優先檔名，其次前幾頁文字，再來 metadata creationDate（避免以建立時間覆蓋實際發表年）
	year = None
	year = _parse_year_from_filename(pdf_path.stem)
	if not year:
		lines = _first_n_page_lines(pdf_path, n_pages=3)
		for ln in lines:
			y = _parse_year_from_string(ln)
			if y:
				year = y
				break
	if not year and meta_creation:
		# PyMuPDF 常見 creationDate 格式：'D:YYYYMMDDHHmmss...'
		m = re.search(r"(?:D:)?((?:19|20)\d{2})", meta_creation)
		if m:
			year = m.group(1)

	# 作者：優先從前幾頁文字擷取，其次 metadata（避免套版作者）
	# 第一次以 3 頁範圍尋找，以避免誤抓太遠內容
	lines = _first_n_page_lines(pdf_path, n_pages=3)
	if very_verbose:
		print("[AUTHOR] Scanning first 3 pages for labeled author...")
	author = _extract_author_from_lines(lines, verbose=very_verbose)
	method: str = ""
	if not author:
		author = _guess_author_from_lines(lines, title=title, verbose=very_verbose)
		if author:
			method = "guess-3p"
	else:
		method = "label-3p"

	# 若尚未取得作者，擴大搜尋範圍到前 8 頁，以涵蓋常見的封面/簽名頁面
	if not author:
		if very_verbose:
			print("[AUTHOR] Not found. Expanding search up to first 8 pages...")
		lines_wide = _first_n_page_lines(pdf_path, n_pages=8, max_lines=800)
		author = _extract_author_from_lines(lines_wide, verbose=very_verbose)
		if author:
			method = "label-8p"
		if not author:
			author = _guess_author_from_lines(lines_wide, title=title, verbose=very_verbose)
			if author:
				method = "guess-8p"

	# 偏好中文姓名：若題目含 CJK 且目前作者為純 ASCII，嘗試再找 CJK 姓名覆寫
	if author and not _has_cjk(author) and title and _has_cjk(title):
		if very_verbose:
			print(f"[AUTHOR] Current author '{author}' is ASCII-only; title is CJK. Attempting CJK override...")
		lines_wide = _first_n_page_lines(pdf_path, n_pages=8, max_lines=800)
		cjk_guess = _guess_author_from_lines(lines_wide, title=title, verbose=very_verbose)
		if cjk_guess and _has_cjk(cjk_guess):
			if very_verbose:
				print(f"[AUTHOR] Overriding ASCII author with CJK candidate: '{cjk_guess}'")
			author = cjk_guess
			method = (method + "+cjk") if method else "cjk-override"

	# 最後才回退到 metadata 的作者資訊
	if not author:
		if very_verbose and meta_author:
			print(f"[AUTHOR] Falling back to PDF metadata author: '{meta_author}'")
		author = meta_author
		method = "metadata"

	# 後處理：正規化中文姓名的空白與順序（若為「名 姓」且姓氏一字，調整為「姓名字」）
	def _normalize_cjk_name(name: str) -> str:
		if not name:
			return name
		s = name.strip()
		# 單一空白分隔且全為 CJK
		m = re.match(fr"^([{_CJK}]{{1,3}})\s+([{_CJK}]{{1,2}})$", s)
		if m:
			first, second = m.group(1), m.group(2)
			# 若第二段為一字，極可能為姓氏，調整為 姓 + 名
			if len(second) == 1:
				return second + first
		# 將內部多個空白收斂為單一空白
		s = re.sub(r"\s{2,}", " ", s)
		return s

	if author:
		# 避免使用無意義的 metadata 佔位詞（如 USER/ADMIN/UNKNOWN）
		if author.strip().lower() in {"user", "admin", "unknown", "untitled", "test"}:
			if very_verbose:
				print(f"[AUTHOR] Suppressing placeholder metadata author: '{author}'")
			author = None
		else:
			# 優先輸出中文姓名：若同時含中英文，僅保留中文部分
			def _to_cjk_author(name: str) -> str:
				parts = re.findall(f"[{_CJK}]+", name)
				if parts:
					# 取最長的連續中文片段視為姓名
					return max(parts, key=len)
				return name.strip()
			# 先正規化中英間空白，再轉為僅中文
			author = _normalize_cjk_name(author)
			author = _to_cjk_author(author)
	# 單行摘要輸出（-v）：僅一行顯示作者與來源；-vv 顯示完整追蹤
	if verbose:
		print(f"[AUTHOR] {author or ''} ({method or 'n/a'})")
	if very_verbose:
		print(f"[AUTHOR] Final author: '{author or ''}'")

	return title, year, author


def compile_patterns(patterns: Sequence[str]) -> List[re.Pattern[str]]:
	return [re.compile(p, flags=re.IGNORECASE) for p in patterns]


START_REGEXES = compile_patterns(START_PATTERNS)
END_REGEXES = compile_patterns(END_PATTERNS)


def _looks_like_header(line: str) -> bool:
	"""判斷是否看起來像是章節標題（軟性結束條件）。

	- 短行且大寫比例高的英文文字
	- 常見標題字眼（中/英），或前綴是常見章節名稱
	"""
	s = line.strip()
	if not s:
		return False
	# 全英文且幾乎全為大寫（例如章節標題）
	if (
		len(s) <= 60
		and all(ord(ch) < 128 for ch in s)  # 僅限 ASCII，避免中英混合被誤判
		and s.isupper()
		and re.search(r"[A-Z]", s)
	):
		return True
	# Typical patterns like "1. Introduction" already covered; here we add small capsish words
	if re.match(r"^(?:Abstract|ABSTRACT|Introduction|Conclusions?)\b", s):
		return True
	if re.match(r"^(?:摘要|引言|結論|結語|緒論|關鍵字|目錄)\b", s):
		return True
	return False


def _join_with_hyphen_fix(lines: List[str]) -> str:
	"""保留原始換行；僅修正英文在換行處的連字號斷行。

	規則：若上一行以 '-' 結尾，且下一行以小寫英文字母開頭，移除連字號並直接銜接；
	其餘情況不進行合併，維持每行一換行，貼近原始排版與提供的答案格式。
	"""
	out: List[str] = []
	for line in lines:
		if out:
			prev = out[-1]
			if prev.endswith('-') and line and line[:1].islower():
				out[-1] = prev[:-1] + line.lstrip()
				continue
		out.append(line)
	return "\n".join(out)


_CJK = r"\u2E80-\u2FFF\u3000-\u303F\u31C0-\u31EF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF"  # 中日韓多區段


def _normalize_cjk_ascii_spacing(text: str, mode: str) -> str:
	"""控制中英/數字混排時的空白處理。

	mode:
	  - 'none': 不做處理（預設）
	  - 'insert': 在 CJK 與 ASCII/數字 之間插入單一空白
	  - 'remove': 移除 CJK 與 ASCII/數字 之間的空白
	"""
	if mode == "none":
		return text

	if mode == "insert":
		text = re.sub(fr"([{_CJK}])([A-Za-z0-9])", r"\1 \2", text)
		text = re.sub(fr"([A-Za-z0-9])([{_CJK}])", r"\1 \2", text)
		# 收斂多重空白為單一空白（不跨行）
		text = re.sub(r"[\t ]{2,}", " ", text)
		return text

	if mode == "remove":
		# CJK + 空白 + ASCII/數字 -> 直接相連
		text = re.sub(fr"([{_CJK}])[\t ]+([A-Za-z0-9])", r"\1\2", text)
		# ASCII/數字 + 空白 + CJK -> 直接相連
		text = re.sub(fr"([A-Za-z0-9])[\t ]+([{_CJK}])", r"\1\2", text)
		return text

	return text


def extract_abstract_from_pdf(pdf_path: Path, verbose: bool = False) -> ExtractionResult:
	"""依啟發式規則，自 PDF 擷取摘要區段。

	回傳 ExtractionResult：若成功，包含摘要文字與起訖頁、使用到的起訖模式；若失敗，於 reason 說明原因。
	"""
	lines: List[Tuple[int, str]] = list(iter_pdf_lines(pdf_path))
	if not lines:
		return ExtractionResult(
			abstract=None,
			start_page=None,
			end_page=None,
			start_marker=None,
			end_marker=None,
			reason="No text extracted",
		)

	# 1) 尋找摘要起始標記
	start_idx: Optional[int] = None
	start_page: Optional[int] = None
	start_marker: Optional[str] = None
	inline_first: Optional[str] = None

	for idx, (pg, line) in enumerate(lines):
		for rgx in START_REGEXES:
			m = rgx.search(line)
			if m:
				start_idx = idx
				start_page = pg
				start_marker = m.re.pattern
				inline_first = m.groupdict().get("inline") if m.groupdict() else None
				if verbose:
					print(f"Start found on page {pg+1}: '{line.strip()[:80]}'")
				break
		if start_idx is not None:
			break

	if start_idx is None:
		return ExtractionResult(
			abstract=None,
			start_page=None,
			end_page=None,
			start_marker=None,
			end_marker=None,
			reason="Start marker not found",
		)

	# 2) 自起點往後累積內容，直到遇到結束標記或頁數上限
	collected: List[str] = []
	if inline_first:
		collected.append(inline_first.strip())

	end_idx: Optional[int] = None
	end_page: Optional[int] = None
	end_marker: Optional[str] = None

	start_pg = start_page if start_page is not None else 0
	max_end_pg = start_pg + MAX_ABSTRACT_PAGES - 1

	# 若起始行僅為標題（無內嵌內容），就從下一行開始蒐集；若有內嵌內容，已先加入 collected
	cursor = start_idx + 1 if not inline_first or re.match(r"^\s*(?:Abstract|ABSTRACT|(?:中文)?摘要)\s*[：:]?\s*$", lines[start_idx][1], re.IGNORECASE) else start_idx + 1

	for idx in range(cursor, len(lines)):
		pg, line = lines[idx]

	# 若超過頁數上限仍未遇到結束標記，則以頁數上限作為終止
		if pg > max_end_pg:
			end_idx = idx
			end_page = lines[idx - 1][0] if idx > cursor else start_pg
			end_marker = f"page_limit_{MAX_ABSTRACT_PAGES}"
			if verbose:
				print(f"Stopping at page limit after page {end_page+1}")
			break

	# 硬性結束標記：關鍵字、參考文獻、第一章標題等
		for rgx in END_REGEXES:
			if rgx.search(line):
				end_idx = idx
				end_page = pg
				end_marker = rgx.pattern
				if verbose:
					print(f"End found on page {pg+1}: '{line.strip()[:80]}'")
				break
		if end_idx is not None:
			break

		# 軟性結束：需已收集到一定行數（避免太早截斷），且該行疑似標題
		if len(collected) >= 5 and _looks_like_header(line):
			end_idx = idx
			end_page = pg
			end_marker = "soft_header"
			if verbose:
				print(f"Soft end at page {pg+1}: '{line.strip()[:80]}'")
			break

	# 其他情況：持續加入行內容
		collected.append(line)

	# 若始終未命中結束條件：仍舊收集到頁數上限
	if end_idx is None:
		# Gather until max_end_pg
		for idx in range(cursor, len(lines)):
			pg, line = lines[idx]
			if pg > max_end_pg:
				break
			collected.append(line)
		end_page = min(max_end_pg, lines[-1][0])
		end_marker = "eof_or_limit"

	# 修剪開頭與結尾的空白行
	while collected and not collected[0].strip():
		collected.pop(0)
	while collected and not collected[-1].strip():
		collected.pop()

	# 移除結尾可能夾帶的頁碼/羅馬數字頁碼（常見於前置頁 ii、iii 等）
	roman = re.compile(r"^[ivxlcdmIVXLCDM]{1,4}\.?$")
	digits = re.compile(r"^\d{1,3}$")
	while collected:
		tail = collected[-1].strip()
		if roman.match(tail) or digits.match(tail):
			collected.pop()
			continue
		break

	text = _join_with_hyphen_fix(collected).strip()  # 修正英文連字號換行，保留換行
	text = re.sub(r"\n{3,}", "\n\n", text)  # collapse excessive blank lines

	if not text:
		return ExtractionResult(
			abstract=None,
			start_page=start_page,
			end_page=end_page,
			start_marker=start_marker,
			end_marker=end_marker,
			reason="Empty text after extraction",
		)

	return ExtractionResult(
		abstract=text,
		start_page=start_page,
		end_page=end_page,
		start_marker=start_marker,
		end_marker=end_marker,
		reason=None,
	)


def process_path(input_path: Path, output_dir: Path, recursive: bool, verbose: bool, very_verbose: bool, to_csv:bool) -> int:
	"""處理單一檔案或整個資料夾。回傳成功擷取的 PDF 數量。"""
	pdf_files: List[Path]
	if input_path.is_file() and input_path.suffix.lower() == ".pdf":
		pdf_files = [input_path]
	elif input_path.is_dir():
		if recursive:
			pdf_files = sorted(p for p in input_path.rglob("*.pdf"))
		else:
			pdf_files = sorted(p for p in input_path.glob("*.pdf"))
	else:
		print(f"[WARN] Skipping non-PDF path: {input_path}")
		return 0

	output_dir.mkdir(parents=True, exist_ok=True)  # 確保輸出資料夾存在
	txt_dir = output_dir / "txt"  # 將純文字輸出集中於子資料夾
	txt_dir.mkdir(parents=True, exist_ok=True)

	success = 0
	records: List[Tuple[str, str, str, str]] = []  # (title, year, author, abstract)
	for pdf in pdf_files:
		rel_name = pdf.stem
		out_file = txt_dir / f"{rel_name}.txt"
		try:
			# 1) 題目/年分/作者
			title, year, author = extract_title_year_author(pdf, verbose=verbose, very_verbose=very_verbose)

			# 2) 摘要
			# 僅在 -vv 時才顯示摘要起訖偵測的詳細訊息
			result = extract_abstract_from_pdf(pdf, verbose=very_verbose)
			if result.abstract:
				# 需求格式：摘要作為單段落展示
				abstract_text = result.abstract.replace("\r", "").replace("\n", "")

				# 3) 依指定格式輸出
				content = (
					f"題目：{title or ''}\n\n"
					f"年分：{year or ''}\n\n"
					f"作者：{author or ''}\n\n"
					f"摘要：{abstract_text}"
				)
				out_file.write_text(content, encoding="utf-8-sig")  # Windows 友善的 UTF-8 BOM
				success += 1
				records.append((title or "", year or "", author or "", abstract_text))
				if verbose:
					print(
						f"[OK] {pdf.name}: abstract pages {(result.start_page or 0)+1}-"
						f"{(result.end_page or result.start_page or 0)+1}, saved -> {out_file.name}"
					)
			else:
				reason = result.reason or "Unknown"
				if verbose:
					print(f"[MISS] {pdf.name}: {reason}")
		except Exception as e:
			print(f"[ERROR] {pdf}: {e}")
	# 寫出整併 CSV：論文整理.csv
	if to_csv:
		if records:
			csv_path = output_dir / "論文整理.csv"
			buf = io.StringIO()
			# 使用單一換行符，避免 Windows 檢視器顯示空白列
			writer = csv.writer(buf, lineterminator='\n')
			writer.writerow(["題目", "年分", "作者", "摘要"])
			for (title, year, author, abstract_text) in records:
				writer.writerow([title, year, author, abstract_text])
			csv_content = buf.getvalue()
			buf.close()
			try:
				csv_path.write_text(csv_content, encoding="utf-8-sig")
				if verbose:
					print(f"Saved CSV -> {csv_path} ({len(records)} rows)")
			except PermissionError:
				# 不產生其他檔名，維持目標名稱固定；提示使用者關閉檔案後重跑。
				print(f"[WARN] Unable to write {csv_path} because it is open or locked. Please close it and rerun.")

	return success, title, year, author, abstract_text


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
	# 參數解析：支援檔案/資料夾、輸出路徑、是否遞迴與詳細模式
	p = argparse.ArgumentParser(description="Extract Abstract/摘要 from PDFs")
	p.add_argument("input", type=str, help="PDF file or directory containing PDFs")
	p.add_argument(
		"--output-dir",
		type=str,
		default="abstract_output",
		help="Directory to write extracted abstracts (default: abstract_output)",
	)
	p.add_argument(
		"--recursive",
		action="store_true",
		help="Recursively scan subdirectories when input is a directory",
	)
	p.add_argument("--verbose", "-v", action="store_true", help="Verbose logging (concise; prints one-line author)")
	p.add_argument("--very-verbose", "-vv", action="store_true", help="Very verbose logging (detailed trace)")
	# 調試模式：內嵌原 inspect 工具
	p.add_argument("--inspect", action="store_true", help="Print first pages/lines of a PDF and exit (debug)")
	p.add_argument("--inspect-pages", type=int, default=2, help="Pages to scan in inspect mode (default: 2)")
	p.add_argument("--inspect-lines", type=int, default=200, help="Max lines to print in inspect mode (default: 200)")
	# 移除 spacing 與 answer 對齊參數（按需求改為固定不處理）
	return p.parse_args(argv)


def _inspect_pdf(pdf_path: Path, pages: int, max_lines: int) -> int:
	"""輸出 PDF 前幾頁的前若干行，供除錯觀察。"""
	if fitz is None:
		print("PyMuPDF (pymupdf) is required. Install with: pip install pymupdf")
		return 3
	if not pdf_path.is_file():
		print(f"Inspect requires a PDF file, got: {pdf_path}")
		return 2
	doc = fitz.open(pdf_path)
	try:
		count = 0
		for i in range(min(len(doc), max(1, pages))):
			page = doc[i]
			text = page.get_text("text").replace("\r\n", "\n").replace("\r", "\n")
			for ln in text.split("\n"):
				print(f"{i+1:02d}: {ln}")
				count += 1
				if count >= max_lines:
					return 0
		return 0
	finally:
		doc.close()


def main(argv: Optional[Sequence[str]] = None) -> int:
	# 進入點：檢查路徑與套件，執行擷取流程
	args = parse_args(argv)
	input_path = Path(args.input)
	output_dir = Path(args.output_dir)
	# -vv implies -v 行為
	if args.very_verbose:
		args.verbose = True

	if not input_path.exists():
		print(f"Input path not found: {input_path}")
		return 2
	if fitz is None:
		print("PyMuPDF (pymupdf) is required. Install with: pip install pymupdf")
		return 3

	# Inspect 模式：僅輸出前若干頁/行並結束
	if args.inspect:
		return _inspect_pdf(input_path, pages=args.inspect_pages, max_lines=args.inspect_lines)

	# 使用統一的處理流程（輸出 題目/年分/作者/摘要 到 .txt）
	count, title, year, author, abstract_text = process_path(input_path, output_dir, recursive=args.recursive, verbose=args.verbose, very_verbose=args.very_verbose, to_csv=True)
	if args.verbose:
		print(f"\nCompleted. Processed files: {count}")
	return 0 if count > 0 else 1


if __name__ == "__main__":
	sys.exit(main())