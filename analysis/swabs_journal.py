from __future__ import annotations

import re
import unicodedata
import datetime as dt
import pandas as pd
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple


DATE_COL = "Дата исследования"
DEP_COL = "Подразделение"
BUILDING_COL = "Корпус"

SHEETS_DEFAULT = ("Абиотические", "Воздух", "Персонал")


def _clean_text(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)) or pd.isna(x):
        return ""
    s = unicodedata.normalize("NFKC", str(x))

    s = s.replace("\xa0", " ")
    s = s.replace("\u200b", "")
    s = s.replace("\ufeff", "")
    s = s.replace("\u00ad", "")

    s = s.replace("\r", " ").replace("\n", " ").replace("\t", " ")
    s = s.replace("\v", " ").replace("\f", " ")

    s = re.sub(r"\s+", " ", s).strip()
    return s


def _to_date(x) -> Optional[pd.Timestamp]:
    if x is None or pd.isna(x):
        return None

    if isinstance(x, (pd.Timestamp, dt.datetime, dt.date)):
        try:
            return pd.Timestamp(x.date())
        except Exception:
            return pd.Timestamp(pd.to_datetime(x).date())

    s = _clean_text(x)
    if not s:
        return None

    s = s.replace("г.", "").replace("г", "").strip()

    if re.match(r"^\d{4}-\d{2}-\d{2}(\b|T)", s):
        dtv = pd.to_datetime(s, errors="coerce", yearfirst=True, dayfirst=False)
    else:
        dtv = pd.to_datetime(s, errors="coerce", dayfirst=True)

    if pd.isna(dtv):
        return None

    return pd.Timestamp(dtv.date())


def _canon_operblock(dep_clean: str, building_clean: str) -> Optional[str]:
    dep_l = (dep_clean or "").lower()
    if "оперблок" not in dep_l:
        return None

    b_l = (building_clean or "").lower()

    m = re.search(r"(?<!\d)(2|3|4|5|6|8)(?!\d)", dep_l)
    if not m:
        return None
    floor = m.group(1)

    if floor == "6" and "фпц" in b_l:
        return "ОПЕРБЛОК 6 ЭТАЖ ФПЦ"

    if floor == "4" and "кдц" in b_l:
        return "ОПЕРБЛОК 4 ЭТАЖ КДЦ"

    if floor == "3" and "надстро" in b_l:
        return "ОПЕРБЛОК 3 ЭТАЖ НАДСТРОЙ"

    if "главн" in b_l and floor in ("2", "5", "6", "8"):
        return f"ОПЕРБЛОК {floor} ЭТАЖ"

    return None


def _norm_col(c: object) -> str:
    return _clean_text(c).lower()


def _find_col(df: pd.DataFrame, want: str) -> Optional[str]:
    want_n = _norm_col(want)
    if not want_n:
        return None

    for c in df.columns:
        if _norm_col(c) == want_n:
            return c

    for c in df.columns:
        if want_n in _norm_col(c):
            return c

    return None


def _as_str(v) -> str:
    return v if isinstance(v, str) else ""


@dataclass
class SwabsJournal:
    path: str
    sheets: Tuple[str, ...] = SHEETS_DEFAULT
    data_by_sheet: Dict[str, pd.DataFrame] = None

    def load(self) -> None:
        data: Dict[str, pd.DataFrame] = {}

        for sh in self.sheets:
            df = pd.read_excel(self.path, sheet_name=sh)
            df = df.dropna(how="all").copy()

            df.columns = [_clean_text(c) for c in df.columns]

            date_col = _find_col(df, DATE_COL)
            dep_col = _find_col(df, DEP_COL)
            building_col = _find_col(df, BUILDING_COL)

            if not date_col:
                raise ValueError(f"На листе '{sh}' нет колонки '{DATE_COL}' (или похожей)")
            if not dep_col:
                raise ValueError(f"На листе '{sh}' нет колонки '{DEP_COL}' (или похожей)")

            if not building_col:
                df[BUILDING_COL] = ""
                building_col = BUILDING_COL

            df["_sheet"] = sh
            df["_date"] = df[date_col].apply(_to_date)

            # сохраняем ОРИГИНАЛ (для совместимости старых алиасов)
            df["_dep_raw_orig"] = df[dep_col].astype(str)

            df["_dep_raw_clean"] = df[dep_col].apply(_clean_text)
            df["_building_clean"] = df[building_col].apply(_clean_text)

            df = df[df["_date"].notna() & df["_dep_raw_clean"].astype(bool)].copy()
            data[sh] = df

        self.data_by_sheet = data

    def unique_raw_departments(self) -> List[str]:
        s = set()
        for df in self.data_by_sheet.values():
            for t in df["_dep_raw_clean"].dropna().astype(str):
                t = _clean_text(t)
                if t:
                    s.add(t)
        return sorted(s)

    def apply_department_mapping(self, mapping: Dict[str, str]) -> None:
        for sh, df in self.data_by_sheet.items():

            def map_row(row) -> str:
                dep_clean = row.get("_dep_raw_clean", "") or ""
                b_clean = row.get("_building_clean", "") or ""

                oper = _canon_operblock(dep_clean, b_clean)
                if oper:
                    return oper

                v = _as_str(mapping.get(dep_clean))
                if v:
                    return v

                raw_orig = _clean_text(row.get("_dep_raw_orig", ""))
                v2 = _as_str(mapping.get(raw_orig))
                if v2:
                    return v2

                return ""

            df["_dep_canon"] = df.apply(map_row, axis=1)
            self.data_by_sheet[sh] = df

    def unknown_departments_for_mapping(self, mapping: Dict[str, str]) -> List[str]:
        unknown = set()

        for df in self.data_by_sheet.values():
            if df is None or df.empty:
                continue

            for _, row in df.iterrows():
                dep_clean = row.get("_dep_raw_clean", "") or ""
                b_clean = row.get("_building_clean", "") or ""

                if not dep_clean:
                    continue
                if _canon_operblock(dep_clean, b_clean):
                    continue
                if _as_str(mapping.get(dep_clean)):
                    continue

                unknown.add(dep_clean)

        return sorted(unknown)

    def list_dates_for_department(self, dep: str) -> List[pd.Timestamp]:
        dep = _clean_text(dep)
        if not dep:
            return []

        dates = set()
        for df in self.data_by_sheet.values():
            part = df[df["_dep_canon"] == dep] if "_dep_canon" in df.columns else df[df["_dep_raw_clean"] == dep]
            dates.update(part["_date"].unique().tolist())

        return sorted(dates, reverse=True)

    def filter_day(self, dep: str, day: pd.Timestamp, sheet: str) -> pd.DataFrame:
        dep = _clean_text(dep)
        df = self.data_by_sheet.get(sheet)
        if df is None or not dep:
            return pd.DataFrame()

        day = pd.Timestamp(day).normalize()
        out = df[(df["_dep_canon"] == dep) & (df["_date"] == day)].copy() if "_dep_canon" in df.columns \
            else df[(df["_dep_raw_clean"] == dep) & (df["_date"] == day)].copy()

        return out[[c for c in out.columns if not c.startswith("_")]]

    def list_all_dates(self) -> List[pd.Timestamp]:
        dates = set()
        for df in self.data_by_sheet.values():
            if df is None or df.empty:
                continue
            dates.update(df["_date"].dropna().unique().tolist())
        return sorted(dates, reverse=True)

    def filter_day_all(self, day: pd.Timestamp, sheet: str) -> pd.DataFrame:
        df = self.data_by_sheet.get(sheet)
        if df is None:
            return pd.DataFrame()

        day = pd.Timestamp(day).normalize()
        out = df[df["_date"] == day].copy()
        return out[[c for c in out.columns if not c.startswith("_")]]
