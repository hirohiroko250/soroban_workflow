
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# OZA 出欠ページ スクレイパー v2
#  - ログイン → 退勤ボタン → 生徒出欠簿 → ブランド変更 → 校舎＋日付で表示
#  - 取得結果を Excel（Raw / ActiveSlots / T_Slot）へ出力

import argparse
import os
import sys
import re
import time
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Dict, List, Tuple, Optional

import requests
from bs4 import BeautifulSoup
import pandas as pd

try:
    import tomllib  # Py3.11+
except Exception:
    tomllib = None


DEFAULT_LOGIN_URL = "https://www.o-za.jp/oza/AdminLogin.aspx"   # ← ルート直下
DEFAULT_BASE_ADMIN = "https://www.o-za.jp/oza/AdminArea/"       # ← AdminArea配下用
DEFAULT_ATTENDANCE_URL = DEFAULT_BASE_ADMIN + "toDayAttendanceSeach.aspx"
DEFAULT_CLOCK_URL      = DEFAULT_BASE_ADMIN + "ClockInOut.aspx"
DEFAULT_USER = "5uz.xxxaoi@gmail.com"
DEFAULT_PASS = "hN5keQGc"   # ⚠ 実運用では環境変数/安全な保管方法を推奨


@dataclass
class Config:
    base_url: str = DEFAULT_BASE_ADMIN
    login_url: str = DEFAULT_LOGIN_URL
    attendance_url: str = DEFAULT_ATTENDANCE_URL
    clock_url: str = DEFAULT_CLOCK_URL
    username: str = DEFAULT_USER
    password: str = DEFAULT_PASS
    course_ids: str = "2,145"  # アンそろばんクラブ2, 聖光学院そろばん145
    user_field: str = "txtLog_ID"
    pass_field: str = "txtLog_PW"
    login_button: str = "btnLoginRun"
    user_agent: str = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) OZA-PayrollScraper/2.0"


def yyyymm_to_range(yyyy_mm: str) -> Tuple[date, date]:
    if "-" in yyyy_mm:
        y, m = yyyy_mm.split("-", 1)
    else:
        y, m = yyyy_mm[:4], yyyy_mm[4:]
    y = int(y); m = int(m)
    start = date(y, m, 1)
    end = date(y + (m == 12), (m % 12) + 1, 1) - timedelta(days=1)
    return start, end


def resolve_month_arg(value: Optional[str]) -> str:
    if not value or value.lower() == 'auto':
        today = date.today()
        first = date(today.year, today.month, 1)
        prev = first - timedelta(days=1)
        return prev.strftime('%Y-%m')
    return value


def extract_hidden_fields(soup: BeautifulSoup) -> Dict[str, str]:
    data = {}
    for inp in soup.select("input[type=hidden]"):
        name = inp.get("name")
        if name:
            data[name] = inp.get("value", "")
    return data


def aspnet_post(session: requests.Session, url: str, soup: BeautifulSoup,
                event_target: str = "", event_argument: str = "", extra_form: dict = None):
    payload = extract_hidden_fields(soup)
    if event_target:
        payload["__EVENTTARGET"] = event_target
    if event_argument is not None:
        payload["__EVENTARGUMENT"] = event_argument
    if extra_form:
        payload.update(extra_form)
    resp = session.post(url, data=payload)
    resp.raise_for_status()
    return resp


def login(session: requests.Session, cfg: Config, verbose: bool = True) -> bool:
    r = session.get(cfg.login_url)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "lxml")
    extra = {
        cfg.user_field: cfg.username,
        cfg.pass_field: cfg.password,
    }
    resp = aspnet_post(session, cfg.login_url, soup,
                       event_target=cfg.login_button,
                       event_argument="",
                       extra_form=extra)
    if verbose:
        print(f"[login] status={resp.status_code} url={resp.url}")
    ok = ("btnLogout" in resp.text) or ("ログアウト" in resp.text) or ("/AdminArea/" in resp.url)
    if verbose:
        print(f"[login] success? {ok}")
    return ok


def click_work_end(session: requests.Session, cfg: Config, verbose: bool = True) -> bool:
    # ClockInOut.aspx で ctl00$btnWorkEnd を __doPostBack
    r = session.get(cfg.clock_url)
    if r.status_code != 200:
        if verbose: print(f"[work_end] GET failed: {r.status_code}")
        return False
    soup = BeautifulSoup(r.text, "lxml")
    target = "ctl00$btnWorkEnd"
    if not soup.find(id="ctl00_btnWorkEnd"):
        # 代替: 出欠ページでもヘッダにボタンがあれば押下
        r2 = session.get(cfg.attendance_url)
        if r2.status_code != 200:
            if verbose: print(f"[work_end] alt GET failed: {r2.status_code}")
            return False
        soup = BeautifulSoup(r2.text, "lxml")
        if not soup.find(id="ctl00_btnWorkEnd"):
            if verbose: print("[work_end] btn not found; skip")
            return False
        post_url = cfg.attendance_url
    else:
        post_url = cfg.clock_url

    resp = aspnet_post(session, post_url, soup, event_target=target, event_argument="")
    if verbose:
        print(f"[work_end] post status={resp.status_code}")
    return True


def open_attendance(session: requests.Session, cfg: Config) -> BeautifulSoup:
    r = session.get(cfg.attendance_url)
    r.raise_for_status()
    return BeautifulSoup(r.text, "lxml")


def change_course(session: requests.Session, cfg: Config, soup: BeautifulSoup, course_id: int) -> BeautifulSoup:
    # ブランドDDLの変更ポストバック模倣
    extra = {
        "ctl00$CPH$ddlSeachCourseID": str(course_id),
        "ctl00$CPH$txtSeachCGP_INDEX": "ALL",
    }
    resp = aspnet_post(session, cfg.attendance_url, soup,
                       event_target="ctl00$CPH$ddlSeachCourseID",
                       event_argument="",
                       extra_form=extra)
    return BeautifulSoup(resp.text, "lxml")


def parse_school_options_from_soup(soup: BeautifulSoup) -> List[Tuple[str, str]]:
    options = []
    ddl = soup.select_one("select#ctl00_CPH_ddlSeachSchoolID")
    if not ddl:
        return options
    for opt in ddl.select("option"):
        val = (opt.get("value") or "").strip()
        txt = (opt.get_text() or "").strip()
        if val:
            options.append((val, txt))
    return options


def parse_attendance_table(html: str) -> List[dict]:
    soup = BeautifulSoup(html, "lxml")
    table = soup.find("table", id="TblDataList")
    if not table:
        return []
    school_name = None
    ddl = soup.select_one("select#ctl00_CPH_ddlSeachSchoolID option[selected]")
    if ddl:
        school_name = ddl.get_text(strip=True)

    rows = []
    trs = table.find_all("tr")
    if len(trs) <= 2:
        # 「授業予定はありません」等
        return rows
    for tr in trs[2:]:
        tds = tr.find_all("td")
        if len(tds) != 6:
            continue
        limit_name = tds[0].get_text(strip=True)
        start_label = tds[1].get_text(strip=True)  # "1605～"
        title = tds[2].get_text(strip=True)
        try:
            expected = int(tds[3].get_text(strip=True) or 0)
        except:
            expected = 0
        try:
            trial = int(tds[4].get_text(strip=True) or 0)
        except:
            trial = 0
        hhmm = re.sub(r"[^0-9]", "", start_label)
        if len(hhmm) >= 3:
            hh = int(hhmm[:-2]); mm = int(hhmm[-2:])
            start_time = f"{hh:02d}:{mm:02d}"
        else:
            start_time = None
        has_class = (expected + trial) > 0
        rows.append({
            "limit": limit_name,
            "start_time": start_time,
            "class_name": title,
            "expected_count": expected,
            "trial_count": trial,
            "has_class": has_class,
            "school_name": school_name,
        })
    return rows


def parse_class_detail(html: str) -> dict:
    """
    クラス詳細ページ (toDayAttendanceDetail.aspx) から講師情報と生徒出欠情報を抽出

    Returns:
        dict: {
            "teacher_id": str,
            "teacher_name": str,
            "teacher_attendance": str,  # 講師出席状態（出席/欠席）
            "attendance_count": int,  # 出席人数（合計）
            "attendance_count_regular": int,  # 通常出席人数（振替なし）
            "attendance_count_substitution": int,  # 振替出席人数
            "absent_count": int,  # 欠席人数
            "class_name": str,
            "date": str,
            "start_time": str,
            "school_name": str,
            "students": List[dict]  # 各生徒の詳細情報
        }
    """
    soup = BeautifulSoup(html, "lxml")
    result = {
        "teacher_id": None,
        "teacher_name": None,
        "teacher_attendance": None,
        "teacher_memo": None,
        "attendance_count": 0,
        "attendance_count_regular": 0,
        "attendance_count_substitution": 0,
        "substitution_count": 0,  # 後方互換性のため残す
        "absent_count": 0,
        "class_name": None,
        "date": None,
        "start_time": None,
        "school_name": None,
        "students": []
    }

    # クラス基本情報を取得
    class_name_span = soup.find("span", id="ctl00_CPH_lblClassGroupName")
    if class_name_span:
        result["class_name"] = class_name_span.get_text(strip=True)

    date_span = soup.find("span", id="ctl00_CPH_lblPlanDate")
    if date_span:
        result["date"] = date_span.get_text(strip=True)

    time_span = soup.find("span", id="ctl00_CPH_lblStartRealTime")
    if time_span:
        result["start_time"] = time_span.get_text(strip=True)

    school_span = soup.find("span", id="ctl00_CPH_lblSchoolName")
    if school_span:
        result["school_name"] = school_span.get_text(strip=True)

    # 講師情報と生徒情報を取得（複数のTblDataListがある）
    all_tables = soup.find_all("table", id="TblDataList")

    # 講師情報を取得（最初のTblDataList）
    if all_tables:
        teacher_table = all_tables[0]
        teacher_rows = teacher_table.find_all("tr")
        # ヘッダー行をスキップして最初のデータ行から講師名を取得
        for tr in teacher_rows[1:]:
            # <th>と<td>が混在している場合があるので、両方取得
            tds = tr.find_all("td")
            ths = tr.find_all("th")

            # 講師テーブルの構造:
            # <th>講師 1</th> <td>ID:8211256</td> <td>竹内 真奈美</td> <td><checkbox></td> <td>出席</td> ...
            if len(tds) >= 2:
                # tds[0] = "ID:8211256"
                # tds[1] = "竹内 真奈美"
                if "ID:" in tds[0].get_text(strip=True):
                    # 講師IDを抽出（"ID:8211256" から "8211256" を取り出す）
                    teacher_id_text = tds[0].get_text(strip=True)
                    # 空白で分割して "ID:" を含む部分を探す
                    teacher_id = None
                    for part in teacher_id_text.split():
                        if part.startswith("ID:"):
                            teacher_id = part.replace("ID:", "").strip()
                            break
                    if not teacher_id:
                        # フォールバック：単純にreplaceする
                        teacher_id = teacher_id_text.replace("ID:", "").split()[0].strip()
                    result["teacher_id"] = teacher_id

                    # 講師名を取得
                    teacher_name = tds[1].get_text(strip=True)
                    if teacher_name and teacher_name != "":
                        result["teacher_name"] = teacher_name

                    # 講師の出席状態を取得（tds[3]が出席/欠席のテキスト）
                    if len(tds) >= 4:
                        result["teacher_attendance"] = tds[3].get_text(strip=True)

                    # 講師の備考を取得（tds[5]が備考欄の可能性）
                    if len(tds) >= 6:
                        result["teacher_memo"] = tds[5].get_text(strip=True)

                    # デバッグ出力
                    print(f"    [DEBUG] 講師情報: ID={result['teacher_id']}, 名前={result['teacher_name']}, 出席={result['teacher_attendance']}, 備考={result['teacher_memo']}")

                    break
                # もしくは別の構造の場合
                elif len(tds) >= 3:
                    teacher_name = tds[2].get_text(strip=True)
                    if teacher_name and teacher_name != "":
                        result["teacher_name"] = teacher_name
                        break

    # 生徒出欠情報を取得（2つ目のTblDataList）
    if len(all_tables) >= 2:
        student_table = all_tables[1]
        student_rows = student_table.find_all("tr")

        # ヘッダー行をスキップして生徒データを処理
        for tr in student_rows[1:]:  # 最初の行はヘッダー
            tds = tr.find_all("td")
            if len(tds) < 10:  # 生徒行は多数の列がある
                continue

            # No, 学年, 生徒ID, 名前, チェックボックス, 出欠状態, ..., 備考
            student_name = tds[3].get_text(strip=True) if len(tds) > 3 else ""
            student_id = tds[2].get_text(strip=True) if len(tds) > 2 else ""

            # チェックボックスの状態を確認
            checkbox_td = tds[4] if len(tds) > 4 else None
            is_attended = False
            if checkbox_td:
                checkbox = checkbox_td.find("input", {"type": "checkbox"})
                if checkbox and checkbox.get("checked") is not None:
                    is_attended = True

            # 出欠状態のテキスト（"出席"、"欠席"など）
            attendance_status = tds[5].get_text(strip=True) if len(tds) > 5 else ""

            # 備考欄（振替などの情報）
            memo = tds[8].get_text(strip=True) if len(tds) > 8 else ""

            # 集計
            if is_attended or attendance_status == "出席":
                result["attendance_count"] += 1
                # 備考に「振替」が含まれている場合
                if "振替" in memo or "振り替え" in memo:
                    result["substitution_count"] += 1
                    result["attendance_count_substitution"] += 1
                    print(f"      [DEBUG] 生徒振替: {student_name} (備考: {memo})")
                else:
                    result["attendance_count_regular"] += 1
            elif attendance_status == "欠席":
                result["absent_count"] += 1
                print(f"      [DEBUG] 生徒欠席: {student_name}")

            # 生徒詳細情報を保存
            result["students"].append({
                "name": student_name,
                "student_id": student_id,
                "is_attended": is_attended,
                "status": attendance_status,
                "memo": memo
            })

    return result


def map_end_time(start_time: Optional[str]) -> Optional[str]:
    mapping = {"16:05": "16:55", "17:00": "17:50", "17:55": "18:45"}
    if not start_time:
        return None
    if start_time in mapping:
        return mapping[start_time]
    try:
        hh, mm = map(int, start_time.split(":"))
        total = hh * 60 + mm + 50
        hh2, mm2 = divmod(total, 60)
        return f"{hh2:02d}:{mm2:02d}"
    except Exception:
        return None


def fetch_class_detail_links(html: str, base_url: str) -> List[Tuple[str, str]]:
    """
    出欠一覧ページから各クラスの詳細ページへのリンクを抽出

    Returns:
        List[Tuple[str, str]]: [(class_name, detail_url), ...]
    """
    soup = BeautifulSoup(html, "lxml")
    links = []

    # TblDataListというIDのテーブルが複数ある可能性があるので、全部取得
    # toDayAttendanceSeach.aspxページのクラス一覧テーブルを探す
    # テーブル構造: 日付区分 | 時間帯 | クラス名 | 本予定人数 | 体験人数 | 合計
    tables = soup.find_all("table", id="TblDataList")

    print(f"    [DEBUG] TblDataListテーブル数: {len(tables)}")

    for table_idx, table in enumerate(tables):
        print(f"    [DEBUG] テーブル {table_idx + 1} を解析中...")
        rows = table.find_all("tr")

        # 「授業予定はありません」のチェック
        if rows and len(rows) > 1:
            second_row_text = rows[1].get_text(strip=True)
            if "授業予定はありません" in second_row_text:
                print(f"    [DEBUG] このテーブルには授業予定がありません。スキップ")
                continue

        # ヘッダー行を確認
        # クラス一覧テーブルの構造:
        # 行0: <th colspan='6'>日付　出欠管理</th>
        # 行1: <th>日付区分</th><th>時間帯</th><th>クラス名</th>...
        # 行2以降: データ行
        header_row_idx = None
        if rows and len(rows) > 0:
            for idx, row in enumerate(rows):
                header_cells = row.find_all("th")
                header_text = [th.get_text(strip=True) for th in header_cells]

                if idx == 0:
                    print(f"    [DEBUG] 行{idx}のヘッダー: {header_text}")

                # クラス名（または「名称」）の列が含まれているかチェック
                if "クラス名" in header_text or "名称" in header_text:
                    header_row_idx = idx
                    print(f"    [DEBUG] 行{idx}にクラス一覧のヘッダーを発見: {header_text}")
                    break

        if header_row_idx is None:
            print(f"    [DEBUG] このテーブルはクラス一覧ではありません。スキップ")
            continue

        # ヘッダー行の次からデータ行を処理
        for row_idx, row in enumerate(rows[header_row_idx + 1:], 1):
            tds = row.find_all("td")
            if len(tds) < 3:
                continue

            # クラス名のセル（3列目、インデックス2）からリンクを探す
            class_cell = tds[2]
            link = class_cell.find("a")
            if link:
                class_name = link.get_text(strip=True)

                # 2つのリンクパターンに対応
                # パターン1: <a href="toDayAttendanceDetail.aspx?...">
                # パターン2: <a href='#' onclick="callPlanDetail('2','4200','1','20251001','1');">
                href = link.get("href", "")
                onclick = link.get("onclick", "")

                print(f"    [DEBUG] リンク発見 (行{row_idx}): クラス名='{class_name}'")
                print(f"    [DEBUG]   href='{href}', onclick='{onclick}'")

                detail_url = None

                # JavaScriptのonclick属性からURLを構築
                if onclick and "callPlanDetail" in onclick:
                    # callPlanDetail('2','4200','1','20251001','1')から引数を抽出
                    import re
                    match = re.search(r"callPlanDetail\('([^']+)','([^']+)','([^']+)','([^']+)','([^']+)'\)", onclick)
                    if match:
                        crsIdx, cgpIdx, sclIdx, planDate, sttTime = match.groups()
                        detail_url = f"{base_url}toDayAttendanceDetail.aspx?crsIdx={crsIdx}&cgpIdx={cgpIdx}&sclIdx={sclIdx}&planDate={planDate}&sttTime={sttTime}"
                        print(f"    [DEBUG]   JavaScriptから構築したURL: '{detail_url}'")

                # 通常のhref属性からURL構築
                elif href and href != "#" and not href.startswith("javascript:"):
                    if not href.startswith("http"):
                        if not href.startswith("/"):
                            detail_url = base_url + href
                        else:
                            detail_url = "https://www.o-za.jp" + href
                    else:
                        detail_url = href
                    print(f"    [DEBUG]   hrefから構築したURL: '{detail_url}'")

                if detail_url:
                    links.append((class_name, detail_url))
                else:
                    print(f"    [DEBUG]   URLを構築できませんでした")

    return links


def fetch_class_detail(session: requests.Session, detail_url: str) -> dict:
    """
    クラス詳細ページにアクセスして情報を取得

    Args:
        session: requests.Session
        detail_url: 詳細ページのURL

    Returns:
        dict: parse_class_detail()の戻り値
    """
    resp = session.get(detail_url)
    resp.raise_for_status()
    return parse_class_detail(resp.text)


def fetch_one_day(session: requests.Session, cfg: Config, soup: BeautifulSoup, day: date, course_id: int, school_id: str, fetch_details: bool = False) -> Tuple[List[dict], BeautifulSoup, List[dict]]:
    # 日付/校舎/ブランドをセットして「表示」押下
    extra = {
        "ctl00$CPH$txtTargetDate": day.strftime("%Y/%m/%d"),
        "ctl00$CPH$ddlSeachCourseID": str(course_id),
        "ctl00$CPH$ddlSeachSchoolID": str(school_id),
        "ctl00$CPH$txtSeachCGP_INDEX": "ALL",
    }
    resp = aspnet_post(session, cfg.attendance_url, soup,
                       event_target="ctl00$CPH$btnSeach",
                       event_argument="",
                       extra_form=extra)
    new_soup = BeautifulSoup(resp.text, "lxml")
    rows = parse_attendance_table(resp.text)
    for r2 in rows:
        r2["date"] = day.isoformat()
        r2["school_id"] = school_id
        r2["course_id"] = course_id

    # クラス詳細情報を取得（オプション）
    details = []
    if fetch_details:
        links = fetch_class_detail_links(resp.text, cfg.base_url)
        print(f"  [DEBUG] 見つかったクラス数: {len(links)}")
        for class_name, detail_url in links:
            try:
                print(f"  [DEBUG] アクセス中: {detail_url}")
                detail_info = fetch_class_detail(session, detail_url)
                detail_info["date"] = day.isoformat()
                detail_info["school_id"] = school_id
                detail_info["course_id"] = course_id
                details.append(detail_info)
                teacher_info = f"講師ID={detail_info['teacher_id']}, 講師名={detail_info['teacher_name']}({detail_info['teacher_attendance']})"
                if detail_info.get('teacher_memo'):
                    teacher_info += f" [備考: {detail_info['teacher_memo']}]"
                print(f"  → {class_name}: {teacher_info}, 生徒出席={detail_info['attendance_count_regular']}, 生徒振替={detail_info['attendance_count_substitution']}, 生徒欠席={detail_info['absent_count']}")
                time.sleep(0.3)  # サーバーへの負荷を軽減
            except Exception as e:
                import traceback
                print(f"  [WARN] クラス詳細取得エラー {class_name}: {e}")
                print(f"  [DEBUG] エラー詳細:\n{traceback.format_exc()}")

    return rows, new_soup, details


def aggregate_active_slots(rows: List[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(columns=["date", "school_name", "start_time", "end_time", "has_class"])
    grp = df.groupby(["date", "school_name", "start_time"], as_index=False).agg({"has_class": "max"})
    grp["end_time"] = grp["start_time"].apply(map_end_time)
    return grp[["date", "school_name", "start_time", "end_time", "has_class"]]


def to_tslot(active_df: pd.DataFrame) -> pd.DataFrame:
    if active_df.empty:
        return pd.DataFrame(columns=["slot_date", "campus_name", "slot_start", "slot_end", "memo"])
    use = active_df[active_df["has_class"] == True].copy()
    use.rename(columns={
        "date": "slot_date",
        "school_name": "campus_name",
        "start_time": "slot_start",
        "end_time": "slot_end",
    }, inplace=True)
    use["memo"] = "scraped"
    return use[["slot_date", "campus_name", "slot_start", "slot_end", "memo"]]


def _chunk(rows: List[dict], size: int = 200) -> List[List[dict]]:
    return [rows[i:i + size] for i in range(0, len(rows), size)]


def normalize_start_time(value: Optional[str]) -> Optional[str]:
    if not value:
        return None
    text = str(value).strip()
    match = re.search(r"(\d{1,2}:\d{2})", text)
    if match:
        return match.group(1)
    return text or None


def prepare_detail_rows(details: List[dict]) -> List[dict]:
    prepared = []
    for detail in details:
        teacher_id = str(detail.get("teacher_id") or "").strip()
        teacher_attendance = str(detail.get("teacher_attendance") or "").strip()
        if not teacher_id:
            continue
        if teacher_attendance and "出席" not in teacher_attendance:
            continue
        attendance_count = detail.get("attendance_count")
        try:
            attendance_total = int(float(attendance_count)) if attendance_count is not None else 0
        except Exception:
            attendance_total = 0
        raw_work_type = detail.get("work_type")
        if raw_work_type is not None:
            work_type = str(raw_work_type).strip() or "授業"
        else:
            work_type = "授業" if attendance_total > 0 else "待機"

        start_time = normalize_start_time(detail.get("start_time"))
        prepared.append({
            "date": detail.get("date"),
            "school_id": detail.get("school_id"),
            "school_name": detail.get("school_name"),
            "class_name": detail.get("class_name"),
            "course_id": detail.get("course_id"),
            "start_time": start_time,
            "teacher_id": teacher_id,
            "teacher_name": detail.get("teacher_name"),
            "teacher_attendance": teacher_attendance or '出席',
            "attendance_count": attendance_total,
            "work_type": work_type,
        })
    return prepared


def push_to_gas(details: List[dict], webhook_url: str, api_key: str, batch_size: int = 200, timeout: int = 30):
    """Apps Script Webアプリに講師勤怠データを送信する"""
    payload_rows = prepare_detail_rows(details)
    if not payload_rows:
        print("[gas] 送信対象データがありません (講師出席データなし)")
        return

    batches = _chunk(payload_rows, batch_size)
    for idx, batch in enumerate(batches, start=1):
        payload = {
            "apiKey": api_key,
            "rows": batch,
        }
        resp = requests.post(webhook_url, json=payload, timeout=timeout)
        try:
            resp.raise_for_status()
        except Exception as e:
            print(f"[gas] バッチ{idx}/{len(batches)} の送信に失敗: {e}")
            print(f"[gas] 応答本文: {resp.text[:500]}")
            raise
        print(f"[gas] バッチ{idx}/{len(batches)} を送信 status={resp.status_code}")


def load_toml(path: str) -> Optional[dict]:
    if not path or not os.path.exists(path):
        return None
    if tomllib is None:
        raise RuntimeError("Python 3.11+ is required to parse TOML, or install 'tomli'.")
    with open(path, "rb") as f:
        return tomllib.load(f)


def main():
    ap = argparse.ArgumentParser(description="OZA 出欠ページを『UIフロー通り』に操作して授業枠を抽出")
    ap.add_argument("--config", help="TOML設定ファイル (oza_scraper.toml など)")
    ap.add_argument("--month", default="auto", help="YYYY-MM または YYYYMM。auto=前月")
    ap.add_argument("--school-ids", default="auto", help='例: "1,17,20"。auto=DDLから全取得')
    ap.add_argument("--course-ids", default=None, help='ブランドID（カンマ区切りで複数指定可能）例: "2,145"。未指定なら設定値/既定=145')
    ap.add_argument("--out", default=None, help="出力Excel。未指定は attendance_sessions_YYYYMM.xlsx")
    ap.add_argument("--skip-workend", action="store_true", help="退勤ボタンのクリックをスキップ")
    ap.add_argument("--fetch-details", action="store_true", help="各クラスの詳細情報（講師名、出席数など）を取得")
    ap.add_argument("--gas-webhook", help="Apps Script WebアプリのURL。指定すると取得結果をGASへPOST")
    ap.add_argument("--gas-api-key", help="Apps Script側で検証するAPIキー (X-GAS-KEY)" )
    args = ap.parse_args()
    args.month = resolve_month_arg(args.month)
    print(f"[info] 対象年月: {args.month}")

    cfg = Config()
    conf_dict = load_toml(args.config) if args.config else None
    if conf_dict:
        sec = conf_dict.get("oza", {})
        cfg.base_url = sec.get("base_url", cfg.base_url)
        cfg.login_url = sec.get("login_url", cfg.login_url)
        cfg.attendance_url = sec.get("attendance_url", cfg.attendance_url)
        cfg.clock_url = sec.get("clock_url", cfg.clock_url)
        cfg.username = sec.get("username", cfg.username)
        cfg.password = sec.get("password", cfg.password)
        cfg.course_ids = sec.get("course_ids", cfg.course_ids)
        cfg.user_field = sec.get("user_field", cfg.user_field)
        cfg.pass_field = sec.get("pass_field", cfg.pass_field)
        cfg.login_button = sec.get("login_button", cfg.login_button)

    cfg.username = os.environ.get("OZA_USERNAME", cfg.username)
    cfg.password = os.environ.get("OZA_PASSWORD", cfg.password)
    env_course_ids = os.environ.get("OZA_COURSE_IDS")
    if env_course_ids:
        cfg.course_ids = env_course_ids

    # course_idsの処理
    if args.course_ids:
        course_ids = [int(x.strip()) for x in args.course_ids.split(",") if x.strip()]
    else:
        course_ids = [int(x.strip()) for x in cfg.course_ids.split(",") if x.strip()]

    print(f"[info] 対象ブランドID: {course_ids}")

    s = requests.Session()
    s.headers.update({"User-Agent": cfg.user_agent})

    # 1) ログイン
    if not login(s, cfg, verbose=True):
        print("[ERROR] ログイン失敗（ヒューリスティック）", file=sys.stderr)

    # 2) 退勤ボタンクリック（任意）
    if not args.skip_workend:
        try:
            click_work_end(s, cfg, verbose=True)
        except Exception as e:
            print(f"[WARN] work_end: {e}", file=sys.stderr)

    start, end = yyyymm_to_range(args.month)
    all_rows: List[dict] = []
    all_details: List[dict] = []

    # 各course_idに対してループ
    for course_id in course_ids:
        print(f"\n[info] ブランドID {course_id} の処理を開始")

        # 3) 生徒出欠簿を開く
        soup = open_attendance(s, cfg)

        # 4) ブランド変更（DDLポストバック）
        try:
            soup = change_course(s, cfg, soup, course_id)
        except Exception as e:
            print(f"[WARN] change_course: {e}", file=sys.stderr)
            continue

        # 5) 校舎リスト
        if args.school_ids.strip().lower() == "auto":
            opts = parse_school_options_from_soup(soup)
            school_ids = [val for val, _ in opts if val.isdigit()]
            print(f"[info] ブランドID {course_id} - 取得した校舎数: {len(school_ids)}")
        else:
            school_ids = [x.strip() for x in args.school_ids.split(",") if x.strip()]

        for sid in school_ids:
            d = start
            day_soup = soup  # __VIEWSTATEを保つ
            while d <= end:
                try:
                    rows, day_soup, details = fetch_one_day(s, cfg, day_soup, d, course_id, sid, fetch_details=args.fetch_details)
                    all_rows.extend(rows)
                    all_details.extend(details)
                    print(f"[{d}] course_id={course_id}, school_id={sid} rows={len(rows)}, details={len(details)}")
                    time.sleep(0.2)
                except Exception as e:
                    print(f"[WARN] {d} course_id={course_id}, school_id={sid}: {e}", file=sys.stderr)
                d += timedelta(days=1)

    raw_df = pd.DataFrame(all_rows)
    active_df = aggregate_active_slots(all_rows)
    tslot_df = to_tslot(active_df)

    yyyymm = args.month.replace("-", "")
    out_path = args.out or f"attendance_sessions_{yyyymm}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        raw_df.to_excel(xw, sheet_name="Raw", index=False)
        active_df.to_excel(xw, sheet_name="ActiveSlots", index=False)
        tslot_df.to_excel(xw, sheet_name="T_Slot", index=False)

        # 詳細情報シートを追加
        if all_details:
            # 生徒情報を展開せずに集計情報のみを出力
            details_summary = []
            for detail in all_details:
                details_summary.append({
                    "date": detail.get("date"),
                    "course_id": detail.get("course_id"),
                    "school_name": detail.get("school_name"),
                    "school_id": detail.get("school_id"),
                    "class_name": detail.get("class_name"),
                    "start_time": detail.get("start_time"),
                    "teacher_id": detail.get("teacher_id"),
                    "teacher_name": detail.get("teacher_name"),
                    "teacher_attendance": detail.get("teacher_attendance"),
                    "teacher_memo": detail.get("teacher_memo"),
                    "attendance_count": detail.get("attendance_count"),
                    "attendance_count_regular": detail.get("attendance_count_regular"),
                    "attendance_count_substitution": detail.get("attendance_count_substitution"),
                    "absent_count": detail.get("absent_count"),
                    "total_students": len(detail.get("students", [])),
                })
            details_df = pd.DataFrame(details_summary)
            details_df.to_excel(xw, sheet_name="ClassDetails", index=False)

            # 生徒詳細情報も別シートに出力
            student_details = []
            for detail in all_details:
                for student in detail.get("students", []):
                    student_details.append({
                        "date": detail.get("date"),
                        "course_id": detail.get("course_id"),
                        "school_name": detail.get("school_name"),
                        "class_name": detail.get("class_name"),
                        "teacher_name": detail.get("teacher_name"),
                        "student_name": student.get("name"),
                        "student_id": student.get("student_id"),
                        "status": student.get("status"),
                        "memo": student.get("memo"),
                    })
            if student_details:
                students_df = pd.DataFrame(student_details)
                students_df.to_excel(xw, sheet_name="StudentDetails", index=False)

    print(f"[OK] Exported: {out_path}")

    gas_webhook = args.gas_webhook or os.environ.get("GAS_WEBHOOK")
    gas_api_key = args.gas_api_key or os.environ.get("GAS_API_KEY")
    if gas_webhook:
        if not gas_api_key:
            raise SystemExit("GAS 連携には GAS_API_KEY (環境変数または --gas-api-key) を指定してください")
        if not args.fetch_details:
            raise SystemExit("GAS 連携には --fetch-details を指定して ClassDetails を取得してください")
        push_to_gas(all_details, gas_webhook, gas_api_key)


if __name__ == "__main__":
    main()
