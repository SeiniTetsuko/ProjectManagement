
#!/usr/bin/env python
# -*- coding: utf-8 -*-

import builtins
import sys
import re
import csv
import json
import time
import traceback
from io import StringIO
from datetime import datetime
from collections import OrderedDict
from io import BytesIO
from openpyxl import load_workbook

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from zenv import get_zdkit_env
from zdbase import ZFile  # 保留平台兼容，不直接依赖


# =============================================================================
# 日志编码兜底
# =============================================================================

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass


# =============================================================================
# print flush patch
# =============================================================================

if not getattr(builtins.print, "_patched_flush", False):
    _original_print = builtins.print

    def print(*args, **kwargs):
        kwargs.setdefault("flush", True)
        _original_print(*args, **kwargs)

    print._patched_flush = True
    builtins.print = print


# =============================================================================
# 固定 API 路由
# =============================================================================

TODO_BASE_URL = "http://share.cxmt.com"
TODO_LIST_API = "/openapi/todo/todoList"

TOKEN_API = "/api/user/platform/getAccessToken"
DOC_TREE_ROUTE = "/platform/api/main/doc/treeList"
SIGNED_URL_ROUTE = "/platform/api/main/storage/getSignedUrl"
GET_DOC_API = "/platform/ws/noteInfo/getDocJson"
MD_INSERT_ROUTE = "/middle/server/api/file/md/insert"


# =============================================================================
# 固定业务参数
# =============================================================================

TODO_STATUS_LIST = ["Open", "Ongoing", "Delay"]
TODO_IS_LEAF = False
TODO_TYPE = "wbs"

DAILY_SECTION_TITLE = "TodoList任务进展"
DAILY_TITLE_KEYWORDS = ["日报"]

DAILY_INSERT_MODE = "a"
DAILY_INSERT_LOCATION = 1

MAPPING_DEPT_FIELD = "Dept"
MAPPING_PROJECT_GUID_FIELD = "project_guid"
MAPPING_WORK_LOG_FOLDER_GUID_FIELD = "work_log_folder_guid"
MAPPING_PROJECT_NAME_FIELD = "project_name"

SKIP_EXISTING_TODO = True

RISK_MAP_EN_TO_CN = {
    "high": "高",
    "middle": "中",
    "low": "低"
}

RISK_MAP_CN_TO_EN = {
    "高": "high",
    "中": "middle",
    "低": "low",
}


# =============================================================================
# 全局配置加载
# =============================================================================

zenv_obj = get_zdkit_env()
BASE_URL = zenv_obj.zdkit._http_client.config.get("url")

try:
    with open(config_file.path, "r", encoding="utf-8") as config_fp:
        config = json.load(config_fp)
except Exception as e:
    print(f"❌ 配置文件读取失败: {e}")
    raise

AK = config.get("ak")
SK = config.get("sk")
ORG_GUID = config.get("org_guid")
USER_GUID = config.get("user_guid")

MAPPING_GUID = config.get("mapping_guid")
TARGET_DATE = config.get("target_date") or datetime.now().strftime("%Y-%m-%d")
TODO_PROJECT_ID = config.get("todo_project_id")
if isinstance(TODO_PROJECT_ID, str):
    TODO_PROJECT_ID = [TODO_PROJECT_ID]

if not AK or not SK:
    print("❌ 配置错误：缺少 ak / sk")
    raise ValueError("配置错误：缺少 ak / sk")

if not USER_GUID:
    print("❌ 配置错误：缺少 user_guid")
    raise ValueError("配置错误：缺少 user_guid")

if not TODO_PROJECT_ID:
    print("❌ 配置错误：缺少 todo_project_id")
    raise ValueError("配置错误：缺少 todo_project_id")

if not MAPPING_GUID:
    print("❌ 配置错误：缺少 mapping_guid")
    raise ValueError("配置错误：缺少 mapping_guid")


# =============================================================================
# 通用 HTTP 工具
# =============================================================================

session = requests.Session()

retry_strategy = Retry(
    total=5,
    connect=5,
    read=5,
    backoff_factor=2,
    status_forcelist=[429, 500, 502, 503, 504],
    allowed_methods=["HEAD", "GET", "POST", "PUT", "DELETE", "OPTIONS"],
    raise_on_status=False
)

adapter = HTTPAdapter(
    max_retries=retry_strategy,
    pool_connections=10,
    pool_maxsize=10
)

session.mount("http://", adapter)
session.mount("https://", adapter)

def request_with_retry(method, url, max_retries=3, **kwargs):
    """
    稳定版 HTTP 请求：

    1. 自动 retry
    2. Connection reset 自动恢复
    3. 避免复用坏连接
    4. 支持长时间下载
    """

    kwargs.setdefault("timeout", 30)

    last_error = None

    for attempt in range(1, max_retries + 1):

        try:

            headers = kwargs.pop("headers", {}) or {}

            # 核心：不要复用失效 keep-alive
            headers["Connection"] = "close"

            # 某些网关会校验 UA
            headers.setdefault(
                "User-Agent",
                "Mozilla/5.0 PythonRequests"
            )

            if method.lower() == "post":

                response = session.post(
                    url,
                    headers=headers,
                    **kwargs
                )

            else:

                response = session.get(
                    url,
                    headers=headers,
                    **kwargs
                )

            return response

        except (
            requests.exceptions.Timeout,
            requests.exceptions.ConnectionError,
            requests.exceptions.ChunkedEncodingError,
        ) as e:

            last_error = e

            wait = min(2 ** attempt, 10)

            print(
                f"⚠️ 请求失败 "
                f"attempt={attempt}/{max_retries} "
                f"wait={wait}s "
                f"url={url} "
                f"error={repr(e)}"
            )

            if attempt < max_retries:
                time.sleep(wait)
            else:
                raise last_error

def get_access_token():

    print("🔑 获取 AccessToken")

    resp = request_with_retry(
        "post",
        BASE_URL + TOKEN_API,
        json={
            "ak": AK,
            "sk": SK
        },
        timeout=30
    )

    resp_json = resp.json()

    token = (
        resp_json.get("data", {})
        .get("accessToken")
    )

    if not token:
        raise Exception(
            f"获取 token 失败: "
            f"url={BASE_URL + TOKEN_API}, "
            f"response={resp_json}"
        )

    print("✅ AccessToken 获取成功")

    return token


def get_headers(user_guid=None):
    """
    BASE_URL (workspace) 专用 header。

    默认使用配置里的 USER_GUID。
    写入日报时可传入 Todo owner 的 userGuid，
    以该 owner 身份完成 Workspace 笔记写入。
    """
    token = get_access_token()

    return {
        "Access-Token": token,
        "ak": AK,
        "X-User-GUID": user_guid or USER_GUID,
        "Content-Type": "application/json"
    }


def get_todo_headers():
    """TODO_BASE_URL (share) 专用 header"""
    return {
        "x-user-guid": USER_GUID,
        "Content-Type": "application/json"
    }


# =============================================================================
# mapping 文件读取
# =============================================================================

def get_signed_url(category_guid):

    print(f"🔗 获取签名 URL: category_guid={category_guid}")

    response = request_with_retry(
        "get",
        BASE_URL + SIGNED_URL_ROUTE,
        headers=get_headers(),
        params={
            "categoryGuid": category_guid
        },
        timeout=30
    )

    response_json = response.json()
    signed_url = (response_json.get("data") or {}).get("signedUrl")

    if not signed_url:
        raise Exception(
            f"获取签名 URL 失败: "
            f"category_guid={category_guid}, "
            f"response={response_json}"
        )

    return signed_url


def normalize_dept_name(dept_name):
    """
    统一部门路径格式：
    - 去掉空格
    - 统一斜杠
    - 去掉首尾斜杠
    - 小写

    例如：
    AIX/AIX 2/AIX 2 -> aix/aix2/aix2
    AIX\\AIX2\\AIX2 -> aix/aix2/aix2
    """
    text = str(dept_name or "").strip()
    text = text.replace("\\", "/")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"/+", "/", text)
    text = text.strip("/")
    return text.lower()

def parse_xlsx_mapping(content_bytes):
    """
    读取 xlsx mapping。

    必须字段：
    Dept,project_guid,work_log_folder_guid

    可选字段：
    project_name
    """

    workbook = load_workbook(
        filename=BytesIO(content_bytes),
        read_only=True,
        data_only=True
    )

    sheet = workbook.active

    rows_iter = sheet.iter_rows(values_only=True)

    try:
        headers = next(rows_iter)
    except StopIteration:
        raise ValueError("xlsx 文件为空")

    headers = [str(x).strip() if x is not None else "" for x in headers]

    rows = []

    for excel_row in rows_iter:
        row_dict = {}

        for idx, value in enumerate(excel_row):
            if idx >= len(headers):
                continue

            key = headers[idx]
            row_dict[key] = str(value).strip() if value is not None else ""

        dept = row_dict.get(MAPPING_DEPT_FIELD, "")
        project_guid = row_dict.get(MAPPING_PROJECT_GUID_FIELD, "")
        work_log_folder_guid = row_dict.get(MAPPING_WORK_LOG_FOLDER_GUID_FIELD, "")
        project_name = row_dict.get(MAPPING_PROJECT_NAME_FIELD, "")

        if not dept:
            print(f"⚠️ 跳过 mapping 行：缺少 Dept -> {row_dict}")
            continue

        if not project_guid:
            print(f"⚠️ 跳过 mapping 行：缺少 project_guid -> {row_dict}")
            continue

        if not work_log_folder_guid:
            print(f"⚠️ 跳过 mapping 行：缺少 work_log_folder_guid -> {row_dict}")
            continue

        rows.append({
            "dept": dept,
            "dept_norm": normalize_dept_name(dept),
            "project_guid": project_guid,
            "work_log_folder_guid": work_log_folder_guid,
            "project_name": project_name
        })

    if not rows:
        raise ValueError("xlsx 文件中没有有效记录")

    return rows

def parse_csv_mapping(text):
    """
    读取 CSV mapping。

    必须字段：
    Dept,project_guid,work_log_folder_guid

    可选字段：
    project_name
    """

    text = text.lstrip("\ufeff").strip()

    if not text:
        raise ValueError("mapping 文件内容为空")

    text = text.replace("\x00", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    stream = StringIO(text, newline="")

    reader = csv.DictReader(
        stream,
        delimiter=",",
        quotechar='"',
        skipinitialspace=True
    )

    rows = []

    for row in reader:
        if not row:
            continue

        dept = (row.get(MAPPING_DEPT_FIELD) or "").strip()
        project_guid = (
            row.get(MAPPING_PROJECT_GUID_FIELD) or ""
        ).strip()

        work_log_folder_guid = (
            row.get(MAPPING_WORK_LOG_FOLDER_GUID_FIELD) or ""
        ).strip()

        project_name = (
            row.get(MAPPING_PROJECT_NAME_FIELD) or ""
        ).strip()

        if not dept:
            print(f"⚠️ 跳过 mapping 行：缺少 Dept -> {row}")
            continue

        if not project_guid:
            print(
                f"⚠️ 跳过 mapping 行："
                f"缺少 project_guid -> {row}"
            )
            continue

        if not work_log_folder_guid:
            print(
                f"⚠️ 跳过 mapping 行："
                f"缺少 work_log_folder_guid -> {row}"
            )
            continue

        rows.append({
            "dept": dept,
            "dept_norm": normalize_dept_name(dept),
            "project_guid": project_guid,
            "work_log_folder_guid": work_log_folder_guid,
            "project_name": project_name
        })

    if not rows:
        raise ValueError("mapping 文件中没有有效记录")

    return rows


def load_mapping_from_guid(mapping_guid):

    print(f"🚀 开始读取 mapping 文件: {mapping_guid}")

    signed_url = get_signed_url(mapping_guid)

    response = request_with_retry(
        "get",
        signed_url,
        timeout=120,
        stream=True,
        headers={
            "Connection": "close"
        }
    )

    if response.status_code != 200:
        raise Exception(
            f"下载 mapping 文件失败: "
            f"status={response.status_code}, "
            f"text={response.text[:300]}"
        )

    # 手动读取，避免流式中断
    content = response.content

    content_type = (
        response.headers.get("Content-Type", "")
        .lower()
    )

    print(f"📄 Content-Type: {content_type}")

    signed_url_lower = signed_url.lower()

    try:

        # 优先按 URL 后缀判断
        if ".xlsx" in signed_url_lower:

            print("📘 检测为 XLSX 文件")

            mapping_rows = parse_xlsx_mapping(content)

        elif ".csv" in signed_url_lower:

            print("📄 检测为 CSV 文件")

            text = content.decode(
                "utf-8-sig",
                errors="ignore"
            )

            mapping_rows = parse_csv_mapping(text)

        # fallback 按 content-type
        elif (
            "spreadsheetml" in content_type
            or "excel" in content_type
        ):

            print("📘 根据 Content-Type 检测为 XLSX 文件")

            mapping_rows = parse_xlsx_mapping(content)

        else:

            print("📄 默认按 CSV 解析")

            text = content.decode(
                "utf-8-sig",
                errors="ignore"
            )

            mapping_rows = parse_csv_mapping(text)

    except Exception:

        print("❌ mapping 文件解析失败")
        print(f"Content-Type: {content_type}")
        print(f"signed_url: {signed_url}")

        try:
            preview = content[:500]
            print(f"content[:500]={preview}")
        except Exception:
            pass

        raise

    print(f"✅ mapping 文件读取成功，有效记录数: {len(mapping_rows)}")

    return mapping_rows

def find_mapping_by_dept(full_dept_name, mapping_rows):
    """
    用 todo.owner.fullDeptName 匹配 mapping.Dept。

    当前策略：
    1. 标准化后完全匹配
    2. 如果完全匹配失败，尝试前缀匹配
    """
    dept_norm = normalize_dept_name(full_dept_name)

    if not dept_norm:
        return None

    # 1. 完全匹配
    for row in mapping_rows:
        if row["dept_norm"] == dept_norm:
            return row

    # 2. mapping 比 fullDept 更短时，允许 fullDept 以 mapping 开头
    candidates = []
    for row in mapping_rows:
        map_dept_norm = row["dept_norm"]
        if dept_norm.startswith(map_dept_norm + "/") or dept_norm == map_dept_norm:
            candidates.append(row)

    if candidates:
        candidates.sort(key=lambda x: len(x["dept_norm"]), reverse=True)
        return candidates[0]

    # 3. fullDept 比 mapping 更短时，允许 mapping 以 fullDept 开头
    candidates = []
    for row in mapping_rows:
        map_dept_norm = row["dept_norm"]
        if map_dept_norm.startswith(dept_norm + "/"):
            candidates.append(row)

    if candidates:
        candidates.sort(key=lambda x: len(x["dept_norm"]))
        return candidates[0]

    return None


# =============================================================================
# Todo API
# =============================================================================

def fetch_todo_list():
    """
    按 projectIds 拉取全量 Todo，不再按 owner 逐个查询。
    如需按 owner 过滤，在本地过滤即可。
    """
    print(f"🚀 获取 Todo 列表: projectIds={TODO_PROJECT_ID}")

    request_body = {
        "projectIds": TODO_PROJECT_ID,
        "todoStatus": TODO_STATUS_LIST,
        "isLeaf": TODO_IS_LEAF,
        "type": TODO_TYPE,
    }

    response = request_with_retry(
        "post",
        TODO_BASE_URL + TODO_LIST_API,
        headers=get_todo_headers(),
        json=request_body,
        timeout=60
    )

    response.raise_for_status()
    response_json = response.json()

    if response_json.get("code") not in (0, 200, "0", "200", None):
        print(f"⚠️ Todo API 返回 code 异常: {response_json.get('code')}")

    todo_list = response_json.get("data") or []

    print(f"✅ 获取 Todo 数量: {len(todo_list)}")

    return todo_list


# =============================================================================
# Workspace 日报查找
# =============================================================================

def get_tree_node_title(node):
    for key in ("dataTitle", "title", "name", "fileName", "filename"):
        value = node.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def get_tree_node_guid(node):
    for key in ("categoryGuid", "dataGuid", "guid", "fileGuid", "id"):
        value = node.get(key)
        if value:
            return str(value)
    return ""


def get_date_title_variants(date_str):
    """
    2026-05-11 ->
    2026-05-11
    2026/05/11
    20260511
    2026.05.11
    2026年05月11日
    2026年5月11日
    """
    dt = datetime.strptime(date_str, "%Y-%m-%d")

    return [
        dt.strftime("%Y-%m-%d"),
        dt.strftime("%Y/%m/%d"),
        dt.strftime("%Y%m%d"),
        dt.strftime("%Y.%m.%d"),
        dt.strftime("%Y年%m月%d日"),
        f"{dt.year}年{dt.month}月{dt.day}日",
    ]


def is_daily_note_title(title, target_date, project_name=""):
    title = title or ""

    date_variants = get_date_title_variants(target_date)

    has_date = any(x in title for x in date_variants)
    if not has_date:
        return False

    has_keyword = any(keyword in title for keyword in DAILY_TITLE_KEYWORDS)
    if not has_keyword:
        return False

    return True


def score_daily_note_title(title, target_date, project_name=""):
    """
    多个候选日报时，用 score 选最像目标日报的。
    """
    score = 0
    title = title or ""

    if is_daily_note_title(title, target_date, project_name):
        score += 100

    if project_name and project_name in title:
        score += 30

    if "副本" in title:
        score -= 20

    if "测试" in title:
        score -= 10

    if title.endswith("日报"):
        score += 5

    return score


def list_folder_nodes(user_guid, project_guid, folder_guid):
    response = request_with_retry(
        "post",
        BASE_URL + DOC_TREE_ROUTE,
        headers=get_headers(),
        json={
            "projectGuid": project_guid,
            "parentGuid": folder_guid
        },
        timeout=60
    )

    if response.status_code != 200:
        raise Exception(
            f"treeList 请求失败: status={response.status_code}, text={response.text[:300]}"
        )

    response_json = response.json()
    data = response_json.get("data") or []

    return data


def find_today_daily_note(user_guid, project_guid, folder_guid, target_date, project_name=""):
    """
    在目标项目日报文件夹下查找 target_date 当天日报。
    """
    print(
        f"🔎 查找日报: project_guid={project_guid}, "
        f"folder_guid={folder_guid}, date={target_date}, project_name={project_name}"
    )

    nodes = list_folder_nodes(
        user_guid=user_guid,
        project_guid=project_guid,
        folder_guid=folder_guid
    )

    candidates = []

    for node in nodes:
        title = get_tree_node_title(node)
        guid = get_tree_node_guid(node)

        if not title or not guid:
            continue

        if not is_daily_note_title(title, target_date, project_name):
            continue

        candidates.append({
            "note_guid": guid,
            "note_title": title,
            "score": score_daily_note_title(title, target_date, project_name),
            "node": node
        })

    if not candidates:
        print(
            f"❌ 未找到当日日报: project_guid={project_guid}, "
            f"folder_guid={folder_guid}, date={target_date}"
        )
        return None

    candidates.sort(key=lambda x: x["score"], reverse=True)

    best = candidates[0]

    print(
        f"✅ 命中日报: {best['note_title']} | "
        f"note_guid={best['note_guid']} | score={best['score']}"
    )

    if len(candidates) > 1:
        print("⚠️ 存在多个候选日报，已选择 score 最高的一个：")
        for item in candidates:
            print(f"    - score={item['score']} | {item['note_title']} | {item['note_guid']}")

    return best


# =============================================================================
# 读取已有日报内容，用于去重
# =============================================================================

def extract_text_from_note_json_node(node, parts):
    if isinstance(node, dict):
        node_type = node.get("type")

        if node_type == "text":
            text = node.get("text", "")
            if text:
                parts.append(text)
            return

        if node_type in (
            "paragraph",
            "heading",
            "fheading",
            "bulletListItem",
            "numberedListItem",
            "codeBlock"
        ):
            for child in node.get("content", []) or []:
                extract_text_from_note_json_node(child, parts)
            parts.append("\n")
            return

        for key in ("text", "code"):
            value = node.get(key)
            if isinstance(value, str) and value.strip():
                parts.append(value)
                parts.append("\n")
                return

        for child in node.get("content", []) or []:
            extract_text_from_note_json_node(child, parts)

    elif isinstance(node, list):
        for child in node:
            extract_text_from_note_json_node(child, parts)


def extract_text_from_note_json(raw_note_json):
    root = raw_note_json.get("data", {}).get("content", []) or raw_note_json.get("content", [])
    parts = []
    extract_text_from_note_json_node(root, parts)
    return "".join(parts).strip()


def get_note_plain_text(user_guid, note_guid):

    print(f"📄 读取已有日报内容: note_guid={note_guid}")

    response = request_with_retry(
        "get",
        BASE_URL + GET_DOC_API,
        headers=get_headers(),
        params={
            "docId": note_guid
        },
        timeout=60
    )

    response_json = response.json()

    try:
        return extract_text_from_note_json(response_json)
    except Exception as e:
        print(f"⚠️ 解析已有日报内容失败，跳过去重检查: note_guid={note_guid}, error={e}")
        return ""


def is_todo_already_written(existing_text, todo_id):
    """
    根据 Todo 唯一 id 去重。

    说明：
    - Todo API 中 id 是唯一 ID。
    - code 是编码，例如 id=50001 时 code=WBS#50001。
    - 日报标题中写入的是【id】，因此这里也用 id 判断是否已写入。
    """
    if todo_id is None or str(todo_id).strip() == "":
        return False

    return f"【{todo_id}】" in existing_text


# =============================================================================
# 构建 note -> owner -> tasks 中间态
# =============================================================================

def build_note_task_map(todo_list, mapping_rows, target_date):
    """
    新版：
    API 已返回 leaf todo。

    每条 todo 本身就是最终任务：
    - id = 当前任务ID
    - pId = 父任务ID
    """

    note_task_map = OrderedDict()
    daily_note_cache = {}

    total_todo = 0
    routed_todo = 0
    skipped_no_owner = 0
    skipped_no_mapping = 0
    skipped_no_daily_note = 0

    for todo in todo_list:

        total_todo += 1

        todo_id = todo.get("id", "")
        todo_code = todo.get("code", "")

        parent_todo_id = todo.get("pId", "")

        title = todo.get("title", "")
        due_date = todo.get("dueDate", "")
        risk_level = todo.get("riskLevel", "")
        risk_text = RISK_MAP_EN_TO_CN.get(risk_level, risk_level)

        owners = todo.get("owners", []) or []

        if not owners:
            skipped_no_owner += 1

            print(
                f"⚠️ 跳过无 owners 的 Todo: "
                f"todo={todo_id}, title={title}"
            )

            continue

        # 一个 leaf todo 可能多个 owner
        for owner in owners:

            owner_guid = owner.get("userGuid", "")
            owner_name = owner.get("name", "")
            owner_display_name = owner.get(
                "displayName",
                owner_name
            )

            owner_full_dept = owner.get(
                "fullDeptName",
                ""
            )

            if not owner_guid:

                skipped_no_owner += 1

                print(
                    f"⚠️ 跳过无 owner.userGuid 的 Todo: "
                    f"todo={todo_id}, title={title}"
                )

                continue

            mapping = find_mapping_by_dept(
                owner_full_dept,
                mapping_rows
            )

            if not mapping:

                skipped_no_mapping += 1

                print(
                    f"⚠️ 未找到部门 mapping，跳过: "
                    f"owner={owner_display_name}, "
                    f"dept={owner_full_dept}, "
                    f"todo={todo_id}"
                )

                continue

            project_guid = mapping["project_guid"]

            work_log_folder_guid = mapping[
                "work_log_folder_guid"
            ]

            project_name = mapping.get(
                "project_name",
                ""
            )

            mapping_dept = mapping["dept"]

            cache_key = (
                f"{project_guid}|"
                f"{work_log_folder_guid}|"
                f"{target_date}"
            )

            if cache_key not in daily_note_cache:

                daily_note = find_today_daily_note(
                    user_guid=USER_GUID,
                    project_guid=project_guid,
                    folder_guid=work_log_folder_guid,
                    target_date=target_date,
                    project_name=project_name
                )

                daily_note_cache[cache_key] = daily_note

            else:

                daily_note = daily_note_cache[cache_key]

            if not daily_note:

                skipped_no_daily_note += 1

                print(
                    f"⚠️ 未找到目标日报，跳过 Todo: "
                    f"owner={owner_display_name}, "
                    f"dept={owner_full_dept}, "
                    f"todo={todo_id}"
                )

                continue

            note_guid = daily_note["note_guid"]
            note_title = daily_note["note_title"]

            if note_guid not in note_task_map:

                note_task_map[note_guid] = {
                    "project_guid": project_guid,
                    "project_name": project_name,
                    "work_log_folder_guid": work_log_folder_guid,
                    "note_guid": note_guid,
                    "note_title": note_title,
                    "dept": mapping_dept,

                    # 同一个日报命中的第一个 owner 作为写入人。
                    # 后续同部门/同日报 owner 不覆盖，避免一个日报多次切换身份。
                    "writer_user_guid": owner_guid,
                    "writer_user_name": owner_display_name or owner_name,

                    "owners": OrderedDict()
                }

            note_owners = note_task_map[note_guid]["owners"]

            if owner_guid not in note_owners:

                note_owners[owner_guid] = {
                    "owner_guid": owner_guid,
                    "owner_name": owner_name,
                    "owner_display_name": owner_display_name,
                    "owner_dept": owner_full_dept,
                    "tasks": []
                }

            related_project = todo.get("relatedProject") or {}

            related_project_name = (
                related_project.get("name", "")
            )

            note_owners[owner_guid]["tasks"].append({

                # id 是 Todo 唯一 ID，用于日报标题展示和后续回写解析。
                # code 只是编码，例如 id=50001 时 code=WBS#50001。
                "todo_id": todo_id,
                "todo_code": todo_code,

                "task_title": title,

                "due_date": due_date,

                # 写入日报时不再用接口返回值预填 promise/risk，
                # 只保留原始字段，便于日志或后续扩展。
                "promise_date": todo.get("promiseDate", ""),

                "risk_level": risk_level,
                "risk_text": risk_text,

                "status": todo.get("status", ""),

                "parent_todo_id": parent_todo_id,

                "related_project_name": related_project_name
            })

            routed_todo += 1

    print("\n================ 路由统计 ================\n")

    print(f"leaf todo 总数: {total_todo}")
    print(f"成功路由: {routed_todo}")
    print(f"跳过：无 owner: {skipped_no_owner}")
    print(f"跳过：无 mapping: {skipped_no_mapping}")
    print(f"跳过：无当日日报: {skipped_no_daily_note}")
    print(f"目标日报数量: {len(note_task_map)}")

    print("\n==========================================\n")

    return note_task_map

# =============================================================================
# Markdown / HTML 构建
# =============================================================================

def escape_html(text):
    text = str(text or "")
    return (
        text
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def normalize_date_for_display(date_value):
    """
    将 Todo 接口返回的日期尽量规范成 YYYY-MM-DD。

    支持：
    - 2026-05-20
    - 2026-05-20 12:00:00
    - 2026/05/20
    - 20260520

    无法识别时原样返回，避免误删信息。
    """
    text = str(date_value or "").strip()

    if not text:
        return ""

    match = re.match(r"^(\d{4}-\d{1,2}-\d{1,2})", text)
    if match:
        try:
            dt = datetime.strptime(match.group(1), "%Y-%m-%d")
            return dt.strftime("%Y-%m-%d")
        except Exception:
            return match.group(1)

    match = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})", text)
    if match:
        year, month, day = match.groups()
        return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"

    match = re.match(r"^(\d{4})(\d{2})(\d{2})$", text)
    if match:
        year, month, day = match.groups()
        return f"{year}-{month}-{day}"

    return text


def build_task_block(task):
    """
    新版 Todo 写入日报展示格式。

    关键调整：
    1. 标题最后一段写 Todo 唯一 id，不写 code。
       例如 id=50001, code=WBS#50001，则日报写【50001】。
    2. 新增红色任务截止时间 dueDate。
    3. 承诺完成时间、风险等级固定展示格式提示，不用接口返回值预填。
    4. 进展占位符改成明确提示，要求用户删除后填写实际进展。
    """

    task_title = escape_html(
        task.get("task_title", "")
    )

    todo_id = escape_html(
        task.get("todo_id", "")
    )

    due_date = escape_html(
        normalize_date_for_display(task.get("due_date", ""))
    )

    related_project_name = escape_html(
        task.get("related_project_name", "")
    )

    title_parts = []

    if related_project_name:
        title_parts.append(
            f"【{related_project_name}】"
        )

    title_parts.append(f"【{task_title}】")

    # 注意：这里传 id，不传 code。
    # code 通常是 WBS#50001，而 id 才是 Todo 唯一 ID。
    title_parts.append(
        f"【{todo_id}】"
    )

    title_html = "&amp;".join(title_parts)

    return f"""
<p>🚀 <strong>{title_html}</strong></p>

<p><strong style="color:red;">&lt;任务截止时间&gt;: {due_date}</strong><br></p>

<p><strong>&lt;承诺完成时间：YYYY-MM-DD&gt;:</strong><br></p>

<p><strong>&lt;风险等级：高/中/低&gt;:</strong><br></p>

<p><strong>&lt;今日进展描述&gt;:</strong></p>

<ul>
<li>请删除本占位符，并填写今日实际工作进展</li>
</ul>
"""

def build_owner_block(owner_info):
    """
    单个负责人下的任务块。

    展示目标：
    ### @负责人

    task1

    task2
    """
    owner_name = escape_html(owner_info.get("owner_name", ""))

    task_blocks = []

    for task in owner_info.get("tasks", []) or []:
        task_blocks.append(build_task_block(task))

    if not task_blocks:
        return ""

    return f"""
<h3>@{owner_name}</h3>

{''.join(task_blocks)}
"""


def build_markdown_for_note(note_info):
    """
    为单个项目当日日报构建写入内容。

    展示目标：
    ## TodoList任务进展

    ### @负责人A

    负责人A的多个 Todo

    ### @负责人B

    负责人B的多个 Todo

    注意：
    - 不展示”项目：xxx”。
    - 不展示 owner_display_name / owner_dept。
    - 一个 note_guid 只生成一个 h2。
    - 同一个 owner 下多个 Todo 合并到一个 h3 下。
    """
    owner_blocks = []

    for owner_guid, owner_info in note_info.get("owners", {}).items():
        block = build_owner_block(owner_info)
        if block.strip():
            owner_blocks.append(block)

    if not owner_blocks:
        return ""

    return f"""
<h2>{DAILY_SECTION_TITLE}</h2>

{''.join(owner_blocks)}

<hr>
"""


# =============================================================================
# 写入日报
# =============================================================================

def filter_existing_tasks(note_guid, note_info):
    """
    如果 SKIP_EXISTING_TODO=True，则读取已有日报内容，过滤已经写过的 Todo。
    """
    if not SKIP_EXISTING_TODO:
        return note_info

    existing_text = get_note_plain_text(USER_GUID, note_guid)

    if not existing_text:
        return note_info

    filtered_note_info = dict(note_info)
    filtered_owners = OrderedDict()

    skipped_count = 0
    kept_count = 0

    for owner_guid, owner_info in note_info.get("owners", {}).items():
        filtered_owner = dict(owner_info)
        filtered_tasks = []

        for task in owner_info.get("tasks", []) or []:
            todo_id = task.get("todo_id", "")

            if is_todo_already_written(existing_text, todo_id):
                skipped_count += 1
                print(
                    f"ℹ️ 已存在，跳过写入: "
                    f"note={note_guid}, todo_id={todo_id}"
                )
                continue

            filtered_tasks.append(task)
            kept_count += 1

        if filtered_tasks:
            filtered_owner["tasks"] = filtered_tasks
            filtered_owners[owner_guid] = filtered_owner

    filtered_note_info["owners"] = filtered_owners

    print(f"📌 note={note_guid} 去重结果：保留 {kept_count} 条，跳过 {skipped_count} 条")

    return filtered_note_info


def insert_markdown_to_note(note_guid, markdown_content, writer_user_guid=None):
    if not markdown_content.strip():
        print(f"ℹ️ note={note_guid} 无可写入内容，跳过")
        return None

    print(
        f"📝 写入日报: note_guid={note_guid}, "
        f"writer_user_guid={writer_user_guid or USER_GUID}"
    )

    payload = {
        "note_guid": note_guid,
        "markdown_content": markdown_content,
        "mode": DAILY_INSERT_MODE,
        "location": DAILY_INSERT_LOCATION
    }

    response = request_with_retry(
        "post",
        BASE_URL + MD_INSERT_ROUTE,
        headers=get_headers(user_guid=writer_user_guid),
        json=payload,
        timeout=60
    )

    if response.status_code != 200:
        raise Exception(
            f"写入笔记失败: "
            f"note_guid={note_guid}, "
            f"status={response.status_code}, "
            f"text={response.text[:500]}"
        )

    response_json = response.json()

    print(f"✅ 写入成功: note_guid={note_guid}")

    return response_json

def write_all_notes(note_task_map):
    success_count = 0
    skip_count = 0
    fail_count = 0

    for note_guid, note_info in note_task_map.items():
        print("\n------------------------------------------")
        print(f"🚀 开始处理日报: {note_info.get('note_title')} | {note_guid}")

        try:
            filtered_note_info = filter_existing_tasks(note_guid, note_info)
            markdown_content = build_markdown_for_note(filtered_note_info)

            print("\n================ 写入内容预览 ================\n")
            print(markdown_content)
            print("\n==============================================\n")

            if not markdown_content.strip():
                print(f"ℹ️ 当前日报无新增任务，跳过: {note_guid}")
                skip_count += 1
                continue

            writer_user_guid = filtered_note_info.get("writer_user_guid")
            writer_user_name = filtered_note_info.get("writer_user_name", "")

            print(
                f"👤 本次日报写入人: "
                f"{writer_user_name} | {writer_user_guid or USER_GUID}"
            )

            insert_markdown_to_note(
                note_guid=note_guid,
                markdown_content=markdown_content,
                writer_user_guid=writer_user_guid
            )

            success_count += 1

        except Exception as e:
            fail_count += 1
            print(f"❌ 写入失败: note_guid={note_guid}, error={e}")
            print(traceback.format_exc())

    print("\n================ 写入统计 ================\n")
    print(f"成功写入日报数: {success_count}")
    print(f"无新增跳过日报数: {skip_count}")
    print(f"失败日报数: {fail_count}")
    print("\n==========================================\n")

    if fail_count > 0:
        raise Exception(f"存在 {fail_count} 个日报写入失败，请查看日志")


# =============================================================================
# 主流程
# =============================================================================


print("\n==========================================")
print("🚀 Todo 写入项目日报应用启动")
print("==========================================\n")

print(f"BASE_URL: {BASE_URL}")
print(f"TARGET_DATE: {TARGET_DATE}")
print(f"USER_GUID: {USER_GUID}")
print(f"TODO_PROJECT_ID: {TODO_PROJECT_ID}")
print(f"MAPPING_GUID: {MAPPING_GUID}")

mapping_rows = load_mapping_from_guid(MAPPING_GUID)

todo_list = fetch_todo_list()

if not todo_list:
    print("ℹ️ 当前未获取到 Todo 数据，流程结束")
    sys.exit(0)

note_task_map = build_note_task_map(
    todo_list=todo_list,
    mapping_rows=mapping_rows,
    target_date=TARGET_DATE
)

if not note_task_map:
    print("ℹ️ 当前没有可写入日报的 Todo，流程结束")
    sys.exit(0)

write_all_notes(note_task_map)

print("\n==========================================")
print("🏁 Todo 写入项目日报应用完成")
print("==========================================\n")