#!/usr/bin/env python3
"""
GSOE9011 weekly team form automation.

Behavior:
- group is an input parameter
- week is an input parameter
- members are expanded automatically from the chosen group
- Q4-Q8 are fixed to score 5
- Q9 is fixed to a predefined comment
- Q10 is fixed to No
"""

import argparse
import json
import os
import random
import re
import time
import uuid
from datetime import datetime, timedelta, timezone
from http.cookies import SimpleCookie
from pathlib import Path
from urllib.parse import urlparse

import requests


FORM_ID = "pM_2PxXn20i44Qhnufn7owEVDOPsLAJNqd7DEI0ZF-xURTBGRDhGMkgyUTgwVlc0NFkxRklRUFdGTi4u"
BASE_URL = "https://forms.office.com"
FORM_PAGE_URL = f"{BASE_URL}/pages/responsepage.aspx?id={FORM_ID}&route=shorturl"
STARTUP_URL = (
    f"{BASE_URL}/handlers/ResponsePageStartup.ashx"
    f"?id={FORM_ID}&route=shorturl&mobile=true"
)

SCRIPT_DIR = Path(__file__).resolve().parent
COOKIE_FILE = SCRIPT_DIR / "forms_cookie.txt"
FORM_PAGE_FILE = SCRIPT_DIR / "form_page.html"
FORM_PAGE_CONTEXT_FILE = SCRIPT_DIR / "form_page_context.json"
FORM_PAYLOAD_FILE = SCRIPT_DIR / "form_payload.json"
FORM_STARTUP_FILE = SCRIPT_DIR / "form_startup.json"
FORM_SCHEMA_FILE = SCRIPT_DIR / "form_schema.json"

WEEK_QUESTION_TITLE = "Week of Term"
GROUP_QUESTION_TITLE = "Group name"
COMMENT_QUESTION_TITLE = (
    "Comments on team member performance. For this team member (including yourself), "
    "please provide any additional comments and/or justification for your ratings above."
)
COMMENT_FIXED_TEXT = "Exceptional performance in all areas."
MANAGEMENT_QUESTION_TITLE = (
    "Management. Based on your rating of this team member’s performance, do you feel "
    "there are any issues that require intervention by teaching staff?"
)
RATING_QUESTION_TITLES = [
    "Contribution (equity and effort to the team’s work) / Contributing to the team’s work (dimension from paper)",
    "Collaboration (interaction with teammates) / Interacting with teammates (dimension from paper)",
    "Progression (keeping the team on track) / Keeping the team on track (dimension from paper)",
    "Quality (meeting expectations and standards) / Expecting quality (dimension from paper)",
    "Competencies (having related knowledge, skills, and abilities) / Having related knowledge, skills, and abilities (dimension from paper)",
]
REQUEST_HEADER_PROFILES = [
    {"Accept-Language": "en-AU,en;q=0.9", "Cache-Control": "no-cache"},
    {"Accept-Language": "en-US,en;q=0.9", "Cache-Control": "max-age=0"},
    {"Accept-Language": "en-US,en;q=0.8,zh-CN;q=0.6", "Pragma": "no-cache"},
]

SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": (
            "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/141.0.0.0 Mobile Safari/537.36"
        ),
        "sec-ch-ua-platform": '"Android"',
        "sec-ch-ua-mobile": "?1",
        "sec-ch-ua": '"Chromium";v="9", "Not?A_Brand";v="8"',
    }
)


def write_json(path: Path, data):
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def load_cookie_string():
    env_cookie = os.environ.get("FORMS_COOKIE", "").strip()
    if env_cookie:
        return normalize_cookie_string(env_cookie)

    cookie_file_path = os.environ.get("FORMS_COOKIE_FILE", "").strip()
    if cookie_file_path:
        path = Path(cookie_file_path).expanduser()
        if path.exists():
            return normalize_cookie_string(path.read_text(encoding="utf-8"))

    if COOKIE_FILE.exists():
        return normalize_cookie_string(COOKIE_FILE.read_text(encoding="utf-8"))

    return ""


def normalize_cookie_string(raw_text):
    """Normalize copied browser cookie headers into a raw cookie string."""
    text = (raw_text or "").strip()
    if not text:
        return ""

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return ""

    # Common copy format:
    # cookie
    # MUID=...; FormsWebSessionId=...
    if lines[0].lower() == "cookie":
        lines = lines[1:]

    text = " ".join(lines).strip()
    if text.lower().startswith("cookie:"):
        text = text.split(":", 1)[1].strip()

    return text


def apply_cookie_string(cookie_string):
    if not cookie_string:
        return

    jar = SimpleCookie()
    jar.load(cookie_string)
    for name, morsel in jar.items():
        SESSION.cookies.set(name, morsel.value, path="/")


def extract_page_context(html):
    match = re.search(
        r"window\.OfficeFormServerInfo\s*=\s*(\{.*?\});",
        html,
        flags=re.S,
    )
    if not match:
        raise ValueError(
            "未找到页面上下文 / Page context not found. 请检查 Cookie 是否有效 / "
            "Please check whether the cookie is still valid."
        )

    return json.loads(match.group(1))


def fetch_form_page():
    cookie_string = load_cookie_string()
    if not cookie_string:
        raise ValueError(
            "缺少 Cookie / Missing cookie. 请把完整浏览器 Cookie 放到 forms_cookie.txt "
            "或 FORMS_COOKIE。"
        )

    apply_cookie_string(cookie_string)
    headers = {
        "Referer": "https://login.microsoftonline.com/",
        "Upgrade-Insecure-Requests": "1",
    }
    resp = SESSION.get(FORM_PAGE_URL, headers=headers, timeout=30)
    resp.raise_for_status()

    FORM_PAGE_FILE.write_text(resp.text, encoding="utf-8")
    page_context = extract_page_context(resp.text)
    write_json(FORM_PAGE_CONTEXT_FILE, page_context)
    return page_context


def fetch_prefetch_structure(page_context):
    prefetch_url = page_context.get("prefetchFormUrl")
    if not prefetch_url:
        return None

    headers = {
        "Accept": "application/json",
        "X-UserSessionId": page_context.get("serverSessionId", ""),
        "__RequestVerificationToken": page_context.get("antiForgeryToken", ""),
        "Referer": FORM_PAGE_URL,
    }
    resp = SESSION.get(prefetch_url, headers=headers, timeout=30)
    resp.raise_for_status()

    data = resp.json()
    write_json(FORM_PAYLOAD_FILE, data)
    return data


def fetch_startup_structure():
    correlation_id = str(uuid.uuid4())
    headers = {
        "x-fsw-page": "/pages/responsepage.aspx",
        "x-fsw-ring": "Business",
        "x-fsw-enable": "1",
        "x-fsw-startup": "1",
        "x-fsw-baseclient": "formweekly_cd_20260323.2",
        "x-fsw-client": "formweekly_cd_20260323.2",
        "x-fsw-cdn": f"{BASE_URL}/cdn",
        "x-fsw-server": "16.0.19919.42051",
        "X-CorrelationId": correlation_id,
        "Content-Type": "application/json",
        "Referer": FORM_PAGE_URL,
    }
    resp = SESSION.get(STARTUP_URL, headers=headers, timeout=30)
    resp.raise_for_status()
    data = resp.json()
    write_json(FORM_STARTUP_FILE, data)
    return data


def get_form_structure():
    page_context = fetch_form_page()
    try:
        structure_data = fetch_prefetch_structure(page_context)
        if structure_data:
            return page_context, structure_data, "prefetch"
    except Exception as exc:
        print(f"[!] prefetch 拉取失败，回退 startup / Prefetch failed, falling back to startup: {exc}")

    return page_context, fetch_startup_structure(), "startup"


def find_form_info(node):
    if isinstance(node, dict):
        if isinstance(node.get("questions"), list) or isinstance(node.get("questionInfo"), list):
            return node

        for key in ("form", "formInfo", "requestForm", "data", "value"):
            if key in node:
                found = find_form_info(node[key])
                if found is not None:
                    return found

        for value in node.values():
            found = find_form_info(value)
            if found is not None:
                return found

    elif isinstance(node, list):
        for item in node:
            found = find_form_info(item)
            if found is not None:
                return found

    return None


def parse_questions(structure_data):
    form_info = find_form_info(structure_data)
    if form_info is None:
        if isinstance(structure_data, dict):
            error_info = (structure_data.get("data") or {}).get("error") or {}
            error_message = error_info.get("message", "").strip()
            if error_message:
                raise ValueError(
                    f"无法加载题目 / Could not load form questions: {error_message}. "
                    "请刷新 forms_cookie.txt / Please refresh forms_cookie.txt."
                )
        raise ValueError("Could not locate questions in the response payload.")

    questions = form_info.get("questions") or form_info.get("questionInfo") or []
    questions = sorted(questions, key=lambda q: q.get("order", 0))
    meta = {
        "formId": form_info.get("id", form_info.get("formId", "")),
        "tenantId": form_info.get("tenantId", ""),
        "ownerId": form_info.get("ownerId", ""),
        "title": form_info.get("title", form_info.get("name", "")),
    }
    return meta, questions


def enrich_meta_from_page_context(meta, page_context):
    meta = dict(meta)
    prefetch_url = page_context.get("prefetchFormUrl", "")
    if not prefetch_url:
        return meta

    path_parts = [part for part in urlparse(prefetch_url).path.split("/") if part]
    try:
        api_index = path_parts.index("api")
        users_index = path_parts.index("users")
        tenant_id = path_parts[api_index + 1]
        owner_id = path_parts[users_index + 1]
    except (ValueError, IndexError):
        return meta

    if tenant_id and not meta.get("tenantId"):
        meta["tenantId"] = tenant_id
    if owner_id and not meta.get("ownerId"):
        meta["ownerId"] = owner_id
    return meta


def parse_question_info(question):
    raw = question.get("deserializedQuestionInfo") or question.get("questionInfo")
    if isinstance(raw, dict):
        return raw
    if isinstance(raw, str):
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {}
    return {}


def get_question_choices(question):
    direct_choices = question.get("choices") or question.get("options") or []
    if direct_choices:
        return direct_choices

    question_info = parse_question_info(question)
    embedded_choices = question_info.get("Choices") or question_info.get("choices") or []
    normalized = []
    for choice in embedded_choices:
        branch_info = choice.get("BranchInfo") or choice.get("branchInfo") or {}
        normalized.append(
            {
                "id": choice.get("Id", choice.get("id", "")),
                "description": choice.get("Description", choice.get("description", "")),
                "value": choice.get("Value", choice.get("value", "")),
                "text": choice.get("Text", choice.get("text", "")),
                "target_question_id": branch_info.get("TargetQuestionId")
                or branch_info.get("targetQuestionId", ""),
            }
        )
    return normalized


def find_question_by_title(questions, title):
    for question in questions:
        if question.get("title") == title:
            return question
    raise ValueError(f"Question not found: {title}")


def find_question_by_id(questions, question_id):
    for question in questions:
        qid = question.get("id", question.get("questionId", ""))
        if qid == question_id:
            return question
    raise ValueError(f"Question ID not found: {question_id}")


def build_question_schema(meta, questions):
    payload = {"meta": meta, "questions": []}
    for index, question in enumerate(questions, 1):
        payload["questions"].append(
            {
                "index": index,
                "id": question.get("id", question.get("questionId", "")),
                "type": question.get("type", ""),
                "required": question.get("required", question.get("isRequired", False)),
                "title": question.get("title", question.get("questionText", "")),
                "options": [
                    {
                        "id": choice.get("id", ""),
                        "text": choice.get("description")
                        or choice.get("value")
                        or choice.get("text", ""),
                    }
                    for choice in get_question_choices(question)
                ],
            }
        )
    return payload


def format_group_label(group_input):
    raw = str(group_input).strip()
    if raw.startswith("Group "):
        number_text = raw.split(" ", 1)[1].strip()
    else:
        number_text = raw

    if not re.fullmatch(r"\d{1,3}", number_text):
        raise ValueError("Group must be 1-154 or 01-154.")

    number = int(number_text)
    if number < 1 or number > 154:
        raise ValueError("Group must be in the range 01-154.")

    return f"Group {number:02d}"


def resolve_choice_text(question, desired_value):
    target = str(desired_value).strip()
    choices = get_question_choices(question)
    if not choices:
        raise ValueError(f"No options found for question: {question.get('title')}")

    for choice in choices:
        choice_text = str(choice.get("description") or choice.get("value") or choice.get("text", "")).strip()
        if choice_text == target:
            return choice_text

    lowered = target.lower()
    for choice in choices:
        choice_text = str(choice.get("description") or choice.get("value") or choice.get("text", "")).strip()
        if choice_text.lower() == lowered:
            return choice_text

    for choice in choices:
        choice_text = str(choice.get("description") or choice.get("value") or choice.get("text", "")).strip()
        if choice_text.startswith(target):
            return choice_text

    raise ValueError(f"Option not found in [{question.get('title')}]: {desired_value}")


def parse_week_values(week_input, week_question):
    raw = str(week_input).strip()
    if not raw:
        raise ValueError("Week input is required.")

    tokens = [token.strip() for token in raw.split(",") if token.strip()]
    if not tokens:
        raise ValueError("Week input is required.")

    normalized = []
    seen = set()
    for token in tokens:
        if token.lower().startswith("week "):
            suffix = token.split(" ", 1)[1].strip()
        else:
            suffix = token

        if not re.fullmatch(r"\d{1,2}", suffix):
            raise ValueError("Week format is invalid. Use 9 or Week 9 or 9,10.")

        label = f"Week {int(suffix)}"
        label = resolve_choice_text(week_question, label)
        if label not in seen:
            normalized.append(label)
            seen.add(label)

    return normalized


def find_group_member_question(questions, group_question, group_label):
    for choice in get_question_choices(group_question):
        choice_text = str(choice.get("description") or choice.get("value") or choice.get("text", "")).strip()
        if choice_text != group_label:
            continue

        target_question_id = choice.get("target_question_id", "")
        if target_question_id:
            return find_question_by_id(questions, target_question_id)

    return find_question_by_title(questions, f"Group Member Name ({group_label})")


def normalize_choice_answer(question, answer_value):
    answer_text = str(answer_value).strip()
    choices = get_question_choices(question)
    if not choices:
        return answer_text

    for choice in choices:
        choice_text = str(choice.get("description") or choice.get("value") or choice.get("text", "")).strip()
        choice_id = str(choice.get("id", "")).strip()
        if answer_text == choice_text or answer_text == choice_id:
            return choice_text or answer_text
    return answer_text


def build_answer_payload(questions, answers):
    answer_list = []
    for question in questions:
        qid = question.get("id", question.get("questionId", ""))
        if qid not in answers:
            continue

        answer_value = answers[qid]
        question_type = question.get("type", "")
        if "choice" in question_type.lower():
            answer_value = normalize_choice_answer(question, answer_value)
        else:
            answer_value = str(answer_value)

        answer_list.append({"questionId": qid, "answer1": answer_value})

    return json.dumps(answer_list, ensure_ascii=False)


def build_batch_submissions(questions, group_input, week_input):
    week_question = find_question_by_title(questions, WEEK_QUESTION_TITLE)
    group_question = find_question_by_title(questions, GROUP_QUESTION_TITLE)
    management_question = find_question_by_title(questions, MANAGEMENT_QUESTION_TITLE)
    comment_question = find_question_by_title(questions, COMMENT_QUESTION_TITLE)

    group_label = resolve_choice_text(group_question, format_group_label(group_input))
    week_values = parse_week_values(week_input, week_question)
    member_question = find_group_member_question(questions, group_question, group_label)
    member_values = [
        str(choice.get("description") or choice.get("value") or choice.get("text", "")).strip()
        for choice in get_question_choices(member_question)
    ]
    if not member_values:
        raise ValueError(f"No members found for {group_label}.")

    rating_answers = {}
    for title in RATING_QUESTION_TITLES:
        question = find_question_by_title(questions, title)
        qid = question.get("id", question.get("questionId", ""))
        rating_answers[qid] = resolve_choice_text(question, "5")

    management_qid = management_question.get("id", management_question.get("questionId", ""))
    management_answer = resolve_choice_text(management_question, "No")
    comment_qid = comment_question.get("id", comment_question.get("questionId", ""))
    week_qid = week_question.get("id", week_question.get("questionId", ""))
    group_qid = group_question.get("id", group_question.get("questionId", ""))
    member_qid = member_question.get("id", member_question.get("questionId", ""))

    submissions = []
    for week in week_values:
        for member in member_values:
            answers = {
                week_qid: week,
                group_qid: group_label,
                member_qid: member,
                comment_qid: COMMENT_FIXED_TEXT,
                management_qid: management_answer,
            }
            answers.update(rating_answers)
            submissions.append(
                {
                    "week": week,
                    "group": group_label,
                    "member": member,
                    "answers": answers,
                }
            )

    return {
        "group": group_label,
        "weeks": week_values,
        "members": member_values,
        "member_question": member_question.get("title"),
        "submissions": submissions,
    }


def build_submit_headers(page_context):
    headers = {
        "Content-Type": "application/json",
        "X-CorrelationId": str(uuid.uuid4()),
        "Referer": FORM_PAGE_URL,
    }
    anti_forgery_token = page_context.get("antiForgeryToken", "")
    user_session_id = page_context.get("serverSessionId", "")
    if anti_forgery_token:
        headers["__RequestVerificationToken"] = anti_forgery_token
    if user_session_id:
        headers["X-UserSessionId"] = user_session_id
    headers.update(random.choice(REQUEST_HEADER_PROFILES))
    return headers


def submit_form(meta, answers_json, page_context):
    now = datetime.now(timezone.utc)
    payload = {
        "startDate": (now - timedelta(seconds=5)).strftime("%Y-%m-%dT%H:%M:%S.000Z"),
        "submitDate": now.strftime("%Y-%m-%dT%H:%M:%S.000Z"),
        "answers": answers_json,
    }

    form_id = meta["formId"]
    tenant_id = meta.get("tenantId", "")
    owner_id = meta.get("ownerId", "")
    if tenant_id and owner_id:
        submit_url = (
            f"{BASE_URL}/formapi/api/{tenant_id}/users/{owner_id}"
            f"/forms('{form_id}')/responses"
        )
    else:
        submit_url = f"{BASE_URL}/formapi/api/{form_id}/responses"

    headers = build_submit_headers(page_context)
    resp = SESSION.post(submit_url, json=payload, headers=headers, timeout=30)
    return resp


def parse_args():
    parser = argparse.ArgumentParser(description="GSOE9011 weekly team form automation")
    parser.add_argument("--group", help="Group number, for example 60 or 060 or Group 60")
    parser.add_argument("--week", help="Week input, for example 9 or Week 9 or 9,10")
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Only print the plan and generated payloads, do not submit.",
    )
    parser.add_argument(
        "--export-only",
        action="store_true",
        help="Only fetch and export form structure, do not build submissions.",
    )
    parser.add_argument("--delay-min", type=float, default=1.5)
    parser.add_argument("--delay-max", type=float, default=4.0)
    return parser.parse_args()


def prompt_if_missing(value, prompt_text):
    if value:
        return value
    return input(prompt_text).strip()


def collect_runtime_inputs(args):
    """When no CLI args are provided, ask the user interactively."""
    if not args.week and not args.group:
        print("\n[手动输入模式 / Interactive Input Mode]")
        print("请先输入周次，再输入组号。后续成员会自动循环，固定题会自动填写。")
        print("Enter week first, then group. Members will be looped automatically and fixed questions will be filled automatically.")

    week_input = prompt_if_missing(
        args.week,
        "输入周次 / Enter week (例如 / for example 9 or 9,10): ",
    )
    group_input = prompt_if_missing(
        args.group,
        "输入组号 / Enter group (01-154): ",
    )
    return week_input, group_input


def run_batch(meta, questions, page_context, group_input, week_input, dry_run, delay_min, delay_max):
    batch = build_batch_submissions(questions, group_input, week_input)
    submissions = batch["submissions"]
    total = len(submissions)

    print("\n[批量计划]")
    print(f"组别 / Group: {batch['group']}")
    print(f"周次 / Weeks: {', '.join(batch['weeks'])}")
    print(f"成员题 / Member Question: {batch['member_question']}")
    print(f"成员数量 / Member Count: {len(batch['members'])}")
    print(f"预计提交数 / Planned Submissions: {total}")
    print(
        "固定规则 / Fixed Rules: Q4-Q8 选 5 / score 5, "
        f"Q9 固定文本 / fixed comment = {COMMENT_FIXED_TEXT!r}, Q10 选 No"
    )
    print("成员列表 / Member List:")
    for index, member in enumerate(batch["members"], 1):
        print(f"  {index}. {member}")
    print("固定答案明细 / Fixed Answer Details:")
    print("  - Q4-Q8: 5")
    print(f"  - Q9: {COMMENT_FIXED_TEXT}")
    print("  - Q10: No")

    for index, item in enumerate(submissions, 1):
        print(f"  [{index}/{total}] {item['week']} -> {item['member']}")

    if dry_run:
        print("\n[Dry Run] 未实际提交 / No real submission was sent.")
        if submissions:
            first_payload = json.loads(build_answer_payload(questions, submissions[0]["answers"]))
            print("\n[首条 Payload 预览 / First Payload Preview]")
            print(json.dumps(first_payload, indent=2, ensure_ascii=False))
        return

    confirmation = input(
        "\n输入 yes / Y / y 后开始提交，其它任意输入取消 / "
        "Type yes / Y / y to start submission, anything else to cancel: "
    ).strip()
    if confirmation.lower() != "yes" and confirmation not in {"Y", "y"}:
        print("已取消，未提交任何内容 / Cancelled, no submission was sent.")
        return

    if delay_max < delay_min:
        raise ValueError("delay-max must be greater than or equal to delay-min.")

    for index, item in enumerate(submissions, 1):
        print(f"\n[提交 {index}/{total} / Submit {index}/{total}] {item['week']} / {item['member']}")
        answers_json = build_answer_payload(questions, item["answers"])
        resp = submit_form(meta, answers_json, page_context)
        print(f"HTTP {resp.status_code}")
        if resp.status_code not in (200, 201, 204):
            try:
                print(json.dumps(resp.json(), indent=2, ensure_ascii=False))
            except Exception:
                print(resp.text[:2000])
            raise RuntimeError(
                f"Submit failed: HTTP {resp.status_code} ({item['week']} / {item['member']})"
            )

        if index < total and delay_max > 0:
            sleep_seconds = random.uniform(delay_min, delay_max)
            print(f"等待 {sleep_seconds:.2f} 秒后继续 / Waiting {sleep_seconds:.2f}s before next request ...")
            time.sleep(sleep_seconds)


def main():
    args = parse_args()
    try:
        page_context, structure_data, source = get_form_structure()
        meta, questions = parse_questions(structure_data)
    except Exception as exc:
        print(f"[!] 获取表单结构失败 / Failed to load form structure: {exc}")
        return
    meta = enrich_meta_from_page_context(meta, page_context)
    write_json(FORM_SCHEMA_FILE, build_question_schema(meta, questions))

    print(f"表单标题 / Form Title: {meta.get('title', 'N/A')}")
    print(f"题目数量 / Question Count: {len(questions)}")
    print(f"结构来源 / Structure Source: {source}")
    print(f"结构文件 / Structure File: {FORM_SCHEMA_FILE}")

    if args.export_only:
        print("仅导出结构，未生成提交计划 / Structure exported only, no submission plan generated.")
        return

    week_input, group_input = collect_runtime_inputs(args)

    run_batch(
        meta=meta,
        questions=questions,
        page_context=page_context,
        group_input=group_input,
        week_input=week_input,
        dry_run=args.dry_run,
        delay_min=args.delay_min,
        delay_max=args.delay_max,
    )


if __name__ == "__main__":
    main()
