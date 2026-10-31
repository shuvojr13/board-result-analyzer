"""
analysis.py - Comprehensive Result Analysis for Bangladesh Board Examinations
Calculates:
- Regular vs Irregular Candidate Performance Analysis & Comparison
- Subject Wise Grade Analysis (A+, A, A-, B, C, D, F, Appeared, Passed, Pass Rate)
- 1-Subject Fail Breakdown and Student Roster
- 2-Subject Fail Breakdown and Student Roster
- 3+ Subject Fail List
- GPA & Group Distributions
- Excel Analysis Workbook Generation
"""

import io
import re
from collections import defaultdict
from typing import Any, Dict, List, Optional
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


VALID_GRADES = ["A+", "A", "A-", "B", "C", "D", "F"]

GRADE_POINTS = {
    "A+": 5.0,
    "A": 4.0,
    "A-": 3.5,
    "B": 3.0,
    "C": 2.0,
    "D": 1.0,
    "F": 0.0
}


def identify_fourth_subject(subject_grades: Dict[str, str], group: str) -> Optional[str]:
    """
    Identifies the 4th (optional) subject from a student's enrolled subject list.
    In Bangladesh curriculum, regular students have 3 compulsory subjects (Bangla, English, ICT),
    3 group-core electives, and 1 fourth/optional subject (7 subjects total in HSC/SSC).
    """
    if not subject_grades:
        return None

    # Check if any subject explicitly contains 4th/fourth/optional in name
    for s in subject_grades.keys():
        s_low = s.lower()
        if any(term in s_low for term in ["4th", "fourth", "optional", "additional"]):
            return s

    # Typical 4th subjects in Bangladesh education boards
    typical_fourth_keywords = [
        "agriculture", "higher math", "biology", "statistics", "psychology",
        "home science", "social work", "finance", "logic", "geography",
        "production management", "islamic studies", "music", "arts"
    ]

    subjs = list(subject_grades.keys())

    # If student has 7 or more subjects, the 4th subject is placed last in board tables
    if len(subjs) >= 7:
        last_subj = subjs[-1]
        last_subj_low = last_subj.lower()
        # If last subject is not Bangla, English, or ICT:
        if not any(comp in last_subj_low for comp in ["bangla", "english", "information & communication"]):
            return last_subj

    # If typical keyword found in any subject (starting check from reverse):
    for s in reversed(subjs):
        s_low = s.lower()
        if any(kw in s_low for kw in typical_fourth_keywords):
            return s

    # If 7 or more subjects and no other match, return the last subject
    if len(subjs) >= 7:
        return subjs[-1]

    return None


def normalize_record(raw: Dict[str, Any]) -> Dict[str, Any]:
    """
    Normalizes a student record from scraper dictionary or DataFrame row.
    Extracts Roll, Reg, Name, Group, Type (Regular/Irregular), Result, GPA (With & Without 4th),
    Subject grades, 4th subject, and Absent status.
    """
    roll = str(raw.get("Roll", "")).strip().split(".")[0]
    reg = str(raw.get("Registration", "") or raw.get("Reg", "")).strip().split(".")[0]
    if reg == "nan":
        reg = ""
    name = str(raw.get("Name", "N/A")).strip()
    if name == "nan":
        name = "N/A"

    father_name = str(raw.get("Father Name", "") or raw.get("Father's Name", "")).strip()
    if father_name == "nan":
        father_name = ""

    institute = str(raw.get("Institute", "") or raw.get("Institution", "")).strip()
    if institute == "nan":
        institute = ""

    group = str(raw.get("Group", "") or "GENERAL").strip().upper()
    if not group or group in ["NAN", "NONE", "-"]:
        group = "GENERAL"

    result_status = str(raw.get("Result Status", "") or raw.get("Result", "") or "").strip().upper()
    status_field = str(raw.get("Status", "")).strip().upper()

    gpa_raw = str(raw.get("GPA", "0.00")).strip()
    try:
        gpa = float(gpa_raw)
    except Exception:
        gpa = 0.00

    subject_grades: Dict[str, str] = {}

    # 1. Check if Subject_Grades dict is directly present
    if isinstance(raw.get("Subject_Grades"), dict) and raw["Subject_Grades"]:
        for k, v in raw["Subject_Grades"].items():
            g = str(v).strip().upper()
            if g in VALID_GRADES or g in ["ABS", "ABSENT"]:
                subject_grades[str(k).strip()] = g

    # 2. Check if Subject Grades Summary string is present (e.g. "101 BANGLA: F, 107 ENGLISH: B")
    summary_str = str(raw.get("Subject Grades Summary", "") or raw.get("Subject_Grades_Summary", "")).strip()
    if not subject_grades and summary_str and summary_str != "nan":
        parts = summary_str.split(",")
        for part in parts:
            if ":" in part:
                s_name, s_grade = part.split(":", 1)
                s_grade = s_grade.strip().upper()
                if s_grade in VALID_GRADES or s_grade in ["ABS", "ABSENT"]:
                    subject_grades[s_name.strip()] = s_grade

    # 3. Check individual column keys
    if not subject_grades:
        exclude_keys = {
            "roll", "registration", "reg", "name", "father name", "mother name",
            "institute", "board", "group", "exam", "year", "result status",
            "result", "gpa", "status", "subject grades summary", "subject_grades_summary",
            "type", "session", "created_at"
        }
        for k, v in raw.items():
            if str(k).lower().strip() not in exclude_keys:
                grade_val = str(v).strip().upper()
                if grade_val in VALID_GRADES or grade_val in ["ABS", "ABSENT"]:
                    subject_grades[str(k).strip()] = grade_val

    # Detect candidate type (Regular / Irregular / Improvement / Private)
    cand_type_raw = str(
        raw.get("Type") or raw.get("Candidate Type") or raw.get("Student Type") or ""
    ).strip().upper()

    if not cand_type_raw or cand_type_raw in ["NAN", "NONE", "-"]:
        if 0 < len(subject_grades) <= 4:
            cand_type = "IRREGULAR"
        else:
            cand_type = "REGULAR"
    elif "IRREG" in cand_type_raw:
        cand_type = "IRREGULAR"
    elif "REG" in cand_type_raw:
        cand_type = "REGULAR"
    elif "IMP" in cand_type_raw:
        cand_type = "IMPROVEMENT"
    elif "PRIV" in cand_type_raw:
        cand_type = "PRIVATE"
    else:
        cand_type = cand_type_raw

    # Identify 4th Subject
    fourth_subj = identify_fourth_subject(subject_grades, group)
    fourth_grade = subject_grades.get(fourth_subj) if fourth_subj else None
    fourth_gp = GRADE_POINTS.get(fourth_grade, 0.0) if fourth_grade else 0.0

    # Separate compulsory/core subjects from 4th subject
    main_subjects = {k: v for k, v in subject_grades.items() if k != fourth_subj}

    # Detect absent status
    is_absent = (
        ("ABSENT" in result_status) or
        (result_status in ["ABS", "A"]) or
        ("ABSENT" in status_field) or
        ("ABSENT" in summary_str.upper()) or
        (len(subject_grades) > 0 and all(g in ["ABS", "ABSENT", "-"] for g in subject_grades.values()))
    )

    # Detect failed subjects (only grades == 'F')
    failed_subjects = [s for s, g in subject_grades.items() if g == "F"]
    passed_subjects = [f"{s} ({g})" for s, g in subject_grades.items() if g != "F" and g not in ["ABS", "ABSENT"]]

    # In Bangladesh board rules: Failing 4th subject does NOT fail the student if all main subjects are passed!
    main_failed_subjects = [s for s in failed_subjects if s != fourth_subj]
    is_fail = False
    is_passed = False

    if is_absent:
        final_status = "ABSENT"
    else:
        if ("FAIL" in result_status) or (len(main_failed_subjects) > 0):
            is_fail = True
            final_status = "FAIL"
        elif gpa > 0.0 or result_status == "PASSED":
            is_passed = True
            final_status = "PASSED"
        else:
            final_status = result_status if result_status else "UNKNOWN"

    # Calculate GPA With 4th Subject and GPA Without 4th Subject
    if is_passed:
        # GPA Without 4th Subject: strictly on the compulsory & core subjects (Golden A+ = >= 5.00 without 4th)
        if len(main_subjects) > 0:
            sum_main_gp = sum(GRADE_POINTS.get(g, 0.0) for g in main_subjects.values())
            gpa_without_4th = min(5.00, round(sum_main_gp / len(main_subjects), 2))
        else:
            gpa_without_4th = gpa

        # GPA With 4th Subject: official board GPA or computed with 4th subject bonus
        bonus = max(0.0, fourth_gp - 2.0) if fourth_subj else 0.0
        if len(main_subjects) > 0:
            sum_main_gp = sum(GRADE_POINTS.get(g, 0.0) for g in main_subjects.values())
            computed_w4 = min(5.00, round((sum_main_gp + bonus) / len(main_subjects), 2))
            gpa_with_4th = max(gpa, computed_w4) if gpa > 0.0 else computed_w4
        else:
            gpa_with_4th = gpa
    else:
        gpa_with_4th = 0.00
        gpa_without_4th = 0.00

    is_gpa5_with_4th = is_passed and gpa_with_4th >= 5.00
    is_gpa5_without_4th = is_passed and gpa_without_4th >= 5.00

    return {
        "Roll": roll,
        "Registration": reg,
        "Name": name,
        "Father Name": father_name,
        "Institute": institute,
        "Group": group,
        "Type": cand_type,
        "Result": final_status,
        "GPA": f"{gpa_with_4th:.2f}",
        "GPANumeric": gpa_with_4th,
        "GPA_With_4th": f"{gpa_with_4th:.2f}",
        "GPA_Without_4th": f"{gpa_without_4th:.2f}",
        "GPA_With_4th_Numeric": gpa_with_4th,
        "GPA_Without_4th_Numeric": gpa_without_4th,
        "Fourth_Subject": fourth_subj,
        "Fourth_Subject_Grade": fourth_grade,
        "Is_GPA5_With_4th": is_gpa5_with_4th,
        "Is_GPA5_Without_4th": is_gpa5_without_4th,
        "Is_Absent": is_absent,
        "Subject_Grades": subject_grades,
        "Failed_Subjects": failed_subjects,
        "Failed_Count": len(failed_subjects),
        "Passed_Grades_Summary": ", ".join(passed_subjects) if passed_subjects else "None",
        "Raw": raw
    }


def compute_result_analysis(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Computes all analytical aggregations across the given list of student records,
    including Regular vs Irregular candidate analyses.
    """
    normalized = [normalize_record(r) for r in records if r.get("Roll") or r.get("roll")]
    total_students = len(normalized)

    if total_students == 0:
        return {
            "total_students": 0,
            "summary": {},
            "groups_summary": {},
            "types_summary": {},
            "regular_summary": {},
            "irregular_summary": {},
            "group_type_comparison": [],
            "subject_analysis": [],
            "one_subject_fail": {"total": 0, "subjects_summary": {}, "students": []},
            "two_subject_fail": {"total": 0, "subjects_summary": {}, "pair_summary": {}, "students": []},
            "multi_subject_fail": {"total": 0, "students": []},
            "gpa_distribution": {},
            "available_groups": [],
            "available_types": []
        }

    # Tracking containers
    passed_students = []
    failed_students = []
    absent_students = []
    other_students = []

    one_subject_fail_students = []
    two_subject_fail_students = []
    multi_subject_fail_students = []

    # Frequency maps
    one_sub_freq = defaultdict(int)
    two_sub_freq = defaultdict(int)
    two_sub_pairs = defaultdict(int)

    # GPA Distributions (Both With 4th and Without 4th subject)
    gpa_dist_with_4th = {
        "5.00 (A+)": 0,
        "4.00 - 4.99 (A)": 0,
        "3.50 - 3.99 (A-)": 0,
        "3.00 - 3.49 (B)": 0,
        "2.00 - 2.99 (C)": 0,
        "1.00 - 1.99 (D)": 0,
        "0.00 (F / Fail)": 0
    }
    gpa_dist_without_4th = {
        "5.00 (A+)": 0,
        "4.00 - 4.99 (A)": 0,
        "3.50 - 3.99 (A-)": 0,
        "3.00 - 3.49 (B)": 0,
        "2.00 - 2.99 (C)": 0,
        "1.00 - 1.99 (D)": 0,
        "0.00 (F / Fail)": 0
    }

    # Group tracking: group_name -> list of students
    group_students = defaultdict(list)

    # Type tracking: type_name -> list of students
    type_students = defaultdict(list)

    # Subject tracking per group & candidate type:
    # (group, type, subject) -> {"A+": 0, "A": 0, ...}
    subject_stats = defaultdict(lambda: {g: 0 for g in VALID_GRADES})

    for st in normalized:
        grp = st["Group"]
        c_type = st["Type"]
        group_students[grp].append(st)
        type_students[c_type].append(st)

        gpa_w4 = st["GPA_With_4th_Numeric"]
        gpa_no4 = st["GPA_Without_4th_Numeric"]
        res = st["Result"]
        fails = st["Failed_Subjects"]
        fail_cnt = len(fails)
        is_p = (res == "PASSED")

        # Result grouping
        if res == "PASSED":
            passed_students.append(st)
        elif res == "FAIL":
            failed_students.append(st)
        elif res == "ABSENT":
            absent_students.append({
                "Roll": st["Roll"],
                "Registration": st["Registration"],
                "Name": st["Name"],
                "Father Name": st.get("Father Name", ""),
                "Institute": st.get("Institute", ""),
                "Group": grp,
                "Type": c_type,
                "Result": "ABSENT",
                "Status": "Absent in Examination",
                "Remarks": "Did not appear"
            })
        else:
            other_students.append(st)

        # Subject-fail categorization (Only for students who appeared and failed)
        if res != "ABSENT":
            if fail_cnt == 1:
                one_subject_fail_students.append({
                    "Roll": st["Roll"],
                    "Registration": st["Registration"],
                    "Name": st["Name"],
                    "Group": grp,
                    "Type": c_type,
                    "Failed_Subject": fails[0],
                    "Passed_Grades_Summary": st["Passed_Grades_Summary"],
                    "GPA": st["GPA_With_4th"],
                    "GPA_With_4th": st["GPA_With_4th"],
                    "GPA_Without_4th": st["GPA_Without_4th"]
                })
                one_sub_freq[fails[0]] += 1
            elif fail_cnt == 2:
                two_subject_fail_students.append({
                    "Roll": st["Roll"],
                    "Registration": st["Registration"],
                    "Name": st["Name"],
                    "Group": grp,
                    "Type": c_type,
                    "Failed_Subjects": fails,
                    "Passed_Grades_Summary": st["Passed_Grades_Summary"],
                    "GPA": st["GPA_With_4th"],
                    "GPA_With_4th": st["GPA_With_4th"],
                    "GPA_Without_4th": st["GPA_Without_4th"]
                })
                for s in fails:
                    two_sub_freq[s] += 1
                pair_key = " & ".join(sorted(fails))
                two_sub_pairs[pair_key] += 1
            elif fail_cnt >= 3:
                multi_subject_fail_students.append({
                    "Roll": st["Roll"],
                    "Registration": st["Registration"],
                    "Name": st["Name"],
                    "Group": grp,
                    "Type": c_type,
                    "Failed_Count": fail_cnt,
                    "Failed_Subjects": fails,
                    "Passed_Grades_Summary": st["Passed_Grades_Summary"],
                    "GPA": st["GPA_With_4th"],
                    "GPA_With_4th": st["GPA_With_4th"],
                    "GPA_Without_4th": st["GPA_Without_4th"]
                })

        # GPA distribution: With 4th Subject
        if res == "FAIL" or fail_cnt > 0 or not is_p:
            gpa_dist_with_4th["0.00 (F / Fail)"] += 1
        elif gpa_w4 >= 5.00:
            gpa_dist_with_4th["5.00 (A+)"] += 1
        elif gpa_w4 >= 4.00:
            gpa_dist_with_4th["4.00 - 4.99 (A)"] += 1
        elif gpa_w4 >= 3.50:
            gpa_dist_with_4th["3.50 - 3.99 (A-)"] += 1
        elif gpa_w4 >= 3.00:
            gpa_dist_with_4th["3.00 - 3.49 (B)"] += 1
        elif gpa_w4 >= 2.00:
            gpa_dist_with_4th["2.00 - 2.99 (C)"] += 1
        elif gpa_w4 >= 1.00:
            gpa_dist_with_4th["1.00 - 1.99 (D)"] += 1
        else:
            gpa_dist_with_4th["0.00 (F / Fail)"] += 1

        # GPA distribution: Without 4th Subject (Pure / Golden GPA)
        if res == "FAIL" or fail_cnt > 0 or not is_p:
            gpa_dist_without_4th["0.00 (F / Fail)"] += 1
        elif gpa_no4 >= 5.00:
            gpa_dist_without_4th["5.00 (A+)"] += 1
        elif gpa_no4 >= 4.00:
            gpa_dist_without_4th["4.00 - 4.99 (A)"] += 1
        elif gpa_no4 >= 3.50:
            gpa_dist_without_4th["3.50 - 3.99 (A-)"] += 1
        elif gpa_no4 >= 3.00:
            gpa_dist_without_4th["3.00 - 3.49 (B)"] += 1
        elif gpa_no4 >= 2.00:
            gpa_dist_without_4th["2.00 - 2.99 (C)"] += 1
        elif gpa_no4 >= 1.00:
            gpa_dist_without_4th["1.00 - 1.99 (D)"] += 1
        else:
            gpa_dist_without_4th["0.00 (F / Fail)"] += 1

        # Subject-wise tracking across group and candidate type
        if res != "ABSENT":
            for subj, gr in st["Subject_Grades"].items():
                if gr in VALID_GRADES:
                    subject_stats[(grp, c_type, subj)][gr] += 1
                    subject_stats[("ALL", c_type, subj)][gr] += 1
                    subject_stats[(grp, "ALL", subj)][gr] += 1
                    subject_stats[("ALL", "ALL", subj)][gr] += 1

    # Overall Summary calculation
    total_passed = len(passed_students)
    total_failed = len(failed_students)
    total_absent = len(absent_students)
    appeared = total_students - total_absent
    pass_rate = round((total_passed / appeared * 100), 2) if appeared > 0 else 0.0
    fail_rate = round((total_failed / appeared * 100), 2) if appeared > 0 else 0.0
    absent_rate = round((total_absent / total_students * 100), 2) if total_students > 0 else 0.0

    # GPA 5.0 With 4th Subject vs Without 4th Subject
    gpa_5_count_with_4th = sum(1 for s in passed_students if s["Is_GPA5_With_4th"])
    gpa_5_rate_with_4th = round((gpa_5_count_with_4th / appeared * 100), 2) if appeared > 0 else 0.0

    gpa_5_count_without_4th = sum(1 for s in passed_students if s["Is_GPA5_Without_4th"])
    gpa_5_rate_without_4th = round((gpa_5_count_without_4th / appeared * 100), 2) if appeared > 0 else 0.0

    # Group Summary calculation
    groups_summary = {}
    for grp, s_list in group_students.items():
        grp_tot = len(s_list)
        grp_pass = sum(1 for s in s_list if s["Result"] == "PASSED")
        grp_fail = sum(1 for s in s_list if s["Result"] == "FAIL")
        grp_absent = sum(1 for s in s_list if s["Result"] == "ABSENT")
        grp_app = grp_tot - grp_absent
        grp_gpa5_w4 = sum(1 for s in s_list if s["Is_GPA5_With_4th"])
        grp_gpa5_no4 = sum(1 for s in s_list if s["Is_GPA5_Without_4th"])
        grp_1fail = sum(1 for s in s_list if s["Failed_Count"] == 1 and s["Result"] != "ABSENT")
        grp_2fail = sum(1 for s in s_list if s["Failed_Count"] == 2 and s["Result"] != "ABSENT")
        grp_3fail = sum(1 for s in s_list if s["Failed_Count"] >= 3 and s["Result"] != "ABSENT")
        grp_rate = round((grp_pass / grp_app * 100), 2) if grp_app > 0 else 0.0
        grp_absent_rate = round((grp_absent / grp_tot * 100), 2) if grp_tot > 0 else 0.0

        groups_summary[grp] = {
            "group": grp,
            "total": grp_tot,
            "appeared": grp_app,
            "passed": grp_pass,
            "failed": grp_fail,
            "absent": grp_absent,
            "absent_rate": grp_absent_rate,
            "pass_rate": grp_rate,
            "gpa5_count": grp_gpa5_w4,
            "gpa5_count_with_4th": grp_gpa5_w4,
            "gpa5_count_without_4th": grp_gpa5_no4,
            "one_fail_count": grp_1fail,
            "two_fail_count": grp_2fail,
            "multi_fail_count": grp_3fail
        }

    # Candidate Type Summary (Regular vs Irregular)
    types_summary = {}
    for c_type, s_list in type_students.items():
        t_tot = len(s_list)
        t_pass = sum(1 for s in s_list if s["Result"] == "PASSED")
        t_fail = sum(1 for s in s_list if s["Result"] == "FAIL")
        t_abs = sum(1 for s in s_list if s["Result"] == "ABSENT")
        t_app = t_tot - t_abs
        t_pass_rate = round((t_pass / t_app * 100), 2) if t_app > 0 else 0.0
        t_fail_rate = round((t_fail / t_app * 100), 2) if t_app > 0 else 0.0
        t_abs_rate = round((t_abs / t_tot * 100), 2) if t_tot > 0 else 0.0
        t_gpa5_w4 = sum(1 for s in s_list if s["Is_GPA5_With_4th"])
        t_gpa5_no4 = sum(1 for s in s_list if s["Is_GPA5_Without_4th"])
        t_gpa5_rate_w4 = round((t_gpa5_w4 / t_app * 100), 2) if t_app > 0 else 0.0
        t_gpa5_rate_no4 = round((t_gpa5_no4 / t_app * 100), 2) if t_app > 0 else 0.0
        t_1fail = sum(1 for s in s_list if s["Failed_Count"] == 1 and s["Result"] != "ABSENT")
        t_2fail = sum(1 for s in s_list if s["Failed_Count"] == 2 and s["Result"] != "ABSENT")
        t_mfail = sum(1 for s in s_list if s["Failed_Count"] >= 3 and s["Result"] != "ABSENT")

        # Group-wise breakdown inside this type
        grp_breakdown = {}
        for s in s_list:
            g = s["Group"]
            if g not in grp_breakdown:
                grp_breakdown[g] = {"total": 0, "passed": 0, "failed": 0, "absent": 0, "gpa5_with_4th": 0, "gpa5_without_4th": 0}
            grp_breakdown[g]["total"] += 1
            if s["Result"] == "PASSED":
                grp_breakdown[g]["passed"] += 1
                if s["Is_GPA5_With_4th"]:
                    grp_breakdown[g]["gpa5_with_4th"] += 1
                if s["Is_GPA5_Without_4th"]:
                    grp_breakdown[g]["gpa5_without_4th"] += 1
            elif s["Result"] == "FAIL":
                grp_breakdown[g]["failed"] += 1
            elif s["Result"] == "ABSENT":
                grp_breakdown[g]["absent"] += 1

        for g_data in grp_breakdown.values():
            g_data["pass_rate"] = round((g_data["passed"] / g_data["total"] * 100), 2) if g_data["total"] > 0 else 0.0

        types_summary[c_type] = {
            "type": c_type,
            "total": t_tot,
            "appeared": t_app,
            "passed": t_pass,
            "failed": t_fail,
            "absent": t_abs,
            "absent_rate": t_abs_rate,
            "pass_rate": t_pass_rate,
            "fail_rate": t_fail_rate,
            "gpa5_count": t_gpa5_w4,
            "gpa5_count_with_4th": t_gpa5_w4,
            "gpa5_count_without_4th": t_gpa5_no4,
            "gpa5_rate": t_gpa5_rate_w4,
            "gpa5_rate_with_4th": t_gpa5_rate_w4,
            "gpa5_rate_without_4th": t_gpa5_rate_no4,
            "one_fail_count": t_1fail,
            "two_fail_count": t_2fail,
            "multi_fail_count": t_mfail,
            "group_breakdown": grp_breakdown
        }

    # Cross-tab comparison: Group x Type (e.g. Science Regular vs Science Irregular)
    group_type_comparison = []
    for grp in sorted(list(group_students.keys())):
        for c_type in ["REGULAR", "IRREGULAR", "IMPROVEMENT", "PRIVATE"]:
            s_sub = [s for s in group_students[grp] if s["Type"] == c_type]
            if s_sub:
                sub_tot = len(s_sub)
                sub_pass = sum(1 for s in s_sub if s["Result"] == "PASSED")
                sub_fail = sum(1 for s in s_sub if s["Result"] == "FAIL")
                sub_abs = sum(1 for s in s_sub if s["Result"] == "ABSENT")
                sub_app = sub_tot - sub_abs
                sub_rate = round((sub_pass / sub_app * 100), 2) if sub_app > 0 else 0.0
                sub_gpa5_w4 = sum(1 for s in s_sub if s["Is_GPA5_With_4th"])
                sub_gpa5_no4 = sum(1 for s in s_sub if s["Is_GPA5_Without_4th"])
                group_type_comparison.append({
                    "group": grp,
                    "type": c_type,
                    "total": sub_tot,
                    "appeared": sub_app,
                    "absent": sub_abs,
                    "passed": sub_pass,
                    "failed": sub_fail,
                    "pass_rate": sub_rate,
                    "gpa5_count": sub_gpa5_w4,
                    "gpa5_count_with_4th": sub_gpa5_w4,
                    "gpa5_count_without_4th": sub_gpa5_no4,
                    "one_fail_count": sum(1 for s in s_sub if s["Failed_Count"] == 1 and s["Result"] != "ABSENT"),
                    "two_fail_count": sum(1 for s in s_sub if s["Failed_Count"] == 2 and s["Result"] != "ABSENT"),
                    "multi_fail_count": sum(1 for s in s_sub if s["Failed_Count"] >= 3 and s["Result"] != "ABSENT")
                })

    regular_summary = types_summary.get("REGULAR", {
        "type": "REGULAR", "total": 0, "appeared": 0, "absent": 0, "absent_rate": 0.0,
        "passed": 0, "failed": 0, "pass_rate": 0.0, "fail_rate": 0.0,
        "gpa5_count": 0, "gpa5_count_with_4th": 0, "gpa5_count_without_4th": 0,
        "gpa5_rate": 0.0, "gpa5_rate_with_4th": 0.0, "gpa5_rate_without_4th": 0.0,
        "one_fail_count": 0, "two_fail_count": 0, "multi_fail_count": 0, "group_breakdown": {}
    })
    irregular_summary = types_summary.get("IRREGULAR", {
        "type": "IRREGULAR", "total": 0, "appeared": 0, "absent": 0, "absent_rate": 0.0,
        "passed": 0, "failed": 0, "pass_rate": 0.0, "fail_rate": 0.0,
        "gpa5_count": 0, "gpa5_count_with_4th": 0, "gpa5_count_without_4th": 0,
        "gpa5_rate": 0.0, "gpa5_rate_with_4th": 0.0, "gpa5_rate_without_4th": 0.0,
        "one_fail_count": 0, "two_fail_count": 0, "multi_fail_count": 0, "group_breakdown": {}
    })

    # Absent Analysis aggregation
    absent_analysis = {
        "total_absent": total_absent,
        "absent_rate": absent_rate,
        "regular_absent": sum(1 for s in absent_students if s["Type"] == "REGULAR"),
        "irregular_absent": sum(1 for s in absent_students if s["Type"] == "IRREGULAR"),
        "groups_absent": {
            grp: {
                "total": len(s_list),
                "absent": sum(1 for s in s_list if s["Result"] == "ABSENT"),
                "rate": round(sum(1 for s in s_list if s["Result"] == "ABSENT") / len(s_list) * 100, 2) if len(s_list) > 0 else 0.0
            }
            for grp, s_list in group_students.items()
        },
        "types_absent": {
            t: {
                "total": len(s_list),
                "absent": sum(1 for s in s_list if s["Result"] == "ABSENT"),
                "rate": round(sum(1 for s in s_list if s["Result"] == "ABSENT") / len(s_list) * 100, 2) if len(s_list) > 0 else 0.0
            }
            for t, s_list in type_students.items()
        },
        "students": absent_students
    }

    # Group & Subject-Wise Grade Analysis table records
    subject_analysis = []
    for (grp, c_type, subj), counts in subject_stats.items():
        tot_subj_appeared = sum(counts.values())
        if tot_subj_appeared == 0:
            continue
        subj_failed = counts.get("F", 0)
        subj_passed = tot_subj_appeared - subj_failed
        subj_pass_rate = round((subj_passed / tot_subj_appeared * 100), 2)

        subject_analysis.append({
            "group": grp,
            "type": c_type,
            "subject": subj,
            "appeared": tot_subj_appeared,
            "passed": subj_passed,
            "failed": subj_failed,
            "pass_rate": subj_pass_rate,
            "grades": {g: counts.get(g, 0) for g in VALID_GRADES}
        })

    # Sort subjects: first by group, type, then subject
    subject_analysis.sort(key=lambda x: (x["group"], x["type"], x["subject"]))

    # Sort 1-fail and 2-fail subjects by frequency descending
    sorted_one_sub_freq = dict(sorted(one_sub_freq.items(), key=lambda item: item[1], reverse=True))
    sorted_two_sub_freq = dict(sorted(two_sub_freq.items(), key=lambda item: item[1], reverse=True))
    sorted_two_sub_pairs = dict(sorted(two_sub_pairs.items(), key=lambda item: item[1], reverse=True))

    available_groups = sorted(list(group_students.keys()))
    available_types = sorted(list(type_students.keys()))

    return {
        "total_students": total_students,
        "summary": {
            "total_students": total_students,
            "appeared": appeared,
            "passed": total_passed,
            "failed": total_failed,
            "absent": total_absent,
            "absent_rate": absent_rate,
            "pass_rate": pass_rate,
            "fail_rate": fail_rate,
            "gpa_5_count": gpa_5_count_with_4th,
            "gpa_5_rate": gpa_5_rate_with_4th,
            "gpa_5_count_with_4th": gpa_5_count_with_4th,
            "gpa_5_rate_with_4th": gpa_5_rate_with_4th,
            "gpa_5_count_without_4th": gpa_5_count_without_4th,
            "gpa_5_rate_without_4th": gpa_5_rate_without_4th,
            "one_subject_fail_count": len(one_subject_fail_students),
            "two_subject_fail_count": len(two_subject_fail_students),
            "multi_subject_fail_count": len(multi_subject_fail_students)
        },
        "groups_summary": groups_summary,
        "types_summary": types_summary,
        "regular_summary": regular_summary,
        "irregular_summary": irregular_summary,
        "group_type_comparison": group_type_comparison,
        "subject_analysis": subject_analysis,
        "one_subject_fail": {
            "total": len(one_subject_fail_students),
            "subjects_summary": sorted_one_sub_freq,
            "students": sorted(one_subject_fail_students, key=lambda s: (s["Failed_Subject"], s["Roll"]))
        },
        "two_subject_fail": {
            "total": len(two_subject_fail_students),
            "subjects_summary": sorted_two_sub_freq,
            "pair_summary": sorted_two_sub_pairs,
            "students": sorted(two_subject_fail_students, key=lambda s: (s["Failed_Subjects"][0] if s["Failed_Subjects"] else "", s["Roll"]))
        },
        "multi_subject_fail": {
            "total": len(multi_subject_fail_students),
            "students": sorted(multi_subject_fail_students, key=lambda s: (-s["Failed_Count"], s["Roll"]))
        },
        "gpa_distribution": gpa_dist_with_4th,
        "gpa_distribution_with_4th": gpa_dist_with_4th,
        "gpa_distribution_without_4th": gpa_dist_without_4th,
        "absent_analysis": absent_analysis,
        "available_groups": available_groups,
        "available_types": available_types
    }


def generate_analysis_excel_bytes(analysis: Dict[str, Any], meta: Optional[Dict[str, str]] = None) -> io.BytesIO:
    """
    Generates an Excel workbook (.xlsx) containing formatted analysis worksheets:
    1. Executive Summary & Overview
    2. Regular vs Irregular Analysis
    3. Group & Subject Grade Analysis
    4. 1-Subject Fail Breakdown
    5. 2-Subject Fail Breakdown
    6. 3+ Subject Fail List
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove default sheet

    meta = meta or {}
    exam = meta.get("exam", "HSC / SSC").upper()
    year = meta.get("year", "")
    board = meta.get("board", "").title()
    title_sub = f"Examination: {exam} {year} | Board: {board}"

    # Palette Styles
    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    card_header_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    bold_font = Font(name="Calibri", size=11, bold=True)
    normal_font = Font(name="Calibri", size=11)

    fail_badge_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fail_font = Font(name="Calibri", size=11, bold=True, color="991B1B")

    pass_badge_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    pass_font = Font(name="Calibri", size=11, bold=True, color="166534")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    summary = analysis.get("summary", {})

    # =========================================================================
    # SHEET 1: EXECUTIVE SUMMARY
    # =========================================================================
    ws_sum = wb.create_sheet(title="Executive Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_sum.merge_cells("A1:H1")
    ws_sum["A1"] = "BOARD RESULT COMPREHENSIVE ANALYSIS REPORT"
    ws_sum["A1"].font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    ws_sum["A1"].fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    ws_sum["A1"].alignment = align_center
    ws_sum.row_dimensions[1].height = 36

    ws_sum.merge_cells("A2:H2")
    ws_sum["A2"] = title_sub
    ws_sum["A2"].font = Font(name="Calibri", size=11, italic=True, color="FFFFFF")
    ws_sum["A2"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws_sum["A2"].alignment = align_center
    ws_sum.row_dimensions[2].height = 22

    # KPI Summary Table
    ws_sum.cell(row=4, column=1, value="OVERALL RESULT METRICS").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    kpis = [
        ("Total Students Processed", summary.get("total_students", 0)),
        ("Students Appeared", summary.get("appeared", 0)),
        ("Students Absent", f"{summary.get('absent', 0)} ({summary.get('absent_rate', 0)}%)"),
        ("Total Passed", f"{summary.get('passed', 0)} ({summary.get('pass_rate', 0)}%)"),
        ("Total Failed", f"{summary.get('failed', 0)} ({summary.get('fail_rate', 0)}%)"),
        ("GPA 5.00 With 4th Subject", f"{summary.get('gpa_5_count_with_4th', 0)} ({summary.get('gpa_5_rate_with_4th', 0)}%)"),
        ("GPA 5.00 Without 4th Sub (Golden A+)", f"{summary.get('gpa_5_count_without_4th', 0)} ({summary.get('gpa_5_rate_without_4th', 0)}%)"),
        ("One-Subject Fail (1 Sub)", summary.get("one_subject_fail_count", 0)),
        ("Two-Subject Fail (2 Subs)", summary.get("two_subject_fail_count", 0)),
        ("Three or More Subjects Fail (3+)", summary.get("multi_subject_fail_count", 0)),
    ]

    r_idx = 5
    for label, val in kpis:
        c1 = ws_sum.cell(row=r_idx, column=1, value=label)
        c2 = ws_sum.cell(row=r_idx, column=2, value=val)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_idx += 1

    # GPA Distribution (With 4th Subject)
    ws_sum.cell(row=4, column=4, value="GPA (WITH 4TH SUB)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    r_idx_gpa1 = 5
    for tier, count in analysis.get("gpa_distribution_with_4th", analysis.get("gpa_distribution", {})).items():
        c1 = ws_sum.cell(row=r_idx_gpa1, column=4, value=tier)
        c2 = ws_sum.cell(row=r_idx_gpa1, column=5, value=count)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_idx_gpa1 += 1

    # GPA Distribution (Without 4th Subject)
    ws_sum.cell(row=4, column=7, value="GPA (WITHOUT 4TH SUB)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    r_idx_gpa2 = 5
    for tier, count in analysis.get("gpa_distribution_without_4th", {}).items():
        c1 = ws_sum.cell(row=r_idx_gpa2, column=7, value=tier)
        c2 = ws_sum.cell(row=r_idx_gpa2, column=8, value=count)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_idx_gpa2 += 1

    # Regular vs Irregular Overview in Executive Summary
    type_start_row = max(r_idx, r_idx_gpa1, r_idx_gpa2) + 2
    ws_sum.cell(row=type_start_row, column=1, value="REGULAR VS IRREGULAR CANDIDATE PERFORMANCE").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    type_headers = ["Candidate Type", "Total", "Appeared", "Absent", "Passed", "Failed", "Pass %", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1 Sub Fail", "2 Sub Fail", "3+ Sub Fail"]
    h_row_type = type_start_row + 1
    for col_idx, h_text in enumerate(type_headers, start=1):
        c = ws_sum.cell(row=h_row_type, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = h_row_type + 1
    for t_data in analysis.get("types_summary", {}).values():
        row_vals = [
            t_data["type"],
            t_data["total"],
            t_data["appeared"],
            t_data.get("absent", 0),
            t_data["passed"],
            t_data["failed"],
            f"{t_data['pass_rate']}%",
            t_data.get("gpa5_count_with_4th", t_data.get("gpa5_count", 0)),
            t_data.get("gpa5_count_without_4th", 0),
            t_data["one_fail_count"],
            t_data["two_fail_count"],
            t_data["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_sum.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
            if col_idx == 7:
                if t_data["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif t_data["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1

    # Group Performance Comparison Table
    group_start_row = curr_r + 2
    ws_sum.cell(row=group_start_row, column=1, value="GROUP-WISE PERFORMANCE COMPARISON").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    grp_headers = ["Group", "Total", "Appeared", "Absent", "Passed", "Failed", "Pass %", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1 Sub Fail", "2 Sub Fail", "3+ Sub Fail"]
    h_row = group_start_row + 1
    for col_idx, h_text in enumerate(grp_headers, start=1):
        c = ws_sum.cell(row=h_row, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = h_row + 1
    for grp_data in analysis.get("groups_summary", {}).values():
        row_vals = [
            grp_data["group"],
            grp_data["total"],
            grp_data["appeared"],
            grp_data.get("absent", 0),
            grp_data["passed"],
            grp_data["failed"],
            f"{grp_data['pass_rate']}%",
            grp_data.get("gpa5_count_with_4th", grp_data.get("gpa5_count", 0)),
            grp_data.get("gpa5_count_without_4th", 0),
            grp_data["one_fail_count"],
            grp_data["two_fail_count"],
            grp_data["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_sum.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
            if col_idx == 7:
                if grp_data["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif grp_data["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1
        curr_r += 1

    for col in ws_sum.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_sum.column_dimensions[col_letter].width = max(max_len + 4, 14)

    # =========================================================================
    # SHEET 2: REGULAR VS IRREGULAR ANALYSIS
    # =========================================================================
    ws_type_sheet = wb.create_sheet(title="Regular vs Irregular")
    ws_type_sheet.views.sheetView[0].showGridLines = True

    # Title
    ws_type_sheet.merge_cells("A1:L1")
    ws_type_sheet["A1"] = "REGULAR VS IRREGULAR CANDIDATE RESULT ANALYSIS"
    ws_type_sheet["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws_type_sheet["A1"].fill = header_fill
    ws_type_sheet["A1"].alignment = align_center
    ws_type_sheet.row_dimensions[1].height = 30

    # Section 1: Candidate Type Overall Comparison
    ws_type_sheet.cell(row=3, column=1, value="CANDIDATE TYPE OVERALL PERFORMANCE").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    type_cols = ["Candidate Type", "Total Students", "Appeared", "Absent", "Passed", "Failed", "Pass Rate (%)", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1-Sub Fail", "2-Sub Fail", "3+ Sub Fail"]
    for col_idx, h_text in enumerate(type_cols, start=1):
        c = ws_type_sheet.cell(row=4, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = 5
    for t_data in analysis.get("types_summary", {}).values():
        row_vals = [
            t_data["type"],
            t_data["total"],
            t_data["appeared"],
            t_data.get("absent", 0),
            t_data["passed"],
            t_data["failed"],
            f"{t_data['pass_rate']}%",
            t_data.get("gpa5_count_with_4th", t_data.get("gpa5_count", 0)),
            t_data.get("gpa5_count_without_4th", 0),
            t_data["one_fail_count"],
            t_data["two_fail_count"],
            t_data["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_type_sheet.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
            if col_idx == 7:
                if t_data["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif t_data["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1

    # Section 2: Group x Candidate Type Matrix
    curr_r += 2
    ws_type_sheet.cell(row=curr_r, column=1, value="GROUP-WISE REGULAR VS IRREGULAR BREAKDOWN").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    curr_r += 1

    group_type_cols = ["Group Name", "Candidate Type", "Total Students", "Appeared", "Absent", "Passed", "Failed", "Pass Rate (%)", "GPA 5 (With 4th)", "GPA 5 (Without 4th)", "1-Sub Fail", "2-Sub Fail", "3+ Sub Fail"]
    for col_idx, h_text in enumerate(group_type_cols, start=1):
        c = ws_type_sheet.cell(row=curr_r, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r += 1
    for row_item in analysis.get("group_type_comparison", []):
        row_vals = [
            row_item["group"],
            row_item["type"],
            row_item["total"],
            row_item["appeared"],
            row_item.get("absent", 0),
            row_item["passed"],
            row_item["failed"],
            f"{row_item['pass_rate']}%",
            row_item.get("gpa5_count_with_4th", row_item.get("gpa5_count", 0)),
            row_item.get("gpa5_count_without_4th", 0),
            row_item["one_fail_count"],
            row_item["two_fail_count"],
            row_item["multi_fail_count"],
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_type_sheet.cell(row=curr_r, column=col_idx, value=val)
            c.font = bold_font if col_idx in [1, 2] else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx in [1, 2] else align_center
            if col_idx == 8:
                if row_item["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif row_item["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
        curr_r += 1

    for col in ws_type_sheet.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_type_sheet.column_dimensions[col_letter].width = max(max_len + 3, 13)

    # =========================================================================
    # SHEET 3: GROUP & SUBJECT-WISE GRADE ANALYSIS
    # =========================================================================
    ws_subj = wb.create_sheet(title="Subject Wise Grade Analysis")
    ws_subj.views.sheetView[0].showGridLines = True

    subj_headers = [
        "Group", "Candidate Type", "Subject Code & Name", "Appeared", "Passed", "Failed", "Pass Rate (%)",
        "A+", "A", "A-", "B", "C", "D", "F"
    ]
    for col_idx, h_text in enumerate(subj_headers, start=1):
        c = ws_subj.cell(row=1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border
    ws_subj.row_dimensions[1].height = 28

    curr_r = 2
    for item in analysis.get("subject_analysis", []):
        grades = item["grades"]
        row_vals = [
            item["group"],
            item.get("type", "ALL"),
            item["subject"],
            item["appeared"],
            item["passed"],
            item["failed"],
            f"{item['pass_rate']}%",
            grades.get("A+", 0),
            grades.get("A", 0),
            grades.get("A-", 0),
            grades.get("B", 0),
            grades.get("C", 0),
            grades.get("D", 0),
            grades.get("F", 0)
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_subj.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx in [1, 2, 3] else align_center
            if col_idx == 7:  # Pass rate
                if item["pass_rate"] >= 80:
                    c.fill = pass_badge_fill
                    c.font = pass_font
                elif item["pass_rate"] < 50:
                    c.fill = fail_badge_fill
                    c.font = fail_font
            elif col_idx == 14 and grades.get("F", 0) > 0:  # F count
                c.fill = fail_badge_fill
                c.font = fail_font

        curr_r += 1

    for col in ws_subj.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_subj.column_dimensions[col_letter].width = max(max_len + 3, 11)

    # =========================================================================
    # SHEET 4: ONE SUBJECT FAIL BREAKDOWN
    # =========================================================================
    ws_1fail = wb.create_sheet(title="1 Subject Fail")
    ws_1fail.views.sheetView[0].showGridLines = True

    one_fail_data = analysis.get("one_subject_fail", {})

    # Top summary of 1-fail by subject
    ws_1fail.cell(row=1, column=1, value="ONE-SUBJECT FAIL SUMMARY BY SUBJECT").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    ws_1fail.cell(row=2, column=1, value="Failed Subject").fill = header_fill
    ws_1fail.cell(row=2, column=1).font = header_font
    ws_1fail.cell(row=2, column=2, value="Student Count").fill = header_fill
    ws_1fail.cell(row=2, column=2).font = header_font

    r_sum = 3
    for subj_name, cnt in one_fail_data.get("subjects_summary", {}).items():
        c1 = ws_1fail.cell(row=r_sum, column=1, value=subj_name)
        c2 = ws_1fail.cell(row=r_sum, column=2, value=cnt)
        c1.font = bold_font
        c2.font = fail_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_sum += 1

    # Detailed Student Roster
    start_roster = r_sum + 2
    ws_1fail.cell(row=start_roster, column=1, value="STUDENT ROSTER (1-SUBJECT FAIL)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster_headers = ["Roll No", "Reg No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "Failed Subject (F)", "Passed Subject Grades", "Remarks"]
    for col_idx, h_text in enumerate(roster_headers, start=1):
        c = ws_1fail.cell(row=start_roster + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = start_roster + 2
    for st in one_fail_data.get("students", []):
        row_vals = [
            st["Roll"],
            st["Registration"],
            st["Name"],
            st["Group"],
            st.get("Type", "REGULAR"),
            st.get("GPA_With_4th", st.get("GPA", "0.00")),
            st.get("GPA_Without_4th", "0.00"),
            st["Failed_Subject"],
            st["Passed_Grades_Summary"],
            "Eligible for 1-Subject Improvement"
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_1fail.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_center if col_idx in [1, 2, 4, 5, 6, 7, 10] else align_left
            if col_idx == 8:
                c.fill = fail_badge_fill
                c.font = fail_font
        curr_r += 1

    for col in ws_1fail.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_1fail.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 5: TWO SUBJECT FAIL BREAKDOWN
    # =========================================================================
    ws_2fail = wb.create_sheet(title="2 Subject Fail")
    ws_2fail.views.sheetView[0].showGridLines = True

    two_fail_data = analysis.get("two_subject_fail", {})

    ws_2fail.cell(row=1, column=1, value="TWO-SUBJECT FAIL SUMMARY (TOP FAILED SUBJECTS)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    ws_2fail.cell(row=2, column=1, value="Subject").fill = header_fill
    ws_2fail.cell(row=2, column=1).font = header_font
    ws_2fail.cell(row=2, column=2, value="Occurrences in 2-Fail").fill = header_fill
    ws_2fail.cell(row=2, column=2).font = header_font

    r_sum = 3
    for subj_name, cnt in two_fail_data.get("subjects_summary", {}).items():
        c1 = ws_2fail.cell(row=r_sum, column=1, value=subj_name)
        c2 = ws_2fail.cell(row=r_sum, column=2, value=cnt)
        c1.font = bold_font
        c2.font = fail_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_sum += 1

    # Roster
    start_roster = r_sum + 2
    ws_2fail.cell(row=start_roster, column=1, value="STUDENT ROSTER (2-SUBJECT FAIL)").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster2_headers = ["Roll No", "Reg No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "Failed Subject 1 (F)", "Failed Subject 2 (F)", "Passed Subject Grades"]
    for col_idx, h_text in enumerate(roster2_headers, start=1):
        c = ws_2fail.cell(row=start_roster + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = start_roster + 2
    for st in two_fail_data.get("students", []):
        f_list = st.get("Failed_Subjects", [])
        f1 = f_list[0] if len(f_list) > 0 else "-"
        f2 = f_list[1] if len(f_list) > 1 else "-"

        row_vals = [
            st["Roll"],
            st["Registration"],
            st["Name"],
            st["Group"],
            st.get("Type", "REGULAR"),
            st.get("GPA_With_4th", st.get("GPA", "0.00")),
            st.get("GPA_Without_4th", "0.00"),
            f1,
            f2,
            st["Passed_Grades_Summary"]
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_2fail.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_center if col_idx in [1, 2, 4, 5, 6, 7] else align_left
            if col_idx in [8, 9]:
                c.fill = fail_badge_fill
                c.font = fail_font
        curr_r += 1

    for col in ws_2fail.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_2fail.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 6: 3+ SUBJECT FAIL (MULTI-FAIL)
    # =========================================================================
    ws_mfail = wb.create_sheet(title="3+ Subject Fail")
    ws_mfail.views.sheetView[0].showGridLines = True

    multi_fail_data = analysis.get("multi_subject_fail", {})
    ws_mfail.cell(row=1, column=1, value="STUDENTS FAILING 3 OR MORE SUBJECTS").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster3_headers = ["Roll No", "Student Name", "Group", "Candidate Type", "GPA (With 4th)", "GPA (Without 4th)", "Number of Failed Subjects", "Failed Subjects List", "Passed Grades Summary"]
    for col_idx, h_text in enumerate(roster3_headers, start=1):
        c = ws_mfail.cell(row=2, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = 3
    for st in multi_fail_data.get("students", []):
        f_list = st.get("Failed_Subjects", [])
        row_vals = [
            st["Roll"],
            st["Name"],
            st["Group"],
            st.get("Type", "REGULAR"),
            st.get("GPA_With_4th", st.get("GPA", "0.00")),
            st.get("GPA_Without_4th", "0.00"),
            st["Failed_Count"],
            ", ".join(f_list),
            st["Passed_Grades_Summary"]
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_mfail.cell(row=curr_r, column=col_idx, value=val)
            c.font = normal_font
            c.border = thin_border
            c.alignment = align_center if col_idx in [1, 3, 4, 5, 6, 7] else align_left
            if col_idx == 7:
                c.fill = fail_badge_fill
                c.font = fail_font
        curr_r += 1

    for col in ws_mfail.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_mfail.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    # =========================================================================
    # SHEET 7: ABSENT STUDENTS ANALYSIS & ROSTER
    # =========================================================================
    ws_absent = wb.create_sheet(title="Absent Students")
    ws_absent.views.sheetView[0].showGridLines = True

    absent_data = analysis.get("absent_analysis", {})

    # Title
    ws_absent.merge_cells("A1:I1")
    ws_absent["A1"] = "ABSENT STUDENTS ANALYSIS & ROSTER"
    ws_absent["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws_absent["A1"].fill = header_fill
    ws_absent["A1"].alignment = align_center
    ws_absent.row_dimensions[1].height = 30

    # Section 1: Summary Statistics
    ws_absent.cell(row=3, column=1, value="ABSENT SUMMARY METRICS").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    absent_kpis = [
        ("Total Registered Students", summary.get("total_students", 0)),
        ("Students Appeared", summary.get("appeared", 0)),
        ("Total Absent Candidates", absent_data.get("total_absent", summary.get("absent", 0))),
        ("Absent Rate (%)", f"{absent_data.get('absent_rate', summary.get('absent_rate', 0))}%"),
        ("Regular Candidates Absent", absent_data.get("regular_absent", 0)),
        ("Irregular Candidates Absent", absent_data.get("irregular_absent", 0)),
    ]
    r_abs = 4
    for label, val in absent_kpis:
        c1 = ws_absent.cell(row=r_abs, column=1, value=label)
        c2 = ws_absent.cell(row=r_abs, column=2, value=val)
        c1.font = bold_font
        c2.font = normal_font
        c1.border = thin_border
        c2.border = thin_border
        c2.alignment = align_center
        r_abs += 1

    # Section 2: Group-wise Absent Breakdown
    r_grp_abs_title = r_abs + 1
    ws_absent.cell(row=r_grp_abs_title, column=1, value="GROUP-WISE ABSENT BREAKDOWN").font = Font(name="Calibri", size=12, bold=True, color="0F172A")
    grp_abs_headers = ["Group Name", "Total Registered", "Absent Count", "Absent Rate (%)"]
    for col_idx, h_text in enumerate(grp_abs_headers, start=1):
        c = ws_absent.cell(row=r_grp_abs_title + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r_grp_abs = r_grp_abs_title + 2
    for grp_name, g_info in absent_data.get("groups_absent", {}).items():
        row_vals = [
            grp_name,
            g_info.get("total", 0),
            g_info.get("absent", 0),
            f"{g_info.get('rate', 0)}%"
        ]
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_absent.cell(row=curr_r_grp_abs, column=col_idx, value=val)
            c.font = bold_font if col_idx == 1 else normal_font
            c.border = thin_border
            c.alignment = align_left if col_idx == 1 else align_center
        curr_r_grp_abs += 1

    # Section 3: Detailed Absent Student Roster
    start_roster_abs = curr_r_grp_abs + 2
    ws_absent.cell(row=start_roster_abs, column=1, value="ABSENT STUDENT ROSTER").font = Font(name="Calibri", size=12, bold=True, color="0F172A")

    roster_abs_headers = ["Roll No", "Reg No", "Student Name", "Father's Name", "Group", "Candidate Type", "Institute", "Status", "Remarks"]
    for col_idx, h_text in enumerate(roster_abs_headers, start=1):
        c = ws_absent.cell(row=start_roster_abs + 1, column=col_idx, value=h_text)
        c.fill = header_fill
        c.font = header_font
        c.alignment = align_center
        c.border = thin_border

    curr_r = start_roster_abs + 2
    abs_students = absent_data.get("students", [])
    if not abs_students:
        c = ws_absent.cell(row=curr_r, column=1, value="No absent students recorded.")
        c.font = normal_font
        ws_absent.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=9)
    else:
        for st in abs_students:
            row_vals = [
                st.get("Roll", "-"),
                st.get("Registration", "-"),
                st.get("Name", "N/A"),
                st.get("Father Name", "-"),
                st.get("Group", "GENERAL"),
                st.get("Type", "REGULAR"),
                st.get("Institute", "-"),
                "ABSENT",
                st.get("Remarks", "Did not appear")
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws_absent.cell(row=curr_r, column=col_idx, value=val)
                c.font = normal_font
                c.border = thin_border
                c.alignment = align_center if col_idx in [1, 2, 5, 6, 8] else align_left
                if col_idx == 8:
                    c.fill = fail_badge_fill
                    c.font = fail_font
            curr_r += 1

    for col in ws_absent.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_absent.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream
