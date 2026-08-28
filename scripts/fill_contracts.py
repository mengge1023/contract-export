#!/usr/bin/env python3
"""Fill verified student slots in a DOCX template and export a checked batch."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
import zipfile
from xml.sax.saxutils import escape


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True)
    parser.add_argument("--students", required=True, help="Normalized UTF-8 JSON array")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--zip", dest="zip_path", required=True)
    parser.add_argument("--exceptions", required=True)
    parser.add_argument("--name-prefix", default="乙方：")
    parser.add_argument("--id-prefix", default="乙方身份证号码：")
    return parser.parse_args()


def safe_filename(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip().rstrip(".")
    return value or "未命名"


def first_phone_suffix(phone: str) -> str:
    match = re.search(r"(?<!\d)(1\d{10})(?!\d)", phone)
    if match:
        return match.group(1)[-4:]
    digits = re.sub(r"\D", "", phone)
    return digits[-4:] if len(digits) >= 4 else "无号码"


def filename_serial(serial_number: str) -> str:
    value = str(serial_number).strip()
    return value.zfill(3) if value.isdigit() else value


def classify(students: list[dict]) -> tuple[list[dict], list[dict]]:
    counts: dict[tuple[str, str], int] = {}
    for key in ("serialNumber", "phone", "idNumber"):
        for student in students:
            value = str(student.get(key, "")).strip().lower()
            if value:
                counts[(key, value)] = counts.get((key, value), 0) + 1

    valid, invalid = [], []
    for student in students:
        s = {k: str(student.get(k, "")).strip() for k in ("serialNumber", "realName", "wechatName", "phone", "idNumber")}
        s["excelRow"] = student.get("excelRow")
        issues = []
        for key, label in (("serialNumber", "序号为空"), ("realName", "真实姓名为空"), ("wechatName", "微信名为空"), ("phone", "手机号为空"), ("idNumber", "身份证号为空")):
            if not s[key]:
                issues.append(label)
        if s["idNumber"] and not re.fullmatch(r"\d{17}[\dXx]", s["idNumber"]):
            issues.append(f"身份证号格式异常（{len(s['idNumber'])}位）")
        if s["serialNumber"] and counts.get(("serialNumber", s["serialNumber"].lower()), 0) > 1:
            issues.append("序号重复")
        if s["phone"] and counts.get(("phone", s["phone"].lower()), 0) > 1:
            issues.append("手机号重复")
        if s["idNumber"] and counts.get(("idNumber", s["idNumber"].lower()), 0) > 1:
            issues.append("身份证号重复")
        s["issues"] = issues
        (invalid if issues else valid).append(s)
    return valid, invalid


def patch_xml(xml_bytes: bytes, student: dict, name_prefix: str, id_prefix: str) -> bytes:
    xml = xml_bytes.decode("utf-8")
    name_slot = f"{name_prefix}</w:t>"
    id_slot = f"{id_prefix}</w:t>"
    if xml.count(name_slot) != 1 or xml.count(id_slot) != 1:
        raise RuntimeError("Template slots are missing or ambiguous; inspect the template and specify exact prefixes")
    xml = xml.replace(name_slot, f"{escape(name_prefix + student['realName'])}</w:t>", 1)
    xml = xml.replace(id_slot, f"{escape(id_prefix + student['idNumber'].upper())}</w:t>", 1)
    return xml.encode("utf-8")


def build_docx(template: Path, output: Path, student: dict, name_prefix: str, id_prefix: str) -> None:
    with zipfile.ZipFile(template, "r") as src, zipfile.ZipFile(output, "w") as dst:
        for info in src.infolist():
            data = src.read(info.filename)
            if info.filename == "word/document.xml":
                data = patch_xml(data, student, name_prefix, id_prefix)
            dst.writestr(info, data)


def verify(template: Path, output: Path, student: dict, name_prefix: str, id_prefix: str) -> None:
    with zipfile.ZipFile(template) as src, zipfile.ZipFile(output) as dst:
        src_parts = {i.filename: src.read(i.filename) for i in src.infolist()}
        dst_parts = {i.filename: dst.read(i.filename) for i in dst.infolist()}
    if set(src_parts) != set(dst_parts):
        raise RuntimeError(f"Package parts changed: {output.name}")
    for part, source_data in src_parts.items():
        if part != "word/document.xml" and dst_parts[part] != source_data:
            raise RuntimeError(f"Preserve-only part changed: {output.name} {part}")
    xml = dst_parts["word/document.xml"].decode("utf-8")
    if xml.count(f"{name_prefix}{student['realName']}</w:t>") != 1:
        raise RuntimeError(f"Name slot mismatch: {output.name}")
    if xml.count(f"{id_prefix}{student['idNumber'].upper()}</w:t>") != 1:
        raise RuntimeError(f"ID slot mismatch: {output.name}")


def main() -> None:
    args = parse_args()
    template = Path(args.template).resolve()
    students_path = Path(args.students).resolve()
    output_dir = Path(args.output_dir).resolve()
    zip_path = Path(args.zip_path).resolve()
    exceptions_path = Path(args.exceptions).resolve()
    for target in (output_dir, zip_path, exceptions_path):
        if target.exists():
            raise FileExistsError(f"Refusing to overwrite existing target: {target}")

    students = json.loads(students_path.read_text(encoding="utf-8"))
    if not isinstance(students, list):
        raise TypeError("students JSON must be an array")
    valid, invalid = classify(students)
    output_dir.mkdir(parents=True)

    manifest = []
    for student in valid:
        seq = filename_serial(student["serialNumber"])
        filename = safe_filename(
            f"{seq}_{student['wechatName']}_{student['realName']}_{first_phone_suffix(student['phone'])}_合同.docx"
        )
        output = output_dir / filename
        build_docx(template, output, student, args.name_prefix, args.id_prefix)
        verify(template, output, student, args.name_prefix, args.id_prefix)
        manifest.append({
            "seq": student["serialNumber"],
            "excelRow": student.get("excelRow"),
            "wechatName": student["wechatName"],
            "realName": student["realName"],
            "filename": filename,
            "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
        })

    (output_dir / "生成清单.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [f"以下{len(invalid)}条记录存在异常，本批未生成合同：", ""]
    for student in invalid:
        lines.append(
            f"Excel第{student.get('excelRow')}行｜微信名：{student['wechatName'] or '空'}｜"
            f"真实姓名：{student['realName'] or '空'}｜问题：{'；'.join(student['issues'])}"
        )
    exceptions_path.parent.mkdir(parents=True, exist_ok=True)
    exceptions_path.write_text("\n".join(lines) + "\n", encoding="utf-8")

    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for contract in sorted(output_dir.glob("*.docx")):
            archive.write(contract, arcname=contract.name)
        archive.write(exceptions_path, arcname=exceptions_path.name)

    print(json.dumps({
        "ok": True,
        "total": len(students),
        "generated": len(valid),
        "excluded": len(invalid),
        "outputDir": str(output_dir),
        "zip": str(zip_path),
        "exceptions": str(exceptions_path),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
