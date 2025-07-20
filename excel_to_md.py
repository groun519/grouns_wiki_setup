# excel_to_md.py
import os
import re
import sys
import pandas as pd

def sanitize_filename(name: str) -> str:
    # 파일명으로 사용할 수 없는 문자를 언더스코어로 대체
    return re.sub(r'[\\/:"*?<>|]+', '_', name)

def list_excel_files(dir_path: str):
    files = [f for f in os.listdir(dir_path) if f.lower().endswith(('.xlsx', '.xls'))]
    return files

def choose_from_list(items: list, prompt: str):
    for i, item in enumerate(items, 1):
        print(f"{i}. {item}")
    choice = input(prompt).strip()
    if not choice.isdigit() or not (1 <= int(choice) <= len(items)):
        print("잘못된 선택입니다.")
        sys.exit(1)
    return int(choice) - 1

def parse_indices(s: str, max_idx: int):
    idxs = []
    for part in s.split(','):
        part = part.strip()
        if '-' in part:
            start, end = part.split('-', 1)
            if start.isdigit() and end.isdigit():
                idxs.extend(range(int(start), int(end) + 1))
        elif part.isdigit():
            idxs.append(int(part))
    # 1-based to 0-based, 필터링
    return sorted({i-1 for i in idxs if 1 <= i <= max_idx})

def main():
    cwd = os.getcwd()
    src_dir = os.path.join(cwd, 'CharacterSheets')
    dst_dir = os.path.join(cwd, 'MarkdownDocs')

    if not os.path.isdir(src_dir):
        print(f"'{src_dir}' 폴더가 없습니다.")
        sys.exit(1)
    os.makedirs(dst_dir, exist_ok=True)

    # 1) Excel 파일 선택
    excels = list_excel_files(src_dir)
    if not excels:
        print("CharacterSheets 폴더에 Excel 파일이 없습니다.")
        sys.exit(1)
    print("변환할 Excel 파일을 선택하세요:")
    fi = choose_from_list(excels, "번호 입력: ")
    excel_path = os.path.join(src_dir, excels[fi])
    base_name = os.path.splitext(excels[fi])[0]

    # 2) 시트 선택
    xls = pd.ExcelFile(excel_path)
    sheets = xls.sheet_names
    print("\n시트 목록:")
    for i, name in enumerate(sheets, 1):
        print(f"{i}. {name}")
    sel = input("변환할 시트 번호(여러 개 시 콤마 구분, 범위 가능 ex.1,3-5): ").strip()
    indices = parse_indices(sel, len(sheets))
    if not indices:
        print("변환할 시트가 선택되지 않았습니다.")
        sys.exit(1)

    # 3) 선택된 시트만 Markdown으로 저장
    for idx in indices:
        sheet_name = sheets[idx]
        df = pd.read_excel(xls, sheet_name=sheet_name)
        safe_name = sanitize_filename(f"{base_name}_{sheet_name}")
        out_path = os.path.join(dst_dir, f"{safe_name}.md")
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(f"# {sheet_name}\n\n")
            f.write(df.to_markdown(index=False))
        print(f"-> '{sheet_name}' 시트를 '{out_path}' 으로 저장했습니다.")

if __name__ == "__main__":
    main()
