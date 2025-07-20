# grouns_wiki_setup :
# CharacterSheets → MarkdownDocs 변환기

## 개요
`CharacterSheets` 폴더 내 Excel 파일(.xlsx/.xls)에서 특정 시트를 선택하여  
`MarkdownDocs` 폴더에 Markdown(`.md`) 형식으로 저장하는 스크립트 및 배치 파일 모음입니다.

## 파일 구성
```
프로젝트 루트/
├─ CharacterSheets/       # 변환할 Excel 파일(.xlsx/.xls) 저장
├─ MarkdownDocs/          # 변환된 .md 파일 저장 디렉토리
├─ excel_to_md.py         # 변환용 Python 스크립트
├─ run_convert.bat        # Windows CMD에서 일괄 실행용 배치 파일
└─ README.md              # 사용법 안내 문서 (현재 파일)
```

## 요구 사항
- Python 3.7 이상
- pandas (`pip install pandas`)
- openpyxl (`pip install openpyxl`)

## 사용 방법

### 1. 폴더 준비
- 프로젝트 루트에 `CharacterSheets` 폴더를 만들고 변환할 Excel 파일을 넣습니다.
- `MarkdownDocs` 폴더는 스크립트 실행 시 자동 생성됩니다.

### 2. Python 스크립트 실행
```bash
pip install pandas openpyxl
python excel_to_md.py
```
- Excel 파일 목록이 출력되면 번호를 입력하여 파일 선택
- 시트 목록이 출력되면 변환할 시트 번호(콤마 구분 또는 범위 입력) 선택
- 선택한 시트가 MarkdownDocs 폴더에 `.md` 파일로 저장됩니다.

### 3. Windows 배치 파일로 일괄 실행
- `run_convert.bat` 파일을 더블 클릭하면 CMD 창에서 자동으로:
  1. `excel_to_md.py` 실행
  2. 완료 후 일시 정지(pause)  
- 스크립트 출력 결과를 확인한 뒤 아무 키나 눌러 창을 닫을 수 있습니다.

## 주의 사항
- 스크립트와 배치 파일은 프로젝트 루트에 위치해야 합니다.
- 시트 이름에 사용된 특수문자는 파일명으로 저장할 때 `_`로 자동 변환됩니다.
- 대량의 시트를 변환할 경우 실행 시간이 소요될 수 있습니다.
