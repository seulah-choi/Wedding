"""dashboard.html → apps-script/index.html 변환.

dashboard.html 을 수정했다면 이 스크립트를 실행한 뒤
apps-script/index.html 내용을 Apps Script 편집기의 index.html 에 다시 붙여넣으세요.
실행: python3 build_apps_script.py
"""
from pathlib import Path

here = Path(__file__).parent
src = (here / "dashboard.html").read_text(encoding="utf-8")
# 웹앱에서는 구글 시트가 데이터 원본 → 샘플 데이터·엑셀 라이브러리 제외
for line in [
    '<script src="sample_data.js"></script>\n',
    '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>\n',
]:
    assert line in src, line
    src = src.replace(line, "")
banner = "<!-- 자동 생성 파일: dashboard.html 을 수정한 뒤 build_apps_script.py 로 다시 만드세요 -->\n"
(here / "apps-script" / "index.html").write_text(banner + src, encoding="utf-8")
print("apps-script/index.html written")
