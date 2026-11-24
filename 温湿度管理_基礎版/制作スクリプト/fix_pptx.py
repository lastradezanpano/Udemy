"""pptxgenjs が出力した PPTX を PowerPoint で「修復」を求められない形に直す後処理。

直す点:
1. [Content_Types].xml に登録されているが実在しない部品(slideMaster2.xml 以降など)の登録を消す
2. 同じスライド内で図形の ID(p:cNvPr id)が重複している場合、後から出てきたものに新しい番号を振る

使い方: python3 fix_pptx.py 入力.pptx 出力.pptx
"""
import re
import sys
import zipfile


def fix(src, dst):
    zin = zipfile.ZipFile(src)
    names = set(zin.namelist())
    out = zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED)
    report = []
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename == "[Content_Types].xml":
            text = data.decode("utf-8")

            def keep(m):
                part = m.group(1)
                if part.lstrip("/") in names:
                    return m.group(0)
                report.append(f"登録削除: {part}")
                return ""

            text = re.sub(r'<Override PartName="([^"]+)"[^>]*/>', keep, text)
            data = text.encode("utf-8")
        elif re.match(r"ppt/slides/slide\d+\.xml$", item.filename):
            text = data.decode("utf-8")
            ids = [int(i) for i in re.findall(r'<p:cNvPr id="(\d+)"', text)]
            next_id = max(ids, default=0) + 1
            seen = set()

            def renum(m):
                nonlocal next_id
                i = int(m.group(1))
                if i in seen:
                    report.append(f"ID変更: {item.filename} {i}→{next_id}")
                    i = next_id
                    next_id += 1
                seen.add(i)
                return f'<p:cNvPr id="{i}"'

            text = re.sub(r'<p:cNvPr id="(\d+)"', renum, text)
            data = text.encode("utf-8")
        out.writestr(item, data)
    out.close()
    return report


if __name__ == "__main__":
    for line in fix(sys.argv[1], sys.argv[2]):
        print(line)
