import sys
import os
import io
from datetime import datetime
from pathlib import Path

# Windows CMD 預設編碼可能不支援中文，強制改為 utf-8 輸出
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd


def parse_dates(date_str):
    """將 '115/01/01,115/01/02' 轉成 [1150101, 1150102]"""
    parts = str(date_str).split(',')
    result = []
    for d in parts:
        d = d.strip()
        if d and d != 'nan':
            result.append(int(d.replace('/', '')))
    return result


def process_a_sheet(df_a, df_b):
    """
    處理 A碼項目清冊，為每筆加上補助款欄並展開子行。

    規則：
    - 小計一半（補助款）= 小計 / 2（公司抽成後剩下的才算居服員的）
    - 若該個案在該日期範圍只有一位居服員 → 單行，填入補助款
    - 若有多位居服員 → 主行（保留小計）+ 子行（每位居服員一行）
      - 子行金額 = 天數 × 給付價格 ÷ 當天人數 ÷ 2
    """

    # 建立 B碼 每天每個案的居服員查找表
    # key: (個案姓名, 日期) → set(居服員)
    b_index = (
        df_b.groupby(['個案姓名', '服務日期(請輸入7碼)'])['居服員姓名']
        .apply(set)
        .to_dict()
    )

    output_rows = []

    for _, row in df_a.iterrows():
        dates = parse_dates(row['服務日期'])
        if not dates:
            output_rows.append(row.to_dict())
            continue

        給付價格 = float(row['給付價格'])
        小計 = float(row['小計'])
        個案 = row['個案姓名']

        # 以 A碼「服務人員」欄決定有哪些人（格式：「甲、乙、丙」）
        listed_workers = [w.strip() for w in str(row['服務人員']).split('、') if w.strip()]

        if len(listed_workers) <= 1:
            # 單一居服員 → 直接加補助款，不產生子行
            r = row.to_dict()
            r['補助款'] = 小計 / 2
            output_rows.append(r)
            continue

        # 多位居服員 → 查 B碼，只保留 A碼列出的人
        listed_set = set(listed_workers)
        daily_workers = {}
        for d in dates:
            all_w = b_index.get((個案, d), set())
            matched = all_w & listed_set
            if matched:
                daily_workers[d] = frozenset(matched)

        # 找出 A碼列出但 B碼完全沒有紀錄的人
        found_workers = set(w for ws in daily_workers.values() for w in ws)
        missing_workers = listed_set - found_workers

        if missing_workers:
            # 有人在 B碼找不到 → 無法自動計算，發出警告並預留位置
            warn_msg = f'[待確認] 序號{row["序號"]} {row["服務代碼"]} {個案}：{sorted(missing_workers)} 在B碼無對應紀錄，請人工核對'
            print(f'  [!] {warn_msg}')
            main = row.to_dict()
            main['補助款'] = None
            main['備註'] = warn_msg
            output_rows.append(main)
            # 為每位列出的居服員預留空白子行
            for w in sorted(listed_set):
                sub = row.to_dict()
                sub['數量'] = None
                sub['小計'] = None
                sub['補助款'] = None
                sub['服務日期'] = None
                sub['服務人員'] = w
                sub['備註'] = '待確認'
                output_rows.append(sub)
            continue

        if not daily_workers:
            # B碼完全找不到該個案的紀錄 → 同上
            warn_msg = f'[待確認] 序號{row["序號"]} {row["服務代碼"]} {個案}：B碼無任何對應紀錄，請人工核對'
            print(f'  [!] {warn_msg}')
            r = row.to_dict()
            r['補助款'] = None
            r['備註'] = warn_msg
            output_rows.append(r)
            continue

        # 正常計算：主行（保留小計，補助款空白）
        main = row.to_dict()
        main['補助款'] = None
        output_rows.append(main)

        # 依「當天服務人員組合」分群，產生子行
        groups = {}
        for d, ws in daily_workers.items():
            groups.setdefault(ws, []).append(d)

        for group_workers, group_dates in sorted(groups.items(), key=lambda x: min(x[1])):
            workers = sorted(group_workers)
            days = sorted(group_dates)
            n = len(days)
            amount = n * 給付價格 / len(workers) / 2

            for w in workers:
                sub = row.to_dict()
                sub['數量'] = n
                sub['小計'] = None
                sub['補助款'] = amount
                sub['服務日期'] = '、'.join(str(d) for d in days)
                sub['服務人員'] = w
                output_rows.append(sub)

    # 組成 DataFrame，插入「補助款」欄在「小計」後面
    df_out = pd.DataFrame(output_rows)
    cols = list(df_a.columns)
    idx = cols.index('小計')
    new_cols = cols[:idx+1] + ['補助款'] + cols[idx+1:]
    df_out = df_out.reindex(columns=new_cols)

    return df_out


def find_input_file(folder: Path) -> Path:
    """掃描資料夾內的 xlsx，排除 output 子資料夾，讓使用者選擇。"""
    xlsx_files = [
        f for f in folder.glob('*.xlsx')
        if not f.name.startswith('~$')  # 排除 Excel 暫存鎖定檔
    ]

    if not xlsx_files:
        print('找不到任何 .xlsx 檔案，請將原始檔放在同一資料夾後重新執行。')
        sys.exit(1)

    if len(xlsx_files) == 1:
        print(f'找到檔案：{xlsx_files[0].name}')
        return xlsx_files[0]

    print('找到多個 .xlsx 檔案，請選擇要處理的原始檔：')
    for i, f in enumerate(xlsx_files, 1):
        print(f'  {i}. {f.name}')
    while True:
        choice = input('請輸入編號：').strip()
        if choice.isdigit() and 1 <= int(choice) <= len(xlsx_files):
            return xlsx_files[int(choice) - 1]
        print('  輸入有誤，請重新輸入。')


def main():
    # PyInstaller 打包後 __file__ 會指向暫存目錄，要改用 sys.executable
    if getattr(sys, 'frozen', False):
        folder = Path(sys.executable).parent
    else:
        folder = Path(__file__).parent

    # 決定輸入檔
    if len(sys.argv) > 1:
        input_path = Path(sys.argv[1])
    else:
        input_path = find_input_file(folder)

    # 輸出資料夾：同層的 output/
    output_dir = folder / 'output'
    output_dir.mkdir(exist_ok=True)

    stem = input_path.stem
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    output_path = output_dir / f'{stem}_完成_{timestamp}.xlsx'

    print(f'讀取：{input_path.name}')
    df_b = pd.read_excel(input_path, sheet_name='中央服務紀錄(B碼)+姓名')

    # A碼可能有 1~3 行說明文字在最上面，自動偵測真正的標題列
    df_a_raw = pd.read_excel(input_path, sheet_name='A碼項目清冊', header=None)
    header_row = next(
        i for i, row in df_a_raw.iterrows()
        if row.astype(str).str.contains('服務日期').any()
    )
    df_a = pd.read_excel(input_path, sheet_name='A碼項目清冊', header=header_row)

    # 給付價格可能是 '770/925' 格式，取斜線前的數字
    df_a['給付價格'] = df_a['給付價格'].astype(str).str.split('/').str[0]

    print(f'A碼筆數：{len(df_a)}，B碼筆數：{len(df_b)}')

    df_a_out = process_a_sheet(df_a, df_b)

    print(f'處理後 A碼筆數：{len(df_a_out)}（含子行）')

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_b.to_excel(writer, sheet_name='中央服務紀錄(B碼)+姓名', index=False)
        df_a_out.to_excel(writer, sheet_name='A碼項目清冊', index=False)

    print(f'\n輸出完成：output/{output_path.name}')
    input('\n按 Enter 鍵關閉...')


if __name__ == '__main__':
    main()
