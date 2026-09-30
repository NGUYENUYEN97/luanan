import fitz, os, sys
sys.stdout.reconfigure(encoding='utf-8')

base04 = r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\04_BaiBao_TrongNuoc'
base06 = r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\06_TaiLieu_BaiBao_DeTai'

# SHTT-related files we haven't fully read yet
targets = [
    (base04, 'CVv358S102020019.pdf'),      # Tạp chí Nghề Luật older
    (base04, 'CVv266S01A2024110.pdf'),
    (base04, 'CVv266S06A2024144.pdf'),
    (base04, 'CVv266S2622024093.pdf'),
    (base04, 'CVv266S2962025005.pdf'),
    (base04, 'CVv266S4A2025009.pdf'),
    (base04, 'CVv266S6A22025050.pdf'),
    (base04, 'CVv146S132019032.pdf'),
    (base04, 'CVv146S162024068.pdf'),
    (base04, 'CVv146S242024036.pdf'),
    (base04, 'CVv146S82024008.pdf'),
    (base04, 'CVv328V1S062025345.pdf'),
    (base04, 'CVv39S202024007.pdf'),
    (base04, '194503-f64d53aede71810a9a41720314c90d19.pdf'),
    (base04, '2132-Article Text-10440-2-10-20250611.pdf'),
    (base04, 'CTv184V38S32022041.pdf'),
    (base04, 'CVv133S8232024031.pdf'),
    (base04, 'CVv133S8332024010.pdf'),
    (base04, 'CVv133S8532025064.pdf'),
    (base04, 'CVv168S6902025011.pdf'),
    (base04, '13293-33808-1-PB.pdf'),
    (base04, '987acecf5409fac6f7973d5a5b1a91a99575.pdf'),
    (base04, '5-2-25-938.pdf'),
    (base06, 'Tuyen-tap.pdf'),
    (base06, 'CVv266S2992025084.pdf'),
    (base06, 'CVv133S8212024121.pdf'),
    (base06, 'CVv133S8332024031.pdf'),
    (base06, 'CVv39S212024027.pdf'),
    (base06, '287513-24a232e93919f6ddbe9662c6893f0bf1.pdf'),
    (base06, '324763-37f8b391cf6fae7c4af8859f1d533895.pdf'),
]

output = r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\extra_shtt_articles.txt'
results = []
ok = 0

for folder, fname in targets:
    fp = os.path.join(folder, fname)
    if not os.path.exists(fp):
        continue
    try:
        doc = fitz.open(fp)
        # Extract first 3 pages for screening
        text = ""
        for i, page in enumerate(doc):
            if i < 3:
                text += page.get_text() + "\n"
        doc.close()
        # Check if SHTT-related
        lower = text.lower()
        if any(kw in lower for kw in ['sở hữu trí tuệ', 'shtt', 'bản quyền', 'nhãn hiệu', 'sáng chế', 'quyền tác giả', 'intellectual property']):
            results.append(f"\n{'='*80}\nFILE: {fname}\nFOLDER: {os.path.basename(folder)}\nPAGES: {doc.page_count}\n{'='*80}\n{text}\n")
            ok += 1
            print(f'SHTT: {fname}')
        else:
            print(f'SKIP: {fname} (not SHTT)')
    except Exception as e:
        print(f'ERR: {fname}: {e}')

with open(output, 'w', encoding='utf-8') as f:
    for r in results:
        f.write(r)
print(f'\nDone. SHTT articles: {ok}, Output: {os.path.getsize(output)//1024}KB')
