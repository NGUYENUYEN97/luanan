import fitz

doc = fitz.open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe\Research_on_Collaborative_Management_Mechanism_of_.pdf")
print("=== Research_on_Collaborative_Management_Mechanism_of_.pdf ===")
for p in range(min(2, len(doc))):
    print(f"--- PAGE {p+1} ---")
    print(doc[p].get_text()[:1500])
