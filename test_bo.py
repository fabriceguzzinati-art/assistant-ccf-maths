from utils import CHAPITRES_MATHS_BAC_PRO, PROGRAMME_BO, build_bloc_programme

for niveau in ("2nde Pro", "1ère Pro", "Term Pro"):
    print(f"\n── {niveau} ──")
    for chap in CHAPITRES_MATHS_BAC_PRO[niveau]:
        ok = chap in PROGRAMME_BO.get(niveau, {})
        print("✅" if ok else "❌", chap[:80])

print(build_bloc_programme("Term Pro",
      "Suites géométriques — terme général, sens de variation, somme"))