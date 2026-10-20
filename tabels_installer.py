from pathlib import Path
import textwrap, zipfile, os

from pathlib import Path
import argparse
import time
import requests
from io import BytesIO
import pandas as pd

BASE_URLS = [
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_RF_Continente_{}.xlsx",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_Ret_IRS_{}_Continente.xlsx",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_RF_{}_Continente.xlsx",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_continente {}_v1.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_continente{}.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_{}_continente.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/TAB{}_CONTINENTE.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/TABELAS_{}-CONTINENTE.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/TABELAS{}.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_continente_{}.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_continente {}.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/TRF_{}_continente.xlsx",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/TRF_{}_continente.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/TRF{}.xls",
    "https://info.portaldasfinancas.gov.pt/pt/apoio_contribuinte/tabela_ret_doclib/Documents/Tabelas_Ret_IRS_{}.xlsx"
]

out_dir = Path(__file__).resolve().parent / "tabels"
out_dir.mkdir(parents=True, exist_ok=True)

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; ECO-local-downloader/1.0)"
})

failed = []
success = []
for year in range(2004, 2026 + 1):
    path = out_dir / f"Tabelas_RF_Continente_{year}.xlsx"

    if path.exists():
        print(f"[SKIP] {year}: already exists")
        continue

    for url in BASE_URLS:
        url = url.format(year)

        try:
            response = session.get(url, timeout=30)
            response.raise_for_status()
            content = response.content

            if content.startswith(b"PK"):
                # Already an XLSX file
                path = out_dir / f"Tabelas_RF_Continente_{year}.xlsx"
                path.write_bytes(content)

            elif content.startswith(b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"):
                # Convert XLS to XLSX
                path = out_dir / f"Tabelas_RF_Continente_{year}.xlsx"

                sheets = pd.read_excel(
                    BytesIO(content),
                    sheet_name=None,
                    header=None,
                    dtype=object,
                    engine="xlrd",
                )

                with pd.ExcelWriter(path, engine="openpyxl") as writer:
                    for sheet_name, df in sheets.items():
                        df.to_excel(
                            writer,
                           sheet_name=sheet_name[:31],
                            index=False,
                            header=False,
                        )

            else:
                raise ValueError("Response is not a valid XLS or XLSX file")

            print(f"[OK] {year}: saved {path.name}")
            success.append(year)
            break

        except Exception as exc:
            print(f"[FAIL] {year}: {exc}")


    if success.pop != year:
        failed.append(year)

    time.sleep(0.25)

print(f"\nFinished. Files are in: {out_dir}")
if failed:
    print("Years that failed:", ", ".join(map(str, failed)))
    print("Try again later or check the source URL for those years.")
