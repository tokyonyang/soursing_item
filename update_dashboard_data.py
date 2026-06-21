from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path


REPORT_DIR = Path("reports")
DASHBOARD_DIR = Path("dashboard_data")


def latest(pattern: str) -> Path | None:
    files = sorted(REPORT_DIR.glob(pattern), reverse=True)
    return files[0] if files else None


def copy_if_exists(src: Path | None, dst: Path) -> bool:
    if not src or not src.exists():
        return False
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return True


def main() -> None:
    DASHBOARD_DIR.mkdir(parents=True, exist_ok=True)

    latest_xlsx = latest("sourcing_rank_*.xlsx")
    latest_product_csv = latest("product_keywords_*.csv")
    latest_discovered_csv = latest("discovered_keywords_*.csv")
    latest_history_csv = latest("history_*.csv")

    copied = {
        "latest_sourcing_rank.xlsx": copy_if_exists(latest_xlsx, DASHBOARD_DIR / "latest_sourcing_rank.xlsx"),
        "latest_product_keywords.csv": copy_if_exists(latest_product_csv, DASHBOARD_DIR / "latest_product_keywords.csv"),
        "latest_discovered_keywords.csv": copy_if_exists(latest_discovered_csv, DASHBOARD_DIR / "latest_discovered_keywords.csv"),
        "latest_history.csv": copy_if_exists(latest_history_csv, DASHBOARD_DIR / "latest_history.csv"),
    }

    meta = {
        "updated_at": datetime.now().isoformat(timespec="seconds"),
        "source_files": {
            "sourcing_rank": str(latest_xlsx) if latest_xlsx else None,
            "product_keywords": str(latest_product_csv) if latest_product_csv else None,
            "discovered_keywords": str(latest_discovered_csv) if latest_discovered_csv else None,
            "history": str(latest_history_csv) if latest_history_csv else None,
        },
        "copied": copied,
    }

    (DASHBOARD_DIR / "latest_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print("Dashboard data prepared:")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
