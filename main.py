#!/usr/bin/env python3
"""Scrape Longhu market subpages and export to an Excel workbook."""
from __future__ import annotations

import argparse
from typing import Dict, List
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://data.10jqka.com.cn/market/longhu/"
TAB_NAMES = ["全部股票", "机构参与", "散户队上榜", "游资上榜", "跟风高手上榜"]
DEFAULT_OUTPUT = "longhu.xlsx"


def fetch_tab_links(session: requests.Session) -> Dict[str, str]:
    response = session.get(BASE_URL, timeout=15)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    links: Dict[str, str] = {}

    for anchor in soup.find_all("a"):
        text = anchor.get_text(strip=True)
        if text in TAB_NAMES and text not in links:
            href = anchor.get("href")
            if href:
                links[text] = urljoin(BASE_URL, href)

    if "全部股票" not in links:
        links["全部股票"] = BASE_URL

    missing = [name for name in TAB_NAMES if name not in links]
    if missing:
        raise RuntimeError(
            "无法从页面中解析这些栏目链接: " + ", ".join(missing)
        )

    return links


def extract_table(html: str) -> pd.DataFrame:
    tables: List[pd.DataFrame] = pd.read_html(html)
    if not tables:
        raise RuntimeError("未解析到表格数据")
    return tables[0]


def scrape_to_excel(output_path: str) -> None:
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
                "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            ),
            "Referer": BASE_URL,
        }
    )

    tab_links = fetch_tab_links(session)

    with pd.ExcelWriter(output_path) as writer:
        for name in TAB_NAMES:
            url = tab_links[name]
            response = session.get(url, timeout=15)
            response.raise_for_status()
            df = extract_table(response.text)
            df.to_excel(writer, sheet_name=name, index=False)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="抓取同花顺龙虎榜页面子栏目数据并导出 Excel"
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=f"输出 Excel 文件名 (默认: {DEFAULT_OUTPUT})",
    )
    args = parser.parse_args()

    scrape_to_excel(args.output)
    print(f"已导出: {args.output}")


if __name__ == "__main__":
    main()
