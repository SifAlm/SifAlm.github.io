"""Entry point for the Tamam Justlife scraper pipeline."""
from __future__ import annotations

import asyncio

from tamam.config import ScraperConfig
from tamam.env import setup_env
from tamam.pipeline import Pipeline


async def async_main() -> None:
    setup_env()
    config = ScraperConfig()
    pipeline = Pipeline(config)
    await pipeline.run()

    pages_df = pipeline.datastore.pages_dataframe()
    prices_df = pipeline.datastore.prices_dataframe()
    availability_df = pipeline.datastore.availability_dataframe()

    print("===== Tamam Justlife KPI Summary =====")
    print(f"Pages discovered: {len(pages_df)}")
    print(f"Services discovered: {pages_df[pages_df['level'] == 'service'].shape[0] if not pages_df.empty else 0}")
    print(f"Price combinations captured: {len(prices_df)}")
    print(f"Availability rows: {len(availability_df)}")
    print("Outputs written to:")
    print("  - Justlife_CRM.xlsx")
    print("  - NOTES.txt")


if __name__ == "__main__":
    asyncio.run(async_main())
