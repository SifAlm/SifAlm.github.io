"""Pipeline orchestration for the Tamam Justlife scraper."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from .config import ScraperConfig
from .excel import ExcelBuilder
from .notes import NotesLogger
from .scraper import JustlifeScraper
from .storage import DataStore


class Pipeline:
    """Execute scraping, validation, and packaging."""

    def __init__(self, config: ScraperConfig) -> None:
        self.config = config
        self.notes = NotesLogger()
        self.datastore = DataStore(config.output_dir, selectors_dir=config.selectors_dir)
        self.scraper = JustlifeScraper(config, self.datastore, self.notes)

    async def run(self) -> dict[str, Path]:
        await self.scraper.run()
        csv_paths = self.datastore.export_csvs()
        self._generate_excel()
        self._write_notes()
        return csv_paths

    # ------------------------------------------------------------------
    def _generate_excel(self) -> Path:
        pages_df = self.datastore.pages_dataframe()
        options_df = self.datastore.options_dataframe()
        prices_df = self.datastore.prices_dataframe()
        addons_df = self.datastore.addons_dataframe()
        availability_df = self.datastore.availability_dataframe()
        offers_df = self.datastore.offers_dataframe()

        summary_df = self._build_summary_sheet(pages_df, prices_df, availability_df)
        price_matrix_df = prices_df.copy()
        addons_sheet_df = addons_df.copy()
        availability_sheet_df = self._build_availability_matrix(availability_df)
        offers_sheet_df = offers_df.copy()
        pages_options_df = self._build_pages_options(pages_df, options_df)

        excel_path = Path(self.config.output_dir) / "Justlife_CRM.xlsx"
        builder = ExcelBuilder(excel_path)
        builder.build_workbook(
            summary_df,
            price_matrix_df,
            addons_sheet_df,
            availability_sheet_df,
            offers_sheet_df,
            pages_options_df,
        )
        return excel_path

    def _write_notes(self) -> Path:
        notes_path = Path(self.config.output_dir) / "NOTES.txt"
        if not self.notes._lines:
            self.notes.add("No notes captured during this run.")
        return self.notes.write(notes_path)

    # ------------------------------------------------------------------
    def _build_summary_sheet(
        self,
        pages_df: pd.DataFrame,
        prices_df: pd.DataFrame,
        availability_df: pd.DataFrame,
    ) -> pd.DataFrame:
        metrics = {
            "Run Date": [pd.Timestamp.utcnow().strftime("%Y-%m-%d %H:%M UTC")],
            "# Pages": [len(pages_df)],
            "# Services": [pages_df[pages_df["level"] == "service"].shape[0] if "level" in pages_df else 0],
            "# Price Combinations": [len(prices_df)],
            "Earliest Lead Time": [self._earliest_lead_time(availability_df)],
            "Average Grand Total": [prices_df["grand_total"].mean() if "grand_total" in prices_df else None],
        }
        if "model_type" in prices_df:
            model_avg = prices_df.groupby("model_type")["grand_total"].mean().dropna()
            for model, value in model_avg.items():
                metrics[f"Avg Total ({model})"] = [value]
        summary_df = pd.DataFrame(metrics)
        return summary_df

    def _earliest_lead_time(self, availability_df: pd.DataFrame) -> str:
        if availability_df.empty:
            return "n/a"
        lead_times = availability_df.get("lead_time_days")
        if lead_times is None or lead_times.empty:
            return "n/a"
        return f"{lead_times.min()} days"

    def _build_availability_matrix(self, availability_df: pd.DataFrame) -> pd.DataFrame:
        if availability_df.empty:
            return pd.DataFrame(columns=["page_id"])
        pivot = (
            availability_df
            .groupby(["page_id", "date_str"]).size()
            .unstack(fill_value=0)
            .reset_index()
        )
        return pivot

    def _build_pages_options(self, pages_df: pd.DataFrame, options_df: pd.DataFrame) -> pd.DataFrame:
        if options_df.empty:
            return pages_df
        merged = pages_df.merge(options_df, on="page_id", how="left")
        return merged

