import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DASHBOARD = (ROOT / "dashboard.html").read_text(encoding="utf-8")


class DashboardContractTests(unittest.TestCase):
    def test_browser_mapping_uses_nbos_registered_name_as_alias(self):
        self.assertIn("매체코드명(NBOS등록)", DASHBOARD)
        self.assertRegex(DASHBOARD, r"aliases\[[^\]]*nbos")
        self.assertRegex(DASHBOARD, r"aliases\[rawCode\]\s*\|\|\s*rawCode")

    def test_browser_retains_sources_for_sequential_uploads(self):
        self.assertIn("const FILE_INPUT=", DASHBOARD)
        self.assertIn("mapByCode", DASHBOARD)
        self.assertRegex(DASHBOARD, r"FILE_INPUT\.(?:perf|map|led)=rows")

    def test_flow_chart_keeps_a_readable_responsive_ratio(self):
        chart_grid = re.search(r"\.charts\.two\{([^}]+)\}", DASHBOARD)
        self.assertIsNotNone(chart_grid)
        self.assertIn("grid-template-columns:repeat(3,minmax(0,1fr))", chart_grid.group(1))

        chart_rule = re.search(r"\.cc-b\{([^}]+)\}", DASHBOARD)
        self.assertIsNotNone(chart_rule)
        self.assertIn("aspect-ratio", chart_rule.group(1))
        self.assertIn("min-height", chart_rule.group(1))
        self.assertIn("max-height:260px", chart_rule.group(1))
        self.assertNotIn('preserveAspectRatio="none"', DASHBOARD)

    def test_list_charts_use_exactly_three_equal_columns(self):
        chart_grid = re.search(r"#listCharts\{([^}]+)\}", DASHBOARD)
        self.assertIsNotNone(chart_grid)
        self.assertIn(
            "grid-template-columns:repeat(3,minmax(0,1fr))",
            chart_grid.group(1),
        )

    def test_browser_restores_last_successful_upload_from_indexeddb(self):
        self.assertIn("indexedDB.open", DASHBOARD)
        self.assertIn("saveBackup(D)", DASHBOARD)
        self.assertRegex(DASHBOARD, r"async function restoreLastBackup")
        self.assertRegex(DASHBOARD, r"await loadBackup\(\)")

    def test_week_labels_use_month_and_week_of_month(self):
        self.assertIn("function weekLabel", DASHBOARD)
        self.assertIn("+'월 '+week+'주차'", DASHBOARD)
        self.assertNotIn("+'년 '+Number(id.slice(6))+'주'", DASHBOARD)


if __name__ == "__main__":
    unittest.main()
