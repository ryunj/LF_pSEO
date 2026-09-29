import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DASHBOARD = (ROOT / "dashboard.html").read_text(encoding="utf-8")
APP = (ROOT / "app.py").read_text(encoding="utf-8")


class DashboardContractTests(unittest.TestCase):
    def test_browser_mapping_uses_nbos_registered_name_as_alias(self):
        self.assertIn("매체코드명(NBOS등록)", DASHBOARD)
        self.assertRegex(DASHBOARD, r"aliases\[[^\]]*nbos")
        self.assertRegex(DASHBOARD, r"aliases\[rawCode\]\s*\|\|\s*rawCode")

    def test_browser_retains_sources_for_sequential_uploads(self):
        self.assertIn("const FILE_INPUT=", DASHBOARD)
        self.assertIn("mapByCode", DASHBOARD)
        self.assertRegex(DASHBOARD, r"FILE_INPUT\.(?:perf|map|led)=item\.rows")

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

    def test_browser_accepts_and_parses_excel_workbooks(self):
        self.assertIn('accept=".csv,.tsv,.txt,.xlsx,.xlsm,.xls"', DASHBOARD)
        self.assertIn('<script src="vendor/xlsx.full.min.js"></script>', DASHBOARD)
        self.assertNotIn("cdn.sheetjs.com", DASHBOARD)
        vendor = ROOT / "vendor" / "xlsx.full.min.js"
        self.assertTrue(vendor.exists())
        self.assertGreater(vendor.stat().st_size, 500_000)
        self.assertIn("XLSX.read(await file.arrayBuffer()", DASHBOARD)
        self.assertIn("XLSX.utils.sheet_to_json", DASHBOARD)

    def test_unknown_or_metadata_only_files_are_not_forced_through_perf_dates(self):
        self.assertIn("function detectUploadKind", DASHBOARD)
        self.assertIn("const kind=detectUploadKind(rows)", DASHBOARD)
        self.assertIn("if(!kind){unknown.push", DASHBOARD)
        self.assertNotIn("else{FILE_INPUT.perf=rows", DASHBOARD)

    def test_excel_workbook_prefers_named_business_sheets(self):
        self.assertIn("function sheetPriority", DASHBOARD)
        self.assertIn("kind==='map'&&name==='PSEO'", DASHBOARD)
        self.assertIn("kind==='led'&&name.includes('전체대장')", DASHBOARD)
        self.assertIn("const selected={}", DASHBOARD)

    def test_streamlit_inlines_excel_parser_inside_component(self):
        self.assertIn('read(os.path.join("vendor", "xlsx.full.min.js"))', APP)
        self.assertIn("html = html.replace", APP)
        self.assertIn("<script>%s</script>", APP)


if __name__ == "__main__":
    unittest.main()
