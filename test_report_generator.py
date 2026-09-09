import csv
import tempfile
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path

import report_generator
from openpyxl import Workbook


class ReportGeneratorTest(unittest.TestCase):
    def setUp(self):
        self.root = Path(__file__).resolve().parent
        self.csv_source = self.root / "tests" / "fixtures" / "线路汇总样例.csv"
        self.temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp_dir.cleanup)
        self.xlsx_source = Path(self.temp_dir.name) / "线路汇总样例.xlsx"

        workbook = Workbook(write_only=True)
        worksheet = workbook.create_sheet("线路汇总")
        with self.csv_source.open("r", encoding="utf-8-sig", newline="") as csv_file:
            for row in csv.reader(csv_file):
                worksheet.append(row)
        workbook.save(self.xlsx_source)
        workbook.close()
        self.source = self.xlsx_source

    def test_csv_generates_complete_report(self):
        self.assertTrue(self.csv_source.exists())
        excel_lines, excel_issues = report_generator.load_lines(self.xlsx_source)
        csv_lines, csv_issues = report_generator.load_lines(self.csv_source)

        self.assertGreater(len(csv_lines), 0)
        self.assertEqual(csv_lines, excel_lines)
        self.assertEqual(csv_issues, excel_issues)
        self.assertEqual(
            len(csv_lines), len({line.line_id for line in csv_lines})
        )
        self.assertIsInstance(csv_issues, list)

        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = report_generator.generate_report(
                self.csv_source,
                Path(temp_dir),
                as_of="2026-08-25",
                generated_at=datetime(
                    2026, 8, 25, 10, 30, 0, tzinfo=report_generator.TIMEZONE
                ),
                unique_id="c5a6b7d8",
            )
            html_path = output_dir / "专线运营分析报告_20260825_c5a6b7d8.html"
            markdown_path = output_dir / "专线运营分析报告_20260825_c5a6b7d8.md"
            self.assertTrue(html_path.exists())
            self.assertTrue(markdown_path.exists())

            for content in (
                html_path.read_text(encoding="utf-8"),
                markdown_path.read_text(encoding="utf-8"),
            ):
                self.assertIn(self.csv_source.name, content)
                self.assertIn("线路状态汇总", content)
                self.assertIn("月销金额分析", content)
                self.assertIn("未来6个月月销预测", content)
                self.assertIn("UC 平台地域分析", content)

    def test_fixed_as_of_summary_and_reclaimed_status(self):
        as_of = date(2026, 8, 11)
        lines, issues = report_generator.load_lines(self.source)
        summary = report_generator.summarize(lines, as_of, issues)

        self.assertGreater(len(lines), 0)
        self.assertEqual(summary["line_count"], len(lines))
        self.assertEqual(summary["total_bandwidth"], sum(line.bandwidth for line in lines))
        self.assertAlmostEqual(
            summary["total_settlement"],
            sum(line.settlement_total for line in lines),
        )
        self.assertAlmostEqual(
            sum(row["settlement"] for row in summary["status_rows"]),
            summary["total_settlement"],
        )
        self.assertEqual(
            summary["reclaimed_line_count"],
            sum(line.end_date < as_of for line in lines),
        )
        self.assertEqual(
            summary["active_line_count"],
            sum(line.end_date >= as_of for line in lines),
        )
        self.assertEqual(
            summary["reclaimed_bandwidth"],
            sum(line.bandwidth for line in lines if line.end_date < as_of),
        )
        self.assertEqual(
            len(summary["expiring"]),
            sum(0 <= (line.end_date - as_of).days <= 7 for line in lines),
        )

        monthly = report_generator.creation_metrics(lines, as_of)
        self.assertEqual(sum(row["line_count"] for row in monthly), len(lines))
        self.assertEqual(sum(row["bandwidth"] for row in monthly), summary["total_bandwidth"])
        self.assertAlmostEqual(
            sum(row["settlement"] for row in monthly),
            summary["total_settlement"],
        )
        self.assertEqual(
            sum(row["reclaimed_line_count"] for row in monthly),
            summary["reclaimed_line_count"],
        )
        self.assertEqual(
            sum(row["active_line_count"] for row in monthly),
            summary["active_line_count"],
        )

        active_bandwidth = report_generator.monthly_active_bandwidth_metrics(lines, as_of)
        self.assertEqual(
            active_bandwidth[0]["month"],
            min(line.start_date for line in lines).strftime("%Y-%m"),
        )
        self.assertEqual(active_bandwidth[-1]["month"], as_of.strftime("%Y-%m"))
        for row in active_bandwidth:
            month_start = date.fromisoformat(f'{row["month"]}-01')
            natural_month_end = report_generator.add_months(month_start, 1) - timedelta(days=1)
            snapshot_date = min(natural_month_end, as_of)
            expected_active_lines = [
                line
                for line in lines
                if line.start_date <= snapshot_date <= line.end_date
            ]
            self.assertEqual(row["active_line_count"], len(expected_active_lines))
            self.assertEqual(
                row["active_bandwidth"],
                sum(line.bandwidth for line in expected_active_lines),
            )

        reclaimed = report_generator.monthly_reclaimed_metrics(lines, as_of)
        expected_reclaimed_lines = [line for line in lines if line.end_date < as_of]
        self.assertEqual(
            sum(row["reclaimed_line_count"] for row in reclaimed),
            len(expected_reclaimed_lines),
        )
        self.assertEqual(
            sum(row["reclaimed_bandwidth"] for row in reclaimed),
            sum(line.bandwidth for line in expected_reclaimed_lines),
        )
        self.assertEqual(reclaimed[-1]["month"], as_of.strftime("%Y-%m"))
        for row in reclaimed:
            self.assertEqual(
                row["reclaimed_line_count"],
                sum(
                    line.end_date.strftime("%Y-%m") == row["month"]
                    for line in expected_reclaimed_lines
                ),
            )

        monthly_sales = report_generator.service_monthly_sales_metrics(lines)
        self.assertAlmostEqual(
            sum(row["contract_monthly_sales"] for row in monthly_sales),
            sum(line.sales_total for line in lines),
        )
        self.assertAlmostEqual(
            sum(row["settlement_monthly_sales"] for row in monthly_sales),
            sum(line.settlement_total for line in lines),
        )
        uc_rows = report_generator.platform_region_metrics(lines, "uc", as_of)
        uc_lines = [line for line in lines if line.platform.casefold() == "uc"]
        self.assertTrue(uc_rows)
        self.assertEqual(sum(row["line_count"] for row in uc_rows), len(uc_lines))
        self.assertAlmostEqual(
            sum(row["bandwidth"] for row in uc_rows),
            sum(line.bandwidth for line in uc_lines),
        )
        self.assertAlmostEqual(
            sum(row["sales"] for row in uc_rows),
            sum(line.sales_total for line in uc_lines),
        )
        self.assertAlmostEqual(
            sum(row["settlement"] for row in uc_rows),
            sum(line.settlement_total for line in uc_lines),
        )
        self.assertAlmostEqual(sum(row["bandwidth_share"] for row in uc_rows), 1.0)
        self.assertAlmostEqual(sum(row["sales_share"] for row in uc_rows), 1.0)
        self.assertAlmostEqual(sum(row["settlement_share"] for row in uc_rows), 1.0)
        multi_month_line = next(line for line in lines if line.term_months >= 3)
        monthly_sales_by_month = {row["month"]: row for row in monthly_sales}
        expected_contract_monthly = multi_month_line.sales_total / multi_month_line.term_months
        expected_settlement_monthly = (
            multi_month_line.settlement_total / multi_month_line.term_months
        )
        for offset in range(int(round(multi_month_line.term_months))):
            month = report_generator.add_months(multi_month_line.start_date, offset).strftime(
                "%Y-%m"
            )
            self.assertIn(month, monthly_sales_by_month)
            self.assertGreaterEqual(
                monthly_sales_by_month[month]["contract_monthly_sales"],
                expected_contract_monthly,
            )
            self.assertGreaterEqual(
                monthly_sales_by_month[month]["settlement_monthly_sales"],
                expected_settlement_monthly,
            )

        forecast = report_generator.forecast_monthly_sales_metrics(lines, as_of, horizon=6)
        self.assertEqual(
            [row["month"] for row in forecast],
            ["2026-09", "2026-10", "2026-11", "2026-12", "2027-01", "2027-02"],
        )
        active_lines = [line for line in lines if line.end_date >= as_of]
        self.assertTrue(all(row["line_count"] == len(active_lines) for row in forecast))
        self.assertAlmostEqual(
            sum(row["contract_monthly_sales"] for row in forecast),
            6 * sum(line.sales_total / line.term_months for line in active_lines),
        )
        self.assertAlmostEqual(
            sum(row["settlement_monthly_sales"] for row in forecast),
            6 * sum(line.settlement_total / line.term_months for line in active_lines),
        )
        displayed_monthly_sales = report_generator.service_monthly_sales_metrics(
            lines, through_date=as_of
        )
        self.assertLessEqual(
            max(row["month"] for row in displayed_monthly_sales),
            as_of.strftime("%Y-%m"),
        )

        future_summary = report_generator.summarize(
            lines, max(line.end_date for line in lines), issues
        )
        self.assertEqual(
            future_summary["reclaimed_line_count"],
            sum(line.end_date < max(item.end_date for item in lines) for line in lines),
        )

    def test_output_is_unique_and_contains_status_and_monthly_sections(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_root = Path(temp_dir)
            generated_at = datetime(2026, 8, 11, 15, 30, 0, tzinfo=report_generator.TIMEZONE)
            first = report_generator.generate_report(
                self.source,
                output_root,
                as_of="2026-08-11",
                generated_at=generated_at,
                unique_id="a1b2c3d4",
            )
            second = report_generator.generate_report(
                self.source,
                output_root,
                as_of="2026-08-11",
                generated_at=generated_at,
                unique_id="e5f6a7b8",
            )
            self.assertNotEqual(first, second)
            first_html = first / "专线运营分析报告_20260811_a1b2c3d4.html"
            first_markdown = first / "专线运营分析报告_20260811_a1b2c3d4.md"
            self.assertTrue(first_html.exists())
            self.assertTrue(first_markdown.exists())
            self.assertTrue((second / "专线运营分析报告_20260811_e5f6a7b8.html").exists())

            markdown = first_markdown.read_text(encoding="utf-8")
            html = first_html.read_text(encoding="utf-8")
            self.assertIn("线路状态汇总", markdown)
            self.assertIn("按创建时间的月度新增", markdown)
            self.assertIn("已回收线路", markdown)
            self.assertIn("在售总带宽", markdown)
            lines, _ = report_generator.load_lines(self.source)
            expected_active_bandwidth = sum(
                line.bandwidth for line in lines if line.end_date >= date(2026, 8, 11)
            )
            self.assertIn(
                f"{report_generator.fmt_number(expected_active_bandwidth)} Mbps",
                markdown,
            )
            self.assertIn("实际收入（结算价格）", markdown)
            self.assertIn("月销金额分析", markdown)
            self.assertIn("结算月销（实际收入）", markdown)
            self.assertIn("按线路开始时间所在月份作为第一个服务月", markdown)
            self.assertIn("季度、半年、年付线路", markdown)
            self.assertIn("未来6个月月销预测", markdown)
            self.assertIn("2027-02", markdown)
            self.assertNotIn("2027-03", markdown)
            monthly_section = markdown[
                markdown.index("## 七、月销金额分析") : markdown.index(
                    "## 八、未来6个月月销预测"
                )
            ]
            self.assertIn("2026-08", monthly_section)
            self.assertNotIn("2026-09", monthly_section)
            self.assertNotIn("2027-", monthly_section)
            self.assertNotIn("预计实际收入", markdown)
            self.assertIn("class=\"creation-bars\"", html)
            self.assertIn("## 四、每月在售统计", markdown)
            self.assertLess(
                markdown.index("## 三、7 天内到期线路提醒"),
                markdown.index("## 四、每月在售统计"),
            )
            self.assertLess(
                markdown.index("## 四、每月在售统计"),
                markdown.index("## 五、每月回收线路统计"),
            )
            self.assertLess(
                markdown.index("## 五、每月回收线路统计"),
                markdown.index("## 六、按创建时间的月度新增"),
            )
            self.assertIn("每月在售线路数量", markdown)
            self.assertIn("在售线路数", markdown)
            self.assertIn("class=\"active-statistics-bars\"", html)
            self.assertIn("每月在售线路数量和总带宽柱状图", html)
            self.assertIn("每月回收线路数量", markdown)
            self.assertIn("回收带宽（Mbps）", markdown)
            self.assertIn("class=\"reclaimed-statistics-bars\"", html)
            self.assertIn("每月回收线路数量和回收带宽柱状图", html)
            self.assertIn("按创建月份的已回收线路明细", markdown)
            reclaimed_detail_section = markdown[
                markdown.index("### 按创建月份的已回收线路明细") : markdown.index(
                    "## 七、月销金额分析"
                )
            ]
            for line in lines:
                if line.end_date < date(2026, 8, 11):
                    self.assertIn(line.line_id, reclaimed_detail_section)
                    self.assertIn(line.name, reclaimed_detail_section)
            self.assertIn('class="reclaimed-creation-detail"', html)
            self.assertIn("线路状态汇总", html)
            self.assertIn("在售总带宽", html)
            self.assertIn("实际收入（结算价格）", html)
            self.assertIn("月销金额分析", html)
            self.assertIn("class=\"monthly-sales-bars\"", html)
            self.assertIn("按线路开始时间所在月份作为第一个服务月", html)
            self.assertIn("按服务月份统计合同月销和结算月销柱状图", html)
            self.assertIn("未来6个月月销预测", html)
            self.assertIn("class=\"forecast-sales-bars\"", html)
            self.assertIn("未来6个月合同月销与结算月销预测柱状图", html)
            self.assertIn("UC 平台地域分析", markdown)
            self.assertIn("上海-台北", markdown)
            self.assertIn("UC 平台地域分析", html)
            self.assertIn("上海-台北", html)
            uc_markdown_section = markdown[
                markdown.index("### UC 平台地域分析") : markdown.index("## 十、地域/产品结构分析")
            ]
            uc_html_section = html[
                html.index("UC 平台地域分析") : html.index("十、地域/产品结构分析")
            ]
            self.assertNotIn("毛利率", uc_markdown_section)
            self.assertNotIn("毛利率", uc_html_section)
            self.assertIn('class="uc-region-table"', html)
            self.assertIn("overflow-x: hidden", html)
            self.assertIn("table-layout: fixed", html)
            self.assertIn("table { width: 100%; table-layout: fixed;", html)
            self.assertNotIn("预计实际收入", html)
            self.assertIn("fill: #173b63", html)


if __name__ == "__main__":
    unittest.main()
