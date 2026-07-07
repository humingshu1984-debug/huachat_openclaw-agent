import asyncio
import yaml

class DataExporter:
    def __init__(self, page, config_path="config/config.yaml"):
        self.page = page
        self.downloaded_files = []
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)

    async def export_all_data(self):
        print("=" * 50)
        print("开始执行数据导出流程...")
        print("=" * 50)

        print("\n等待项目页面完全加载...")
        await self.page.wait_for_load_state('domcontentloaded')
        await self.page.wait_for_timeout(10000)

        try:
            print("\n[步骤1] 点击'项目报告'...")
            try:
                await self.page.wait_for_selector('span.yxt-tabs__item-label', timeout=10000)
                await self.page.click('span.yxt-tabs__item-label:has-text("项目报告")')
            except Exception:
                print("未找到'项目报告'，打印页面元素...")
                await self._print_page_elements()
                raise
            await self.page.wait_for_timeout(3000)

            print("\n[步骤2] 点击右上角'更多'按钮...")
            try:
                await self.page.wait_for_selector('text=更多', timeout=10000)
                more_btn = await self.page.query_selector('button.el-button:has-text("更多")')
                if not more_btn:
                    more_btn = await self.page.query_selector('text=更多')
                await more_btn.click()
            except Exception:
                print("未找到'更多'按钮")
                raise
            await self.page.wait_for_timeout(2000)

            print("\n[步骤3] 点击'旧版报告'...")
            try:
                await self.page.wait_for_selector('text=旧版报告', timeout=10000)
                await self.page.click('text=旧版报告')
            except Exception:
                print("未找到'旧版报告'")
                raise
            await self.page.wait_for_timeout(3000)

            print("\n[步骤4] 点击'导出全部数据'...")
            try:
                await self.page.wait_for_selector('text=导出全部数据', timeout=10000)
                await self.page.click('text=导出全部数据')
            except Exception:
                print("未找到'导出全部数据'")
                raise
            await self.page.wait_for_timeout(3000)

            print("\n[步骤5] 点击'查看'...")
            try:
                await self.page.wait_for_selector('text=查看', timeout=10000)
                await self.page.click('text=查看')
            except Exception:
                print("未找到'查看'")
                raise
            await self.page.wait_for_timeout(5000)

            print("\n[步骤6] 等待文件生成并下载...")
            await self._wait_for_file_ready()

            print("\n" + "=" * 50)
            print("数据导出流程完成！")
            print("=" * 50)
            return True

        except Exception as e:
            print(f"\n导出流程出错: {e}")
            await self.page.screenshot(path='export_error.png')
            return False

    async def _print_page_elements(self):
        print("正在打印页面上的主要元素...")
        elements = await self.page.evaluate('''() => {
            const els = Array.from(document.querySelectorAll('*'))
                .filter(el => el.offsetParent !== null && el.textContent.trim().length > 0 && el.textContent.trim().length < 50)
                .map(el => ({
                    tag: el.tagName,
                    text: el.textContent.trim().substring(0, 50)
                }))
                .slice(0, 50);
            return els;
        }''')
        print(f"页面元素: {elements}")

    async def _wait_for_file_ready(self, max_wait_minutes=10):
        max_attempts = max_wait_minutes

        for attempt in range(1, max_attempts + 1):
            print(f"\n[检查 {attempt}/{max_attempts}] 检查文件状态...")

            page_content = await self.page.content()
            if "文件正在生成中" in page_content:
                print("文件仍在生成中，等待60秒...")
                await asyncio.sleep(60)
                await self.page.reload()
                await self.page.wait_for_timeout(3000)
            else:
                print("文件已生成！正在查找下载按钮...")
                download_btn = await self.page.query_selector('button:has-text("下载")')
                if download_btn:
                    print("找到下载按钮！点击下载...")
                    await download_btn.click()
                    await self.page.wait_for_timeout(3000)
                    return True
                else:
                    print("未找到下载按钮，继续等待...")
                    await asyncio.sleep(60)
                    await self.page.reload()

        print("等待超时，文件仍未生成完成")
        return False

    async def extract_table_data(self):
        print("正在提取页面表格数据...")
        await self.page.wait_for_load_state('domcontentloaded')
        await self.page.wait_for_timeout(5000)

        try:
            table = await self.page.query_selector('.el-table')
            if not table:
                print("未找到表格，尝试其他选择器...")
                table = await self.page.query_selector('table')

            if not table:
                print("仍然未找到表格")
                return None

            rows = await table.query_selector_all('tr')
            data = []
            for row in rows:
                cells = await row.query_selector_all('td')
                row_data = [await cell.inner_text() for cell in cells]
                if row_data:
                    data.append(row_data)

            print(f"成功提取 {len(data)} 行数据")
            return data
        except Exception as e:
            print(f"提取表格数据出错: {e}")
            return None
