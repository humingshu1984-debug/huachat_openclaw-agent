import asyncio
import yaml
import os
import json

class PageNavigator:
    def __init__(self, page, context, config_path="config/config.yaml"):
        self.page = page
        self.context = context
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)

    async def navigate_to_training_center(self):
        print("正在等待页面加载完成...")
        await self.page.wait_for_load_state('domcontentloaded')
        print("额外等待15秒让页面完全加载...")
        await self.page.wait_for_timeout(15000)

        print("正在查找并点击培训中心图标...")

        try:
            # 优先使用用户提供的精确选择器
            selectors = [
                'div.core-app-car-title:has-text("培训中心")',
                'div.yxt-tooltip:has-text("培训中心")',
                'div.core-app-car-title',
                'div.yxt-tooltip',
                '[class*="core-app-car-title"]',
                'div:has-text("培训中心")',
                'text=培训中心',
            ]

            training_center_elem = None
            for selector in selectors:
                try:
                    elem = await self.page.query_selector(selector)
                    if elem and await elem.is_visible():
                        training_center_elem = elem
                        print(f"使用选择器 '{selector}' 找到了培训中心元素")
                        break
                except:
                    continue

            if training_center_elem:
                is_visible = await training_center_elem.is_visible()
                print(f"培训中心元素可见状态: {is_visible}")

                print("点击培训中心，将打开新页面...")
                async with self.context.expect_page() as new_page_info:
                    await training_center_elem.click()

                new_page = await new_page_info.value
                print("已检测到新页面，等待新页面加载...")
                await new_page.wait_for_load_state('domcontentloaded')
                await new_page.wait_for_timeout(5000)

                self.page = new_page
                print(f"新页面URL: {self.page.url}")
                print("已进入培训中心（新页面）")
            else:
                print("未找到培训中心元素，打印所有包含'培训'的文字元素...")

                training_elements = await self.page.evaluate('''() => {
                    const elements = Array.from(document.querySelectorAll('*'))
                        .filter(el => el.textContent.includes('培训'))
                        .map(el => ({
                            tag: el.tagName,
                            text: el.textContent.trim().substring(0, 80),
                            className: el.className
                        }))
                        .slice(0, 20);
                    return elements;
                }''')
                # 修复 Python 里的 NameError，使用 json.dumps
                print(f"包含'培训'的元素: {json.dumps(training_elements, ensure_ascii=False, indent=2)}")

                print("使用JavaScript方式查找并点击培训中心...")
                clicked = await self.page.evaluate('''() => {
                    const elements = Array.from(document.querySelectorAll('div'));
                    const target = elements.find(el =>
                        el.textContent.trim().includes('培训中心') && (
                            el.className.includes('core-app-car-title') ||
                            el.className.includes('yxt-tooltip')
                        )
                    );
                    if (target) {
                        target.click();
                        return true;
                    }
                    return false;
                }''')

                if clicked:
                    print("使用JS方式点击了培训中心...")
                    await self.page.wait_for_timeout(3000)

                    all_pages = self.context.pages
                    print(f"当前所有页面数量: {len(all_pages)}")
                    if len(all_pages) > 1:
                        self.page = all_pages[-1]
                        print(f"新页面URL: {self.page.url}")
                        print("已进入培训中心（新页面）")
                    else:
                        print("点击可能没有打开新页面")
                else:
                    print("JS方式也未找到培训中心元素")
                    await self.page.screenshot(path='homepage_screenshot.png')
                    print("已保存截图: homepage_screenshot.png")
        except Exception as e:
            print(f"点击培训中心时出错: {e}")
            await self.page.screenshot(path='error_screenshot.png')
            raise

    async def find_and_enter_project(self, project_name=None):
        if project_name is None:
            project_name = self.config.get('target_project', {}).get('name', '')

        print(f"正在查找培训项目: {project_name}")
        # 增加等待时间
        await self.page.wait_for_load_state('domcontentloaded')
        await self.page.wait_for_timeout(10000)

        try:
            search_input = await self.page.query_selector('input[placeholder*="搜索"]')
            if search_input:
                print("找到搜索框，正在输入项目名称...")
                await search_input.fill(project_name)
                await self.page.wait_for_timeout(3000)
            else:
                print("未找到搜索框，打印页面上的可见元素...")

                visible_elements = await self.page.evaluate('''() => {
                    const elements = Array.from(document.querySelectorAll('*'))
                        .filter(el => el.offsetParent !== null && el.textContent.trim().length > 0)
                        .map(el => ({
                            tag: el.tagName,
                            text: el.textContent.trim().substring(0, 50),
                            class: typeof el.className === 'string' ? el.className.substring(0, 30) : ''
                        }))
                        .filter(el => el.text.length > 0)
                        .slice(0, 30);
                    return elements;
                }''')
                print(f"页面上的可见元素: {json.dumps(visible_elements, ensure_ascii=False, indent=2)}")

                await self.page.wait_for_timeout(5000)

        except Exception as e:
            print(f"查找项目时出错: {e}")

        print("正在尝试点击项目...")
        project_item = await self.page.query_selector(f'text={project_name}')
        if project_item:
            await project_item.click()
            await self.page.wait_for_timeout(5000)
            print(f"已进入项目: {project_name}")
            return True
        else:
            print("直接匹配未找到，尝试模糊搜索...")
            # 尝试只匹配核心关键词
            partial_name = "极客行动"
            fuzzy_item = await self.page.query_selector(f'text={partial_name}')
            if fuzzy_item:
                print(f"使用部分名称 '{partial_name}' 找到项目，正在点击...")
                await fuzzy_item.click()
                await self.page.wait_for_timeout(5000)
                print(f"已进入项目: {partial_name}")
                return True

            print(f"未找到项目: {project_name}")
            await self.page.screenshot(path='project_not_found.png')
            return False

    async def find_export_button_and_click(self):
        print("正在查找导出按钮...")
        export_btn = await self.page.query_selector('button:has-text("导出")')
        if export_btn:
            await export_btn.click()
            await self.page.wait_for_timeout(2000)
            print("已点击导出按钮")
            return True
        return False

    async def wait_for_table_and_extract(self):
        print("正在等待数据表格加载...")
        await self.page.wait_for_selector('.el-table', timeout=30000)
        print("表格已加载")
