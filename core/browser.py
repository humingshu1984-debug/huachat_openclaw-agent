import asyncio
from playwright.async_api import async_playwright
import yaml
import os

class BrowserManager:
    def __init__(self, config_path="config/config.yaml"):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None

    async def start(self, headless=False):
        """启动浏览器"""
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(headless=headless)
        self.context = await self.browser.new_context()
        self.page = await self.context.new_page()
        return self.page

    async def stop(self):
        """关闭浏览器"""
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

    async def screenshot(self, path="screenshot.png"):
        """截图保存"""
        await self.page.screenshot(path=path)
