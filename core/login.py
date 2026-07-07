import asyncio
import yaml
import os

class LoginHandler:
    def __init__(self, page, config_path="config/config.yaml"):
        self.page = page
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)

    async def login(self):
        """执行登录逻辑"""
        url = self.config['website']['login_url']
        username = self.config['auth']['username']
        password = self.config['auth']['password']

        print(f"正在访问登录页面: {url}")
        try:
            await self.page.goto(url, timeout=60000)
            await self.page.wait_for_load_state('domcontentloaded')
            await self.page.wait_for_timeout(10000) # 增加等待时间到10秒
        except Exception as e:
            print(f"访问页面出错: {e}")
            await self.page.screenshot(path='login_page_error.png')
            return False

        print(f"正在尝试定位账号输入框...")
        
        # 检测并打印所有 iframe
        iframes = self.page.frames
        print(f"检测到 {len(iframes)} 个 frame")
        
        target_frame = self.page
        for i, frame in enumerate(iframes):
            print(f"正在检查 Frame {i}: {frame.url}")
            try:
                # 在每个 frame 中寻找账号框
                selectors = ['#username', 'input[id="username"]', 'input[placeholder*="账号"]', '.el-input__inner']
                for selector in selectors:
                    elem = await frame.query_selector(selector)
                    if elem and await elem.is_visible():
                        print(f"在 Frame {i} 中使用选择器 '{selector}' 找到了账号输入框")
                        target_frame = frame
                        username_field = selector
                        break
                if target_frame != self.page:
                    break
            except:
                continue

        if target_frame == self.page:
            # 如果在 frame 中没找到，再在主页面找一次
            selectors = ['#username', 'input[id="username"]', 'input[placeholder*="账号"]', '.el-input__inner']
            username_field = None
            for selector in selectors:
                try:
                    elem = await self.page.query_selector(selector)
                    if elem and await elem.is_visible():
                        username_field = selector
                        print(f"在主页面中使用选择器 '{selector}' 找到了账号输入框")
                        break
                except:
                    continue
            
            if not username_field:
                print("未能在任何地方找到账号输入框，保存截图...")
                await self.page.screenshot(path='login_failed_search.png')
                return False
        
        try:
            # 输入账号
            await target_frame.fill(username_field, username)
            
            # 定位并输入密码
            password_selectors = ['#password', 'input[id="password"]', 'input[type="password"]']
            password_field = None
            for selector in password_selectors:
                try:
                    elem = await target_frame.query_selector(selector)
                    if elem and await elem.is_visible():
                        password_field = selector
                        break
                except:
                    continue
            
            if password_field:
                await target_frame.fill(password_field, password)
            else:
                print("未找到密码输入框")
                return False

            # 点击登录
            await target_frame.click('button:has-text("登录")')
            
            print("已提交登录，等待跳转...")
            await self.page.wait_for_timeout(5000)
            return True

        except Exception as e:
            print(f"执行登录动作出错: {e}")
            return False
