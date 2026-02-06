import os
import subprocess
import winreg

def find_wechat():
    """查找微信安装路径"""
    possible_paths = [
        r"C:\Program Files (x86)\Tencent\WeChat\WeChat.exe",
        r"C:\Program Files\Tencent\WeChat\WeChat.exe",
        r"C:\Users\zoufeng\AppData\Local\Programs\WeChat\WeChat.exe",
        r"C:\Users\zoufeng\AppData\Local\Tencent\WeChat\WeChat.exe"
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    # 通过注册表查找
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\WeChat")
        install_path = winreg.QueryValueEx(key, "InstallLocation")[0]
        wechat_path = os.path.join(install_path, "WeChat.exe")
        if os.path.exists(wechat_path):
            return wechat_path
    except:
        pass
    
    return None

def start_wechat():
    wechat_path = find_wechat()
    
    if wechat_path:
        try:
            subprocess.Popen([wechat_path])
            return f"微信已启动: {wechat_path}"
        except Exception as e:
            return f"启动微信失败: {str(e)}"
    else:
        return "未找到微信，请手动启动"

if __name__ == "__main__":
    result = start_wechat()
    print(result)
    # 复制结果到剪贴板
    subprocess.run(['clip.exe'], input=result.encode('gbk'), shell=True)
