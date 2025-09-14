import os
import sys
import requests
import subprocess
import time
import yaml

def main():
    try:
        print("开始更新程序...")
        
        # 获取当前程序路径
        current_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(current_dir, "config.yaml")
        main_script_path = os.path.join(current_dir, "main.pyw")
        temp_script_path = os.path.join(current_dir, "main_temp.pyw")
        
        # 读取当前版本
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            current_version = config.get('version', '1.0')
        
        # 获取最新版本号
        version_url = "http://spj2025.top:19540/apps/version/pyclasstool.txt"
        response = requests.get(version_url)
        if response.status_code != 200:
            print(f"无法获取版本信息: HTTP {response.status_code}")
            return
        
        # 使用GBK编码解码版本号
        latest_version = response.content.decode('gbk').strip()
        
        # 检查是否有新版本
        if latest_version == current_version:
            print("已是最新版本，无需更新")
            return
        
        print(f"发现新版本: {latest_version}")
        print("正在下载更新...")
        
        # 下载新版本
        download_url = "http://spj2025.top:19540/apps/download/main.pyw"
        response = requests.get(download_url)
        if response.status_code != 200:
            print(f"下载更新失败: HTTP {response.status_code}")
            return
        
        # 保存新版本到临时文件
        with open(temp_script_path, 'wb') as f:
            f.write(response.content)
        
        print("更新下载完成")
        
        # 更新配置文件中的版本号
        config['version'] = latest_version
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config, f)
        
        print("配置文件更新完成")
        
        # 关闭当前程序
        print("正在关闭当前程序...")
        # 这里需要根据实际情况关闭程序
        # 在Windows上，我们可以尝试通过进程名关闭
        # 但更可靠的方式是在主程序中提供关闭接口
        # 这里我们假设主程序会在更新前关闭
        
        # 替换主程序文件
        print("正在替换程序文件...")
        if os.path.exists(main_script_path):
            os.remove(main_script_path)
        os.rename(temp_script_path, main_script_path)
        
        print("更新完成，正在启动新版本...")
        
        # 启动新版本
        subprocess.Popen(["pythonw", main_script_path])
        
        print("更新成功，程序已重新启动")
        
    except Exception as e:
        print(f"更新过程中出错: {str(e)}")
        # 清理临时文件
        if os.path.exists(temp_script_path):
            os.remove(temp_script_path)

if __name__ == "__main__":
    main()