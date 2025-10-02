import os
import sys
import requests
import subprocess
import json
import yaml
import time

def main():
    try:
        print("开始更新程序...")
        
        # 获取当前程序路径
        current_dir = os.path.dirname(os.path.abspath(__file__))
        config_path = os.path.join(current_dir, "config.yaml")
        
        # 读取当前版本
        with open(config_path, 'r', encoding='utf-8') as f:
            config = yaml.safe_load(f)
            current_version = config.get('version', '1.0')
        
        # 获取最新版本号（旧方式）
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
        print("正在获取更新列表...")
        
        # 获取更新列表
        update_list_url = "http://spj2025.top:19540/apps/update/SPJClassTool/UpdateList.json"
        response = requests.get(update_list_url)
        if response.status_code != 200:
            print(f"无法获取更新列表: HTTP {response.status_code}")
            return
        
        # 解析更新列表
        update_list = json.loads(response.text)
        
        print("正在下载更新文件...")
        
        # 下载并替换所有需要更新的文件
        for filename, file_url in update_list.items():
                
            print(f"下载文件: {filename}")
            
            # 下载文件
            response = requests.get(file_url)
            if response.status_code != 200:
                print(f"下载文件失败: {filename}, HTTP {response.status_code}")
                continue
            
            # 创建文件路径
            file_path = os.path.join(current_dir, filename)
            
            # 保存文件
            with open(file_path, 'wb') as f:
                f.write(response.content)
            
            print(f"文件已更新: {filename}")
        
        # 更新配置文件中的版本号
        config['version'] = latest_version
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.dump(config, f)
        
        print("配置文件更新完成")
        print("所有文件更新完成，正在启动新版本...")
        
        # 启动新版本的主程序
        main_script_path = os.path.join(current_dir, "main.pyw")
        subprocess.Popen(["pythonw", main_script_path])
        
        print("更新成功，程序即将重新启动")
        time.sleep(3000)
        
    except Exception as e:
        print(f"更新过程中出错: {str(e)}")

if __name__ == "__main__":
    main()
