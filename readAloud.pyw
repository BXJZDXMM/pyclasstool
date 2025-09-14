import sys
import win32com.client

def speak(text):
    """使用Windows语音API朗读文本"""
    try:
        # 创建语音合成对象
        speaker = win32com.client.Dispatch("SAPI.SpVoice")
        # 朗读文本
        speaker.Speak(text)
    except Exception as e:
        print(f"朗读失败: {e}")

if __name__ == "__main__":
    # 检查命令行参数
    if len(sys.argv) < 2:
        print("用法: python readAloud.py \"要朗读的文本\"")
        sys.exit(1)
    
    # 获取要朗读的文本
    text_to_speak = " ".join(sys.argv[1:])
    
    # 朗读文本
    speak(text_to_speak)