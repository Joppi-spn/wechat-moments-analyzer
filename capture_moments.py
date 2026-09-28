# -*- coding: utf-8 -*-
"""
朋友圈截图采集（ADB + OCR 方案的采集端）

为什么用截图：新版微信屏蔽了 UiAutomator/无障碍的控件读取，
本脚本不读取任何 App 控件，只做「截图 + 模拟上滑」两件事，
文字识别与去重交给 export_excel.py 离线完成。

前置条件：
  1. 安卓手机开启「开发者选项 → USB 调试」并用数据线连接电脑
  2. 在手机微信中手动打开目标好友的朋友圈，停在最顶部
  3. 安装 adb（platform-tools），查找顺序：环境变量 ADB_PATH →
     系统 PATH 中的 adb → 项目内 ../tools/platform-tools/adb.exe

用法：
  python capture_moments.py --name 好友名 --max 100        # 首次采集
  python capture_moments.py --name 好友名 --start 101      # 断点续拍
"""
import os
import time
import argparse
import subprocess

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def load_dotenv():
    """零依赖读取项目根目录下的 .env（若存在），把 KEY=VALUE 注入环境变量"""
    path = os.path.join(BASE_DIR, '.env')
    if not os.path.exists(path):
        return
    with open(path, encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#') or '=' not in line:
                continue
            k, v = line.split('=', 1)
            os.environ.setdefault(k.strip(), v.strip())


load_dotenv()

# 屏幕滑动参数（按自己手机分辨率调整：水平居中、从底部上提到中上部）
SWIPE = (540, 1850, 540, 700, 450)   # x1, y1, x2, y2, 毫秒
SETTLE_SECONDS = 1.6                 # 每次滑动后等待内容加载
REMOTE_SHOT = '/sdcard/_moment_cap.png'
DIFF_THRESHOLD = 2.0                 # 相邻截图 16x16 灰度平均差，小于它视为没滚动
SAME_LIMIT = 2                       # 连续多少次没滚动即判定到底

_ADB_CANDIDATES = [
    os.environ.get('ADB_PATH'),
    'adb',
    os.path.join(BASE_DIR, '..', 'tools', 'platform-tools', 'adb.exe'),
]
ADB = next((p for p in _ADB_CANDIDATES if p), 'adb')


def adb(*args):
    """调用外部 adb 命令"""
    return subprocess.run([ADB, *args], capture_output=True)


def shot(path):
    """手机截屏并拉取到电脑"""
    adb('shell', 'screencap', '-p', REMOTE_SHOT)
    adb('pull', REMOTE_SHOT, path)


def similar(p1, p2):
    """两张图缩成 16x16 灰度马赛克后的平均差值；接近 0 说明页面没动"""
    import cv2
    a = cv2.imread(p1, cv2.IMREAD_GRAYSCALE)
    b = cv2.imread(p2, cv2.IMREAD_GRAYSCALE)
    if a is None or b is None:
        return 999
    a = cv2.resize(a, (16, 16))
    b = cv2.resize(b, (16, 16))
    return float(cv2.absdiff(a, b).mean())


def safe_name(name):
    """去掉 Windows 文件名非法字符"""
    import re
    return re.sub(r'[\\/:*?"<>|]', '', name).strip() or 'friend'


def capture(name, max_shots, start):
    caps_dir = os.path.join(BASE_DIR, 'caps_' + safe_name(name))
    os.makedirs(caps_dir, exist_ok=True)

    tip = '停在朋友圈最顶部，5 秒后开始' if start == 1 else f'从第 {start} 张续拍，3 秒后开始'
    print(f'请确认手机已{tip}，请勿触碰手机……')
    time.sleep(5 if start == 1 else 3)

    prev, same = None, 0
    for i in range(start, start + max_shots):
        p = os.path.join(caps_dir, f'cap_{i:03d}.png')
        shot(p)
        if not os.path.exists(p) or os.path.getsize(p) < 10000:
            print(f'第 {i} 张截图异常，重试')
            time.sleep(1)
            shot(p)

        if prev is not None:
            d = similar(prev, p)
            same = same + 1 if d < DIFF_THRESHOLD else 0
            print(f'[#{i}] 已截图，页面差异={d:.1f}')
            if same >= SAME_LIMIT:
                print('连续多页几乎不变，看起来已到底，停止采集。')
                break
        else:
            print(f'[#{i}] 已截图')
        prev = p

        adb('shell', 'input', 'swipe', *map(str, SWIPE))
        time.sleep(SETTLE_SECONDS)

    print(f'采集结束，截图位于 {caps_dir}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='朋友圈截图采集（ADB）')
    parser.add_argument('--name', required=True, help='好友昵称（仅用于截图目录命名）')
    parser.add_argument('--max', type=int, default=100, help='最多截图数量')
    parser.add_argument('--start', type=int, default=1, help='起始编号，用于断点续拍')
    args = parser.parse_args()
    capture(args.name, args.max, args.start)
