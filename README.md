# \# 朋友圈文本采集与词云分析

# Moments_wordCloud

通过 **ADB 截图 + OCR** 采集指定好友的朋友圈文本，并生成词云图。

## 为什么是截图，而不是直接读控件？

新版微信屏蔽了 UiAutomator / 无障碍服务对其界面控件的读取，
程序化地"直接拿文字"已不可行。本项目因此退回到最朴素也最稳定的方式：

> 模拟人工不断上滑 → 逐屏截图 → 离线 OCR 认字 → 跨截图对齐去重 → 分词生成词云。

全程只依赖正常登录的微信 App，不注入、不 Hook、不抓包，对微信版本不敏感。

## 处理流程

```
手机朋友圈
   │  capture_moments.py   截图 + 模拟上滑（ADB）
   ▼
caps_<昵称>/cap_001.png …
   │  export_excel.py      OCR 认字 + 按段落切分 + 全局去重
   ▼
output/<昵称>朋友圈文本.xlsx   （另附 OCR 明细 sheet 便于核对）
output/<昵称>已处理.txt
   │  wordCloud.py         jieba 分词 + TF-IDF 关键词 + wordcloud 出图
   ▼
output/<昵称>词云.jpg
```

## 环境准备

- Python 3.9+
- 一部安卓手机，开启「开发者选项 → USB 调试」，数据线连接电脑
- 安装 [Android Platform Tools](https://developer.android.com/tools/releases/platform-tools)（提供 `adb`）
- Python 依赖：

```bash
pip install -r requirements.txt
```

`res/` 目录需包含：`cloud.jpg`（词云蒙版，白色为禁区）、`simhei.ttf`（中文字体）、
`stopwords.txt`（停用词表，一行一个）。

## 使用步骤

1. 在手机微信中手动打开目标好友的朋友圈，停在最顶部，保持亮屏。
2. 采集截图（可随时用 `--start` 断点续拍）：

```bash
python capture_moments.py --name 好友昵称 --max 100
```

3. OCR 识别、去重并导出 Excel / 纯文本：

```bash
python export_excel.py --name 好友昵称
```

4. 生成词云：

```bash
python wordCloud.py --name 好友昵称
```

## 配置项

- 复制 `.env.example` 为 `.env`，可自定义 `ADB_PATH`（adb 的完整路径）。
  本项目不使用任何在线 API Key，OCR 为本地离线推理。
- `capture_moments.py` 顶部的 `SWIPE` 坐标按 1080x2400 屏幕标定，
  换分辨率需自行调整；`export_excel.py` 的 `TOP_RESERVED / BODY_X / PARA_GAP`
  是手机截图的版面阈值，同样按设备标定。

## 隐私与合规

- `output/` 与 `caps_*/` 含真实朋友圈内容，**已在 `.gitignore` 中排除，请勿上传**。
- 本项目仅供个人学习与对本人账号数据的研究使用；请遵守《个人信息保护法》
  及微信软件许可协议，勿采集、传播他人隐私，勿用于商业用途。
