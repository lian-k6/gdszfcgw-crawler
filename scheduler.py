"""
定时任务调度器 - 每天定时运行政府采购爬虫并发送通知

使用方式：
1. 直接运行: python scheduler.py
2. 后台运行（Windows）: pythonw scheduler.py
3. 或添加到Windows任务计划程序定时执行 run_daily.bat
"""

import logging
import time
import sys
import os

# 确保能导入同级模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import schedule
except ImportError:
    print("正在安装 schedule 库...")
    os.system(f"{sys.executable} -m pip install schedule -q")
    import schedule

from datetime import datetime
from gdszfcgw import crawl
from notifier import notify_all

try:
    from config import CRAWL_CONFIG
except ImportError:
    CRAWL_CONFIG = {}

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('scheduler.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)


def job():
    """每日爬取任务"""
    logging.info("=" * 50)
    logging.info(f"定时任务启动: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("=" * 50)

    try:
        # 执行爬取
        keywords = CRAWL_CONFIG.get("keywords", None)
        auto_open = CRAWL_CONFIG.get("auto_open_csv", True)

        # CI 环境（如 GitHub Actions）没有 GUI，自动禁用打开文件
        if os.environ.get("CI"):
            auto_open = False
            logging.info("检测到 CI 环境，已禁用自动打开 CSV")

        filename, count, stats = crawl(
            keywords=keywords,
            open_file=auto_open
        )

        # 发送通知
        title = "政府采购爬虫 - 每日报告"
        content = (
            f"爬取日期: {datetime.now().strftime('%Y-%m-%d')}\n"
            f"发现新项目: {count} 条\n"
            f"成功关键词: {stats['success']}\n"
            f"无结果关键词: {stats['empty']}\n"
            f"失败关键词: {stats['failed']}\n"
            f"CSV文件: {os.path.abspath(filename)}"
        )

        notify_all(title, content, attachment_path=filename if count > 0 else None)

        logging.info("定时任务完成")

    except Exception as e:
        logging.error(f"定时任务执行失败: {e}", exc_info=True)
        notify_all("政府采购爬虫 - 异常告警", f"任务执行失败:\n{str(e)}")


def run_once():
    """立即执行一次爬取（用于测试）"""
    print("=" * 50)
    print("立即执行一次爬取任务...")
    print("=" * 50)
    job()


def run_scheduler():
    """启动定时调度器"""
    schedule_time = CRAWL_CONFIG.get("schedule_time", "09:00")

    logging.info(f"调度器已启动，每天 {schedule_time} 执行爬取任务")
    print(f"定时器已启动，每天 {schedule_time} 自动运行")
    print("按 Ctrl+C 停止")

    schedule.every().day.at(schedule_time).do(job)

    while True:
        schedule.run_pending()
        time.sleep(60)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="政府采购爬虫定时调度器")
    parser.add_argument(
        "--once",
        action="store_true",
        help="立即执行一次（不启动定时循环，用于测试）"
    )
    args = parser.parse_args()

    if args.once:
        run_once()
    else:
        run_scheduler()
