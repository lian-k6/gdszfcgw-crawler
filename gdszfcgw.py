import requests
import logging
import time
import random
import csv
import os
import platform
import subprocess
import math
from datetime import datetime, timedelta
from urllib.parse import urlencode

# 导入配置和通知模块
try:
    from config import Config, get_date_range, get_csv_filename
    from notifier import send_email, send_wechat
except ImportError:
    # 如果模块不存在，使用内嵌配置
    pass

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(module)s:%(lineno)d - %(message)s',
    handlers=[
        logging.FileHandler('gdzf_crawl.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)


# 爬虫配置
class Config:
    BASE_URL = "https://gdgpo.czt.gd.gov.cn/gpcms/rest/web/v2/info/selectInfoForIndex"
    DETAIL_BASE_URL = "https://gdgpo.czt.gd.gov.cn/freecms/site/gd/ggxx/detail.html"
    HOME_URL = "https://gdgpo.czt.gd.gov.cn/"

    USER_AGENTS = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.5 Safari/605.1.15",
        "Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36 Edg/116.0.1938.62"
    ]

    # 代理池配置
    PROXIES = [
        None  # 不使用代理
    ]

    FIXED_PARAMS = {
        "siteId": "cd64e06a-21a7-4620-aebc-0576bab7e07a",
        "channel": "fca71be5-fc0c-45db-96af-f513e9abda9d",
        "noticeType": "00101",
        "regionCode": "442000",
        "cityOrArea": "6",
        "purchaseManner": "1",
        "verifyCode": "1111",
        "subChannel": "false"
    }

    # 尝试从 config.py 加载用户自定义的筛选参数
    try:
        import config as _user_config
        _crawl_params = getattr(_user_config, "CRAWL_PARAMS", {})
        _param_map = {
            "region_code": "regionCode",
            "city_or_area": "cityOrArea",
            "purchase_manner": "purchaseManner",
            "notice_type": "noticeType",
        }
        for cfg_key, api_key in _param_map.items():
            val = _crawl_params.get(cfg_key)
            if val is not None:
                if val == "":
                    FIXED_PARAMS.pop(api_key, None)  # 空字符串 = 不限，移除该参数
                else:
                    FIXED_PARAMS[api_key] = val
    except ImportError:
        pass

    @classmethod
    def get_fixed_params(cls):
        """返回当前有效的固定参数（过滤掉空值）"""
        return {k: v for k, v in cls.FIXED_PARAMS.items() if v != ""}

    MAX_RETRIES = 5
    MIN_DELAY = 1.5
    MAX_DELAY = 5.0
    TIMEOUT = 25
    MAX_PAGES = 20
    CAPTCHA_RETRY_LIMIT = 2

    # 邮件配置（用户需自行填写）
    EMAIL_CONFIG = {
        "smtp_server": "smtp.qq.com",       # 例如：smtp.qq.com
        "smtp_port": 465,
        "sender_email": "",                  # 发件人邮箱
        "sender_password": "",               # 邮箱授权码（不是登录密码）
        "receiver_emails": []                # 收件人列表
    }

    # 微信通知配置
    WECHAT_CONFIG = {
        # Server酱 SendKey，获取地址：https://sct.ftqq.com/
        "server_chan_key": "",
        # 企业微信机器人 webhook，例如：https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxx
        "qywx_webhook": ""
    }

    # 定时任务时间（24小时制）
    SCHEDULE_TIME = "09:00"


def get_date_range(days=None):
    """获取时间范围：今天往前推 N 天 到 今天 23:59:59"""
    if days is None:
        try:
            import config as _cfg
            days = getattr(_cfg, "CRAWL_CONFIG", {}).get("days_back", 30)
        except ImportError:
            days = 30
    end_time = datetime.now()
    start_time = end_time - timedelta(days=days)
    return (
        start_time.strftime("%Y-%m-%d 00:00:00"),
        end_time.strftime("%Y-%m-%d 23:59:59")
    )


def get_csv_filename():
    """生成带日期的CSV文件名"""
    today = datetime.now().strftime("%Y%m%d")
    return f"政府采购项目_{today}.csv"


def get_headers():
    return {
        "User-Agent": random.choice(Config.USER_AGENTS),
        "Referer": "https://gdgpo.czt.gd.gov.cn/",
        "Accept": "application/json, text/plain, */*",
        "X-Requested-With": "XMLHttpRequest",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Connection": "keep-alive"
    }


def get_proxy():
    return random.choice(Config.PROXIES) if Config.PROXIES else None


def create_session():
    """创建带session的requests对象，先访问首页预热"""
    session = requests.Session()
    # 禁用从环境变量自动读取代理（避免系统代理配置导致连接失败）
    session.trust_env = False
    try:
        session.get(
            Config.HOME_URL,
            headers={"User-Agent": random.choice(Config.USER_AGENTS)},
            timeout=10
        )
        logging.info("Session预热完成")
        time.sleep(1)
    except Exception as e:
        logging.warning(f"Session预热失败（不影响后续请求）: {e}")
    return session


def get_detail_content(session, item, retry=0):
    """通过 getInfoById API 获取公告详情页完整正文"""
    notice_id = item.get("id", "")
    if not notice_id:
        return ""
    try:
        time.sleep(random.uniform(0.5, 1.5))
        url = "https://gdgpo.czt.gd.gov.cn/gpcms/rest/web/v2/info/getInfoById"
        r = session.get(
            url,
            params={"id": notice_id},
            headers=get_headers(),
            timeout=15
        )
        r.raise_for_status()
        data = r.json()
        if data.get("code") == "200":
            d = data.get("data", {})
            # 正文在 content / noticeContent / detail 中
            content = d.get("content") or d.get("noticeContent") or d.get("detail") or ""
            # 同时提取标题、描述等文本
            texts = [
                str(d.get("title", "")),
                str(d.get("description", "")),
                str(d.get("purchaser", "")),
                str(d.get("agency", "")),
                str(d.get("regionName", "")),
                content
            ]
            return " ".join(texts)
        return ""
    except Exception as e:
        if retry < 2:
            time.sleep(2 * (retry + 1))
            return get_detail_content(session, item, retry + 1)
        logging.warning(f"获取详情失败 id={notice_id}: {e}")
        return ""


def solve_captcha(captcha_image_url):
    logging.warning("遇到验证码，需要人工处理或接入识别服务")
    return "1234"


def generate_detail_url(item):
    try:
        params = {
            "htmlIndexnum": item.get("htmlIndexnum", ""),
            "siteId": Config.FIXED_PARAMS["siteId"],
            "channel": Config.FIXED_PARAMS["channel"]
        }
        valid_params = {k: v for k, v in params.items() if v}
        if not valid_params.get("htmlIndexnum"):
            logging.warning(f"缺少htmlIndexnum，无法生成详情链接: {item.get('title')}")
            return ""
        return f"{Config.DETAIL_BASE_URL}?{urlencode(valid_params)}"
    except Exception as e:
        logging.error(f"生成详情链接失败: {str(e)}")
        return ""


def fetch_page_data(session, page_num, keyword="", retry=0, captcha_retry=0):
    try:
        start_time, end_time = get_date_range(90)
        params = {
            **Config.get_fixed_params(),
            "keyword": keyword,
            "currPage": page_num,
            "pageSize": 40,
            "operationStartTime": start_time,
            "operationEndTime": end_time,
            "_t": int(time.time() * 1000)
        }

        delay = random.uniform(Config.MIN_DELAY, Config.MAX_DELAY)
        time.sleep(delay)

        proxies = get_proxy()
        logging.debug(f"使用代理: {proxies}")

        response = session.get(
            url=Config.BASE_URL,
            params=params,
            headers=get_headers(),
            proxies=proxies,
            timeout=Config.TIMEOUT,
            allow_redirects=False
        )

        logging.debug(f"请求URL: {response.url}")
        logging.debug(f"响应状态码: {response.status_code}")

        if response.status_code == 403 and "验证码" in response.text:
            logging.warning(f"第{page_num}页触发验证码保护")
            if captcha_retry < Config.CAPTCHA_RETRY_LIMIT:
                captcha_code = solve_captcha("")
                if captcha_code:
                    logging.info(f"尝试输入验证码: {captcha_code}")
                    params["verifyCode"] = captcha_code
                    return fetch_page_data(session, page_num, keyword, retry, captcha_retry + 1)
            logging.error("超过验证码重试限制")
            return None, 0, 0

        if response.status_code in [403, 429, 503]:
            logging.warning(f"第{page_num}页被反爬限制，状态码: {response.status_code}")
            if retry < Config.MAX_RETRIES:
                wait_time = 10 * (retry + 1)
                logging.info(f"第{retry + 1}次重试，等待{wait_time}秒...")
                time.sleep(wait_time)
                return fetch_page_data(session, page_num, keyword, retry + 1, captcha_retry)
            return None, 0, 0

        response.raise_for_status()
        logging.info(f"第{page_num}页请求成功，耗时{delay:.2f}秒")

        result = response.json()
        logging.debug(f"第{page_num}页响应: {result}")

        code = result.get("code", "")
        msg = result.get("msg", "")
        if code == "200" and msg == "操作成功":
            data = result.get("data", {})
            total_page = data.get("totalPage", 0)
            total = data.get("total", 0)
            rows = data.get("rows", [])

            # API偶尔返回 totalPage=0 但 total>0，根据total和pageSize计算实际页数
            if total > 0 and total_page == 0:
                page_size = params.get("pageSize", 40)
                total_page = math.ceil(total / page_size)
                logging.debug(f"API返回totalPage=0但total={total}，修正为{total_page}页")

            return rows, min(total_page, Config.MAX_PAGES), total
        else:
            logging.error(f"第{page_num}页业务失败: code={code}, msg={msg}")
            if "验证码" in msg and captcha_retry < Config.CAPTCHA_RETRY_LIMIT:
                captcha_code = solve_captcha("")
                if captcha_code:
                    params["verifyCode"] = captcha_code
                    return fetch_page_data(session, page_num, keyword, retry, captcha_retry + 1)
            return None, 0, 0

    except requests.exceptions.Timeout:
        logging.error(f"第{page_num}页请求超时")
        if retry < Config.MAX_RETRIES:
            logging.info(f"第{retry + 1}次重试...")
            time.sleep(5 * (retry + 1))
            return fetch_page_data(session, page_num, keyword, retry + 1, captcha_retry)
        return None, 0, 0

    except Exception as e:
        logging.error(f"第{page_num}页错误: {str(e)}")

    if retry < Config.MAX_RETRIES:
        logging.info(f"第{retry + 1}次重试...")
        time.sleep(5 * (retry + 1))
        return fetch_page_data(session, page_num, keyword, retry + 1, captcha_retry)
    return None, 0, 0


def parse_item(item):
    return {
        "标题": item.get("title", ""),
        "采购人": item.get("purchaser", ""),
        "代理机构": item.get("agency", ""),
        "公告时间": item.get("noticeTime", ""),
        "开标时间": item.get("openTenderTime", ""),
        "地区": item.get("regionName", ""),
        "项目描述": (item.get("description", "")[:500] + "...")
        if len(item.get("description", "")) > 500 else item.get("description", ""),
        "计划编号": item.get("planCodes", ""),
        "详情链接": generate_detail_url(item)
    }


def save_to_csv(new_data, filename=None, mode="overwrite"):
    """
    保存数据到CSV
    :param new_data: 要保存的数据列表
    :param filename: 文件名，默认按日期生成
    :param mode: 写入模式，"overwrite"=覆盖（默认），"append"=追加
    """
    if not new_data:
        logging.warning("无新数据可保存")
        return False

    if filename is None:
        filename = get_csv_filename()

    fieldnames = [
        "标题", "采购人", "代理机构",
        "公告时间", "开标时间", "地区",
        "项目描述", "计划编号", "详情链接"
    ]

    try:
        if mode == "overwrite" and os.path.exists(filename):
            os.remove(filename)
            logging.info(f"已删除旧文件 {filename}，准备重新写入")

        file_exists = os.path.exists(filename)

        with open(filename, "a" if file_exists else "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            if not file_exists:
                writer.writeheader()
            writer.writerows(new_data)

        logging.info(f"已保存 {len(new_data)} 条数据到 {filename}")
        return True
    except Exception as e:
        logging.error(f"保存CSV失败: {str(e)}")
        return False


def open_csv_file(filename):
    try:
        if platform.system() == 'Windows':
            os.startfile(filename)
        elif platform.system() == 'Darwin':
            subprocess.call(('open', filename))
        else:
            subprocess.call(('xdg-open', filename))
        logging.info(f"已打开文件: {filename}")
    except Exception as e:
        logging.error(f"打开文件失败: {str(e)}")
        logging.info(f"请手动打开文件: {os.path.abspath(filename)}")


def crawl(keywords=None, open_file=True):
    """
    执行爬取任务
    策略：先获取全部公告列表，再在前端用关键词做全文过滤
    （API的title参数只在标题搜索，过于严格；先获取全部后本地过滤更完整）
    :param keywords: 关键词列表，默认使用内置列表
    :param open_file: 是否自动打开CSV文件
    :return: (csv_filename, total_count, stats_dict)
    """
    if keywords is None:
        keywords = [
            '地下市政', '改造片区', '工程建设', '管线', '规划', '规划验收', '国土', '基础测绘', '基坑',
            '建设工程', '界线', '联合测绘', '倾斜摄影', '市政', '水准点', '土地变更', '土地储备', '土地整备', '土壤普查', '违法用地',
            '信息化系统维护', '验线', '影像数据', '宅基地', '征拆', '征地', '整治工程', '整治提升工程',
            '不动产', '测绘', '测量', '城市更新', '档案管理', '道路改造', '地理国情', '地理信息', '地名普查'
        ]

    # 去重并保持顺序
    seen = set()
    keywords = [k for k in keywords if not (k in seen or seen.add(k)
                                            )]

    session = create_session()
    stats = {"success": 0, "empty": 0, "failed": 0, "keywords": len(keywords)}

    start_dt, end_dt = get_date_range(90)
    logging.info(f"爬取时间范围: {start_dt} ~ {end_dt}")

    try:
        # ========== 第一步：获取全部公告列表（不带搜索词）==========
        logging.info("\n" + "=" * 40)
        logging.info("第一步：获取全部公告列表")
        logging.info("=" * 40)

        all_raw_items = []
        first_page, total_pages, total_records = fetch_page_data(session, 1, "")

        if first_page is None:
            logging.error("无法获取数据，退出")
            stats["failed"] = len(keywords)
            return get_csv_filename(), 0, stats

        if total_records == 0:
            logging.info("该时间范围内无公告数据")
            return get_csv_filename(), 0, stats

        all_raw_items.extend(first_page)
        logging.info(f"共 {total_records} 条公告，预计 {total_pages} 页")

        # 获取后续页面
        if total_pages > 1:
            for page in range(2, total_pages + 1):
                page_data, _, _ = fetch_page_data(session, page, "")
                if page_data:
                    all_raw_items.extend(page_data)
                    logging.info(f"已获取第{page}/{total_pages}页，累计 {len(all_raw_items)} 条")
                else:
                    logging.warning(f"第{page}页获取失败，跳过")

        logging.info(f"全部列表获取完成，共 {len(all_raw_items)} 条原始公告")

        # ========== 第二步：用关键词在前端做全文过滤 ==========
        logging.info("\n" + "=" * 40)
        logging.info("第二步：关键词本地全文过滤")
        logging.info("=" * 40)

        # 加载排除词（黑名单）
        try:
            import config as _cfg
            exclude_keywords = [k.lower() for k in getattr(_cfg, "CRAWL_CONFIG", {}).get("exclude_keywords", [])]
        except ImportError:
            exclude_keywords = []

        matched_items = []
        seen_keys = set()
        matched_keywords = set()
        unmatched_records = []  # 记录未在标题+描述中匹配到的原始记录
        excluded_count = 0

        # --- 第一轮：标题 + 描述等字段快速过滤 ---
        for item in all_raw_items:
            search_fields = [
                str(item.get("title", "")),
                str(item.get("description", "")),
                str(item.get("purchaser", "")),
                str(item.get("agency", "")),
                str(item.get("regionName", "")),
                str(item.get("catalogueNameList", "")),
            ]
            search_text = " ".join(search_fields).lower()

            # 先检查排除词（黑名单）
            if exclude_keywords:
                is_excluded = any(ex_kw in search_text for ex_kw in exclude_keywords)
                if is_excluded:
                    excluded_count += 1
                    continue

            hit = False
            for kw in keywords:
                if kw.lower() in search_text:
                    parsed = parse_item(item)
                    key = parsed.get("计划编号") or f"{parsed.get('标题', '')}|{parsed.get('公告时间', '')}"
                    if key and key not in seen_keys:
                        seen_keys.add(key)
                        matched_items.append(parsed)
                        matched_keywords.add(kw)
                    hit = True
                    break
            if not hit:
                unmatched_records.append(item)

        first_hit_count = len(matched_items)
        logging.info(f"第一轮快速过滤命中 {first_hit_count} 条，排除 {excluded_count} 条无关项目")

        # --- 第二轮：获取详情页正文做补充过滤 ---
        remaining_keywords = [kw for kw in keywords if kw not in matched_keywords]
        if unmatched_records and remaining_keywords:
            logging.info(f"第二轮：对 {len(unmatched_records)} 条记录获取详情页正文，检查 {len(remaining_keywords)} 个未命中关键词")
            detail_cache = {}
            for idx, item in enumerate(unmatched_records):
                if not remaining_keywords:
                    break
                notice_id = item.get("id", "")
                if not notice_id:
                    continue

                # 获取详情页正文（带缓存）
                if notice_id in detail_cache:
                    full_text = detail_cache[notice_id]
                else:
                    full_text = get_detail_content(session, item)
                    detail_cache[notice_id] = full_text

                if not full_text:
                    continue

                full_text_lower = full_text.lower()

                # 详情页也检查排除词
                if exclude_keywords and any(ex_kw in full_text_lower for ex_kw in exclude_keywords):
                    excluded_count += 1
                    continue

                hit = False
                for kw in list(remaining_keywords):
                    if kw.lower() in full_text_lower:
                        parsed = parse_item(item)
                        key = parsed.get("计划编号") or f"{parsed.get('标题', '')}|{parsed.get('公告时间', '')}"
                        if key and key not in seen_keys:
                            seen_keys.add(key)
                            matched_items.append(parsed)
                            matched_keywords.add(kw)
                            remaining_keywords.remove(kw)
                        hit = True
                        break  # 只记录第一个匹配的关键词

                if (idx + 1) % 10 == 0:
                    logging.info(f"已检查 {idx + 1}/{len(unmatched_records)} 条详情页")

        # 统计匹配情况
        unmatched_keywords = [kw for kw in keywords if kw not in matched_keywords]
        stats["success"] = len(matched_keywords)
        stats["empty"] = len(unmatched_keywords)
        stats["detail_checked"] = len(unmatched_records)
        stats["excluded"] = excluded_count

        logging.info("\n" + "-" * 40)
        for kw in sorted(matched_keywords):
            logging.info(f"[命中] {kw}")
        for kw in sorted(unmatched_keywords):
            logging.info(f"[未命中] {kw}")

        unique_count = len(matched_items)
        logging.info(f"\n{'=' * 40}")
        logging.info(f"过滤完成，共匹配 {unique_count} 条公告")
        logging.info(f"命中关键词: {stats['success']} 个，未命中: {stats['empty']} 个")
        logging.info(f"排除无关项目: {excluded_count} 条")
        logging.info(f"详情页补充检查: {stats['detail_checked']} 条")
        logging.info(f"{'=' * 40}")

        # ========== 第三步：保存结果 ==========
        csv_filename = get_csv_filename()
        if matched_items:
            if save_to_csv(matched_items, csv_filename):
                if open_file:
                    open_csv_file(csv_filename)

            sample_urls = [item["详情链接"] for item in matched_items[:3] if item["详情链接"]]
            if sample_urls:
                logging.info("详情页链接示例:")
                for url in sample_urls:
                    logging.info(f"- {url}")
        else:
            logging.info("本次未匹配到符合条件的数据")

        return csv_filename, unique_count, stats

    except Exception as e:
        logging.error(f"主程序错误: {str(e)}", exc_info=True)
        return get_csv_filename(), 0, stats
    finally:
        session.close()
        logging.info("爬虫任务结束")


def main():
    filename, count, stats = crawl()

    # 如果配置了通知，尝试发送
    try:
        from notifier import send_email, send_wechat
        if count > 0:
            send_wechat(
                title="政府采购爬虫通知",
                content=f"今日爬取完成，共发现 {count} 条新项目\n"
                        f"成功关键词: {stats['success']}, 无结果: {stats['empty']}, 失败: {stats['failed']}"
            )
            send_email(
                subject=f"政府采购项目数据 - {datetime.now().strftime('%Y-%m-%d')}",
                body=f"爬取完成，共 {count} 条新项目，详见附件。",
                attachment_path=filename
            )
    except Exception as e:
        logging.warning(f"通知发送失败（请检查配置）: {e}")


if __name__ == "__main__":
    main()
