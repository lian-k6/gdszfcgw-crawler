import requests
from bs4 import BeautifulSoup
import csv
import time
import re
import random
import os
import webbrowser
from datetime import datetime, timedelta

# 安全爬虫设置
REQUEST_DELAY_RANGE = (2, 5)  # 随机延迟范围（秒）
MAX_RETRIES = 3  # 最大重试次数

# 自定义User-Agent列表（替代fake_useragent）
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.5 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:115.0) Gecko/20100101 Firefox/115.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 16_5 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.5 Mobile/15E148 Safari/604.1"
]


def get_random_user_agent():
    """随机选择一个User-Agent"""
    return random.choice(USER_AGENTS)


def get_random_headers():
    """生成随机请求头，增强爬虫安全性"""
    return {
        "User-Agent": get_random_user_agent(),  # 使用自定义的User-Agent
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
        "Accept-Language": "zh-CN,zh;q=0.8,en-US;q=0.5,en;q=0.3",
        "Referer": "https://ygp.gdzwfw.gov.cn/zjfwcs/gd-zjcs-pub/purchaseNotice",
        "Connection": "keep-alive",
        "Cache-Control": "max-age=0",
        "Upgrade-Insecure-Requests": "1",
    }


def safe_request(url, method='get', data=None, headers=None, retries=0):
    """安全请求函数，带重试机制和错误处理"""
    try:
        # 如果未提供 headers，使用随机生成的
        if not headers:
            headers = get_random_headers()

        # 随机延迟，模拟人类浏览行为
        delay = random.uniform(*REQUEST_DELAY_RANGE)
        time.sleep(delay)

        if method.lower() == 'post':
            response = requests.post(url, data=data, headers=headers, timeout=10)
        else:
            response = requests.get(url, headers=headers, timeout=10)

        response.raise_for_status()  # 触发HTTP错误
        return response

    except requests.exceptions.RequestException as e:
        print(f"请求错误: {str(e)}")
        if retries < MAX_RETRIES:
            print(f"重试 ({retries + 1}/{MAX_RETRIES})...")
            # 重试时更换User-Agent
            new_headers = get_random_headers()
            return safe_request(url, method, data, new_headers, retries + 1)
        print(f"达到最大重试次数，请求失败")
        return None


def fetch_detail_page(url):
    """获取详情页数据，包含采购项目编码用于去重"""
    base_url = "https://ygp.gdzwfw.gov.cn"
    full_url = base_url + url if url.startswith('/') else url

    response = safe_request(full_url)
    if not response:
        return None

    try:
        soup = BeautifulSoup(response.text, 'html.parser')
        detail = {}

        # 标题
        title_tag = soup.find('h2', class_='wrap-title-center zjcsFontFace')
        detail['标题'] = title_tag.text.strip() if title_tag else ''

        # 提取所有列表项信息
        for li in soup.find_all('li'):
            b_tag = li.find('b')
            if b_tag:
                key = b_tag.text.strip().rstrip('：:')
                value_tag = li.find('div', class_='txt') or li.find('div', class_='txt zjcsFontFace')
                if value_tag:
                    # 清除链接和图片标签，但保留文本
                    for tag in value_tag.find_all(['a', 'img']):
                        tag.extract()
                    value = value_tag.text.strip()

                    # 映射到目标字段
                    if key == '项目业主':
                        detail['采购单位'] = value
                    elif key == '项目规模':
                        detail['预算总金额'] = value
                    elif key == '服务内容':
                        detail['项目内容'] = value
                    elif key == '采购项目名称':
                        detail['采购项目名称'] = value
                    elif key == '选取中介服务机构方式':  # 修正匹配的key
                        detail['选取中介服务机构方式'] = value
                    elif key == '采购项目编码':  # 用于去重的唯一标识
                        detail['采购项目编码'] = value

        # 处理可能的选取方式的其他格式
        if '选取中介服务机构方式' not in detail:
            # 专门查找包含"选取中介服务机构方式"的li标签
            select_mode_li = soup.find('li', string=re.compile(r'选取中介服务机构方式'))
            if select_mode_li:
                value_tag = select_mode_li.find('div', class_='txt') or select_mode_li.find('div',
                                                                                            class_='txt zjcsFontFace')
                if value_tag:
                    # 清除图片标签
                    for img in value_tag.find_all('img'):
                        img.extract()
                    detail['选取中介服务机构方式'] = value_tag.text.strip()

        # 如果仍未找到，尝试从正文中提取
        if '选取中介服务机构方式' not in detail:
            for p in soup.find_all('p'):
                if '选取方式' in p.text or '直接选取' in p.text:
                    # 提取相关文本
                    mode_text = re.search(r'(直接选取|随机抽取|竞争性谈判|询价|其他)[^。，,;；]*', p.text)
                    if mode_text:
                        detail['选取中介服务机构方式'] = mode_text.group()
                    else:
                        detail['选取中介服务机构方式'] = p.text.strip()
                    break

        # 从正文提取截止报名时间
        if '截止报名时间' not in detail:
            content_p = soup.find('p', class_='zjcsFontFace')
            if content_p:
                time_match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}', content_p.text)
                if time_match:
                    detail['截止报名时间'] = time_match.group()

        return detail

    except Exception as e:
        print(f"解析详情页 {full_url} 错误: {e}")
        return None


def fetch_purchase_notices():
    # 计算时间范围（当前时间前15天到当前时间）
    end_date = datetime.now()
    start_date = end_date - timedelta(days=5)

    # 格式化为YYYY-MM-DD
    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = end_date.strftime('%Y-%m-%d')

    print(f"时间范围: {start_date_str} 至 {end_date_str}")

    # 基础请求参数
    base_params = {
        "query_params_url": "/zjfwcs/gd-zjcs-pub/purchaseNotice",
        "query_params_rest_url": "purchaseNotice/listPost",
        "reloadQueryParamsReload": "false",
        "listVo.isTrackTotalHits": "true",
        "listVo.projectName": "",
        "listVo.purOrgName": "",
        "listVo.divisionCode": "442000",
        "listVo.selectModeType": "",
        "listVo.publishDateBegin": start_date_str,
        "listVo.publishDateEnd": end_date_str,
        "listVo.projectType": "",
        "listVo.selectServiceTypes": "001",
        "pageNumber": "0"  # 初始页码
    }

    all_data = []
    seen_ids = set()  # 用于存储已见过的唯一标识，实现去重
    page_number = 0
    total_processed = 0
    duplicates_removed = 0  # 统计移除的重复项数量
    url = "https://ygp.gdzwfw.gov.cn/zjfwcs/gd-zjcs-pub/purchaseNotice/listPost"

    try:
        while True:
            print(f"正在处理第 {page_number + 1} 页...")

            # 设置当前页码
            base_params["pageNumber"] = str(page_number)

            # 发送POST请求
            response = safe_request(url, method='post', data=base_params)
            if not response:
                print("获取列表页失败，尝试下一页...")
                page_number += 1
                continue

            soup = BeautifulSoup(response.text, 'html.parser')

            # 提取表格数据
            table = soup.find('table', class_='table normal')
            if not table:
                print("未找到表格数据，可能已无更多页面")
                break

            # 获取表头
            list_headers = [th.text.strip() for th in table.find('thead').find_all('th')]
            if not list_headers:
                print("未找到表头，可能已无更多页面")
                break

            # 处理当前页数据
            row_count = 0
            for tr in table.find('tbody').find_all('tr'):
                row_count += 1
                row_data = {}

                # 提取列表页信息
                for i, td in enumerate(tr.find_all('td')):
                    header = list_headers[i] if i < len(list_headers) else f"字段{i}"
                    a_tag = td.find('a')
                    if a_tag:
                        row_data[header] = a_tag.text.strip()
                        row_data[f"{header}_链接"] = a_tag.get('href', '')
                        # 从链接标题提取采购项目名称
                        if '采购项目名称' not in row_data:
                            row_data['采购项目名称'] = a_tag.text.strip()

                # 获取详情页数据
                if f"{list_headers[0]}_链接" in row_data:
                    print(f"  正在获取 {row_data.get(list_headers[0], '未知项目')} 的详细信息...")
                    detail = fetch_detail_page(row_data[f"{list_headers[0]}_链接"])
                    if detail:
                        row_data.update(detail)

                        # 去重逻辑：优先使用采购项目编码，没有则使用标题+采购单位组合
                        unique_id = detail.get('采购项目编码')
                        if not unique_id and '标题' in detail and '采购单位' in detail:
                            unique_id = f"{detail['标题']}|{detail['采购单位']}"

                        if unique_id:
                            if unique_id not in seen_ids:
                                seen_ids.add(unique_id)
                                all_data.append(row_data)
                                total_processed += 1
                            else:
                                duplicates_removed += 1
                                print(f"  发现重复记录，已跳过 (累计移除: {duplicates_removed})")
                        else:
                            # 如果无法生成唯一标识，仍添加记录但提示
                            all_data.append(row_data)
                            total_processed += 1
                            print("  无法生成唯一标识，可能存在重复记录")

            # 如果当前页没有数据，停止分页
            if row_count == 0:
                print("当前页无数据，停止分页")
                break

            # 准备下一页
            page_number += 1

        # 定义需要保留的字段
        required_fields = [
            '标题',
            '采购项目名称',
            '预算总金额',
            '采购单位',
            '截止报名时间',
            '选取中介服务机构方式',
            '项目内容'
        ]

        # 保存为CSV文件
        filename = 'zjcs.csv'
        with open(filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=required_fields)
            writer.writeheader()

            for row in all_data:
                # 只保留需要的字段，缺失字段留空
                filtered_row = {field: row.get(field, '') for field in required_fields}
                writer.writerow(filtered_row)

        print(f"处理完成，共提取 {total_processed} 条记录，移除 {duplicates_removed} 条重复记录，保存至 {filename}")

        # 自动打开CSV文件
        try:
            if os.name == 'nt':  # Windows
                os.startfile(filename)
            else:  # macOS/Linux
                webbrowser.open(f'file://{os.path.abspath(filename)}')
            print(f"已自动打开 {filename}")
        except Exception as e:
            print(f"自动打开文件失败: {e}，请手动打开 {filename}")

    except Exception as e:
        print(f"处理数据时发生错误: {e}")


if __name__ == "__main__":
    fetch_purchase_notices()