import os
import shutil
import pandas as pd
import time
import logging
import pythoncom
import hashlib
from pyautocad import Autocad
from tqdm import tqdm

# ----------------------
# 配置参数（新增备份和历史记录配置）
# ----------------------
CONFIG = {
    "root_dir": r"D:\cadtu",
    "excel_suffix": ".xlsx",
    "cad_suffix": ".dwg",
    "letter_fields": {
        "DZ": "DZ", "DH": "DH", "JZMJ": "JZMJ", "ZYMJ": "ZYMJ",
        "FTMJ": "FTMJ", "HH": "HH", "SZC": "SZC", "ZCS": "ZCS",
        "FXMJ1": "FX1", "FXMJ2": "FX2", "BH": "BH"
    },
    "filename_match_col": "对应分户图编号",
    "retry_times": 3,
    "load_delay": 3,
    "operation_delay": 0.5,
    "log_path": "building_cad_log.txt",
    "number_format": "%.2f",
    "backup_dir": "cad_backups",  # 新增：CAD备份文件夹
    "history_path": "processing_history.csv"  # 新增：处理历史记录
}

# ----------------------
# 初始化日志和历史记录
# ----------------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(CONFIG["log_path"], encoding='utf-8'),
        logging.StreamHandler()
    ]
)


def init_history():
    """初始化历史记录文件（首次运行创建表头）"""
    if not os.path.exists(CONFIG["history_path"]):
        with open(CONFIG["history_path"], "w", encoding="utf-8") as f:
            f.write("cad_path,excel_row_id,process_time,hash,status\n")


# ----------------------
# 数据格式化函数
# ----------------------
def format_value(value, field_name):
    try:
        if pd.isna(value):
            return ""

        if isinstance(value, (int, float)):
            if field_name in ["SZC", "ZCS", "HH"]:
                return str(int(value))
            return CONFIG["number_format"] % value

        return str(value).strip()
    except:
        return str(value)


# ----------------------
# 新增：CAD文件备份功能
# ----------------------
def backup_cad_file(cad_path):
    """备份CAD文件到指定目录"""
    if not cad_path or not os.path.exists(cad_path):
        return False

    # 创建按日期分类的备份目录
    backup_root = os.path.join(os.path.dirname(cad_path), CONFIG["backup_dir"])
    date_dir = time.strftime("%Y%m%d")
    backup_dir = os.path.join(backup_root, date_dir)
    os.makedirs(backup_dir, exist_ok=True)

    # 生成带时间戳的备份文件名
    cad_name = os.path.basename(cad_path)
    name, ext = os.path.splitext(cad_name)
    backup_name = f"{name}_backup_{time.strftime('%H%M%S')}{ext}"
    backup_path = os.path.join(backup_dir, backup_name)

    try:
        shutil.copy2(cad_path, backup_path)  # 保留文件元数据
        logging.debug(f"已备份CAD文件: {backup_path}")
        return True
    except Exception as e:
        logging.warning(f"备份CAD文件 {cad_name} 失败: {str(e)}")
        return False


# ----------------------
# 新增：记录处理历史
# ----------------------
def log_history(cad_path, excel_row_id, row_hash, status="success"):
    """记录处理历史到CSV文件"""
    with open(CONFIG["history_path"], "a", encoding="utf-8") as f:
        row = f"{cad_path},{excel_row_id},{time.strftime('%Y-%m-%d %H:%M:%S')},{row_hash},{status}\n"
        f.write(row)


# ----------------------
# 新增：计算数据行哈希值（用于增量更新）
# ----------------------
def get_row_hash(row):
    """将Excel行数据转换为哈希值，检测数据是否修改"""
    row_str = ",".join([str(v) for v in row.values])
    return hashlib.md5(row_str.encode()).hexdigest()


# ----------------------
# 处理单个CAD文件（优化版）
# ----------------------
def process_cad_file(acad, cad_path, data_row, row_idx, row_hash):
    for retry in range(CONFIG["retry_times"] + 1):
        try:
            # 处理前先备份
            backup_cad_file(cad_path)

            # 打开文件
            doc = acad.Application.Documents.Open(cad_path)
            time.sleep(CONFIG["load_delay"])

            # 预收集文本对象映射
            text_objects = {}
            for obj in doc.ModelSpace:
                if obj.ObjectName in ["AcDbText", "AcDbMText"]:
                    key = obj.TextString.strip()
                    text_objects[key] = obj

            filled_count = 0
            # 处理所有字段
            for cad_field, excel_col in CONFIG["letter_fields"].items():
                if pd.isna(data_row[excel_col]):
                    continue

                if cad_field in text_objects:
                    obj = text_objects[cad_field]
                    obj.TextString = format_value(data_row[excel_col], cad_field)
                    filled_count += 1

            time.sleep(CONFIG["operation_delay"])

            # 保存并关闭
            try:
                doc.Save()
            except:
                doc.SaveAs(cad_path)
            doc.Close()

            # 记录成功历史
            log_history(cad_path, row_idx, row_hash)
            return filled_count if filled_count > 0 else 0

        except Exception as e:
            if retry == CONFIG["retry_times"]:
                logging.error(f"处理CAD文件 {os.path.basename(cad_path)} 最终失败: {str(e)}")
                log_history(cad_path, row_idx, row_hash, "failed")
                return -1
            wait_time = (retry + 1) * 2
            logging.warning(f"处理CAD文件 {os.path.basename(cad_path)} 第{retry + 1}次失败，等待{wait_time}秒后重试: {str(e)}")
            time.sleep(wait_time)


# ----------------------
# 处理单个楼栋文件夹（支持增量更新）
# ----------------------
def process_building_folder(acad, folder_path):
    try:
        building_name = os.path.basename(folder_path)
        logging.info(f"\n===== 开始处理楼栋: {building_name} =====")

        # 查找Excel和CAD文件
        excel_files = [f for f in os.listdir(folder_path)
                       if f.lower().endswith(CONFIG["excel_suffix"])]
        cad_files = [f for f in os.listdir(folder_path)
                     if f.lower().endswith(CONFIG["cad_suffix"])]

        if not excel_files:
            logging.warning(f"楼栋 {building_name} 未找到Excel文件")
            return (building_name, 0, len(cad_files), "无Excel文件")

        if not cad_files:
            logging.warning(f"楼栋 {building_name} 未找到CAD文件")
            return (building_name, 0, 0, "无CAD文件")

        # 加载Excel数据
        excel_path = os.path.join(folder_path, excel_files[0])
        try:
            df = pd.read_excel(excel_path)
            if CONFIG["filename_match_col"] not in df.columns:
                raise ValueError(f"Excel中缺少匹配列: {CONFIG['filename_match_col']}")

            # 读取历史记录（用于增量更新判断）
            history_df = pd.read_csv(CONFIG["history_path"]) if os.path.exists(CONFIG["history_path"]) else None

            # 构建文件名到数据的映射（包含行索引和哈希）
            filename_to_data = {}
            for idx, row in df.iterrows():
                cad_name = str(row[CONFIG["filename_match_col"]])
                current_hash = get_row_hash(row)

                # 检查是否需要处理：无历史记录 或 数据已修改
                need_process = True
                if history_df is not None:
                    # 查找相同行索引且哈希一致的记录
                    same_records = history_df[
                        (history_df["excel_row_id"] == idx) &
                        (history_df["hash"] == current_hash)
                        ]
                    if not same_records.empty:
                        need_process = False
                        logging.debug(f"Excel行 {idx} 未修改，跳过处理 {cad_name}")

                if need_process:
                    filename_to_data[cad_name] = (idx, row, current_hash)

            logging.info(f"楼栋 {building_name} 加载Excel数据成功，需处理{len(filename_to_data)}条记录")
        except Exception as e:
            logging.error(f"楼栋 {building_name} 加载Excel失败: {str(e)}")
            return (building_name, 0, len(cad_files), "Excel加载失败")

        # 处理CAD文件
        success_count = 0
        for cad_file in tqdm(cad_files, desc=f"处理 {building_name}"):
            cad_path = os.path.join(folder_path, cad_file)
            cad_name = os.path.splitext(cad_file)[0]

            if cad_name in filename_to_data:
                row_idx, row_data, row_hash = filename_to_data[cad_name]
                result = process_cad_file(acad, cad_path, row_data, row_idx, row_hash)
                if result > 0:
                    success_count += 1
                    logging.info(f"楼栋 {building_name} 的 {cad_file} 填充成功，共{result}个字段")
                elif result == 0:
                    logging.warning(f"楼栋 {building_name} 的 {cad_file} 未填充任何字段")
            else:
                logging.debug(f"楼栋 {building_name} 的 {cad_file} 无需处理（数据未修改或无匹配）")

        return (building_name, success_count, len(cad_files), "处理完成")

    except Exception as e:
        logging.error(f"处理楼栋 {folder_path} 出错: {str(e)}", exc_info=True)
        return (os.path.basename(folder_path), 0, 0, f"处理错误: {str(e)}")


# ----------------------
# 主函数
# ----------------------
def main():
    start_time = time.time()
    pythoncom.CoInitialize()
    acad = None

    try:
        # 初始化历史记录
        init_history()

        # 收集所有楼栋文件夹
        building_folders = [
            os.path.join(CONFIG["root_dir"], f)
            for f in os.listdir(CONFIG["root_dir"])
            if os.path.isdir(os.path.join(CONFIG["root_dir"], f))
        ]

        if not building_folders:
            logging.warning(f"在 {CONFIG['root_dir']} 中未找到任何楼栋文件夹")
            return

        logging.info(f"发现 {len(building_folders)} 个楼栋文件夹，采用单线程处理")

        # 初始化AutoCAD
        try:
            acad = Autocad(create_if_not_exists=True)
            acad.Application.Visible = True
            logging.info(f"AutoCAD连接成功，版本: {acad.Application.Version}")
        except Exception as e:
            logging.error(f"AutoCAD连接失败: {str(e)}")
            return

        # 处理所有楼栋
        results = []
        for folder in building_folders:
            result = process_building_folder(acad, folder)
            results.append(result)

        # 生成报告
        total_success = 0
        total_files = 0

        logging.info("\n" + "=" * 60)
        logging.info("所有楼栋处理完成！")
        logging.info("-" * 60)
        logging.info(f"楼栋总数: {len(building_folders)}")

        for building_name, success, total, status in results:
            total_success += success
            total_files += total
            logging.info(f"楼栋 {building_name}: 成功 {success}/{total}，状态: {status}")

        logging.info("-" * 60)
        logging.info(f"总处理CAD文件: {total_files}")
        logging.info(f"成功填充: {total_success}")
        logging.info(f"成功率: {total_success / total_files * 100:.2f}%" if total_files > 0 else "无文件处理")
        logging.info(f"总耗时: {time.time() - start_time:.2f}秒")
        logging.info("=" * 60)

    finally:
        # 确保资源释放
        if acad:
            try:
                acad.Application.Quit()
                logging.info("AutoCAD已关闭")
            except:
                pass
        pythoncom.CoUninitialize()


# ----------------------
# 执行
# ----------------------
if __name__ == "__main__":
    main()