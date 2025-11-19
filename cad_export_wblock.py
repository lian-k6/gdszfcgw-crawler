# -*- coding: utf-8 -*-
import win32com.client
import pythoncom
import pyautogui
import time
import os
import math
import threading
from datetime import datetime

# =============== 配置区 ===============
SRC_FILE = r"D:\cadtu\地下室分户图(采用).dwg"
OUTPUT_DIR = r"D:\output"
LOG_FILE = os.path.join(OUTPUT_DIR, "cad_parking_export_v16.log")

TEXT_LAYER = "0-小车位编号"
CAR_LAYER = "0-小车车位"

BUFFER = 11.0            # 矩形框半宽（单位与 DWG 相同）
LABEL_OFFSET = 4.0       # 注记偏移距离（m）
TEST_FIRST_N = 50         # 测试前 N 个；设 None 为全部

AUTOCAD_PROG_ID = "AutoCAD.Application"
AUTOCAD_VISIBLE = True

# 弹窗检测模板（可选），截图一个能稳定识别的对话框局部（如“文件名”左侧 label）
SAVE_DIALOG_IMAGE = None  # 示例： r"D:\output\save_dialog_template.png" 或 None

# fallback 屏幕坐标（当模板不可用或检测失败时使用）
FALLBACK_INPUT_POS = (694, 653)

# initial_delay 为在触发 SAVET 命令后线程等待的秒数（你要求 20s）
SAVE_DIALOG_INITIAL_DELAY = 10.0
# 在 initial_delay 之后再等待 detect_timeout 秒进行 locateOnScreen 检测
SAVE_DIALOG_DETECT_TIMEOUT = 20.0
# 临时对象图层名称（改这里即可）
TEMP_LAYER = "TEMP_EXPORT"
# =====================================

# ---------- AutoCAD COM ----------
def get_acad_instance():
    pythoncom.CoInitialize()
    try:
        acad = win32com.client.GetActiveObject(AUTOCAD_PROG_ID)
        log("已附着到现有 AutoCAD 实例")
    except Exception:
        log("未发现运行中的 AutoCAD，尝试 Dispatch 启动...")
        acad = win32com.client.Dispatch(AUTOCAD_PROG_ID)
        log("已启动 AutoCAD")
    try:
        acad.Visible = AUTOCAD_VISIBLE
    except Exception:
        pass
    return acad

# ========== TEMP_LAYER 管理 =============
def ensure_temp_layer(doc, layer_name=TEMP_LAYER):
    """
    确保临时图层存在并返回原始活动图层供恢复。
    返回 (orig_active_layer_obj_or_None)
    """
    try:
        layers = doc.Layers
        try:
            temp_layer = layers.Item(layer_name)
        except Exception:
            # 创建图层
            temp_layer = layers.Add(layer_name)
            # 可选：设置颜色（比如 7/白），线型等
            try:
                temp_layer.Color = 7
            except Exception:
                pass
        # 记录当前活动图层
        try:
            orig_layer = doc.ActiveLayer
        except Exception:
            orig_layer = None
        # 切换到临时图层
        try:
            doc.ActiveLayer = temp_layer
        except Exception:
            # 备用：使用命令切换
            try:
                doc.SendCommand(f'(command "._-layer" "set" "{layer_name}" "")\n')
            except Exception:
                pass
        return orig_layer
    except Exception as e:
        log(f"ensure_temp_layer 错误: {e}")
        return None

def restore_active_layer(doc, orig_layer):
    try:
        if orig_layer is not None:
            try:
                doc.ActiveLayer = orig_layer
            except Exception:
                # 备用：通过名字设置
                try:
                    doc.SendCommand(f'(command "._-layer" "set" "{orig_layer.Name}" "")\n')
                except Exception:
                    pass
    except Exception:
        pass

# ---------- 文本 / polyline 提取 ----------
def collect_texts(msp, layer=TEXT_LAYER):
    res = []
    for ent in msp:
        try:
            if getattr(ent, "Layer", "") == layer and "Text" in getattr(ent, "ObjectName", ""):
                txt = getattr(ent, "TextString", None) or getattr(ent, "Contents", "")
                ip = getattr(ent, "InsertionPoint", None)
                if txt is not None and ip:
                    res.append((str(txt).strip(), float(ip[0]), float(ip[1])))
        except Exception:
            continue
    return res

def find_car_polyline(msp, text, tol=200.0):
    tx, ty = text[1], text[2]
    candidates = []
    for ent in msp:
        try:
            if getattr(ent, "Layer", "") == CAR_LAYER and "Polyline" in getattr(ent, "ObjectName", ""):
                coords = getattr(ent, "Coordinates", None)
                if not coords:
                    continue
                pts = [(float(coords[i]), float(coords[i+1])) for i in range(0, len(coords), 2)]
                if not pts:
                    continue
                cx = sum(p[0] for p in pts) / len(pts)
                cy = sum(p[1] for p in pts) / len(pts)
                dist = math.hypot(cx - tx, cy - ty)
                if dist <= tol:
                    candidates.append((dist, ent, pts))
        except Exception:
            continue
    if candidates:
        candidates.sort(key=lambda x: x[0])
        return candidates[0][1], candidates[0][2]
    return None, None

# ---------- 辅助几何函数 ----------
def poly_bbox(pts):
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))

def convex_hull(points):
    pts = sorted(set(points))
    if len(pts) <= 1:
        return pts
    def cross(o, a, b):
        return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lower = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    upper = []
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    hull = lower[:-1] + upper[:-1]
    return hull

def angle_between(a, b, c):
    bax = a[0] - b[0]; bay = a[1] - b[1]
    cbx = c[0] - b[0]; cby = c[1] - b[1]
    da = math.hypot(bax, bay)
    db = math.hypot(cbx, cby)
    if da < 1e-9 or db < 1e-9:
        return 0.0
    dot = (bax*cbx + bay*cby) / (da*db)
    dot = max(-1.0, min(1.0, dot))
    return math.acos(dot)

def oriented_bbox_from_points(pts):
    if not pts:
        return None
    mx = sum(p[0] for p in pts) / len(pts)
    my = sum(p[1] for p in pts) / len(pts)
    sxx = sum((p[0]-mx)*(p[0]-mx) for p in pts) / len(pts)
    sxy = sum((p[0]-mx)*(p[1]-my) for p in pts) / len(pts)
    syy = sum((p[1]-my)*(p[1]-my) for p in pts) / len(pts)
    trace = sxx + syy
    disc = max(0.0, (sxx - syy)*(sxx - syy) + 4.0*sxy*sxy)
    sqrt_disc = math.sqrt(disc)
    eig1 = (trace + sqrt_disc) / 2.0
    if abs(sxy) > 1e-12:
        v1x = eig1 - syy
        v1y = sxy
    else:
        v1x = 1.0 if sxx >= syy else 0.0
        v1y = 0.0 if sxx >= syy else 1.0
    norm = math.hypot(v1x, v1y)
    if norm < 1e-12:
        v1x, v1y = 1.0, 0.0
        norm = 1.0
    v1x /= norm; v1y /= norm
    v2x, v2y = -v1y, v1x
    proj = [((p[0]-mx)*v1x + (p[1]-my)*v1y, (p[0]-mx)*v2x + (p[1]-my)*v2y) for p in pts]
    xs = [p[0] for p in proj]; ys = [p[1] for p in proj]
    xmin, xmax = min(xs), max(xs)
    ymin, ymax = min(ys), max(ys)
    corners_rot = [(xmin, ymin), (xmax, ymin), (xmax, ymax), (xmin, ymax)]
    corners = []
    for xr, yr in corners_rot:
        wx = mx + xr * v1x + yr * v2x
        wy = my + xr * v1y + yr * v2y
        corners.append((wx, wy))
    return corners

def get_polygon_corners(pts):
    uniq = []
    seen = set()
    for p in pts:
        k = (round(p[0], 6), round(p[1], 6))
        if k not in seen:
            seen.add(k)
            uniq.append((p[0], p[1]))
    if len(uniq) == 0:
        return []
    if len(uniq) == 4:
        return uniq[:]

    hull = convex_hull(uniq)
    def simplify_hull(h):
        if len(h) <= 4:
            return h
        changed = True
        while changed and len(h) > 4:
            changed = False
            n = len(h)
            for i in range(n):
                a = h[(i-1) % n]
                b = h[i]
                c = h[(i+1) % n]
                ang = angle_between(a, b, c)
                if abs(math.pi - ang) < math.radians(6):
                    del h[i]
                    changed = True
                    break
        return h
    hull_s = simplify_hull(hull[:])
    if len(hull_s) == 4:
        return order_points_clockwise(hull_s)
    obb = oriented_bbox_from_points(uniq)
    if not obb:
        xmin, ymin, xmax, ymax = poly_bbox(uniq)
        obb = [(xmin, ymin), (xmax, ymin), (xmax, ymax), (xmin, ymax)]
    diag = math.hypot(obb[1][0] - obb[0][0], obb[2][1] - obb[1][1])
    max_allow = max(1.0, diag * 0.25)
    mapped = []
    for oc in obb:
        best = None
        bd = 1e9
        for up in uniq:
            d = math.hypot(oc[0]-up[0], oc[1]-up[1])
            if d < bd:
                bd = d; best = up
        if bd <= max_allow:
            mapped.append(best)
        else:
            mapped.append(oc)
    return order_points_clockwise(mapped)

def order_points_clockwise(pts):
    cx = sum(p[0] for p in pts) / len(pts)
    cy = sum(p[1] for p in pts) / len(pts)
    def angle(p):
        return math.atan2(p[1]-cy, p[0]-cx)
    pts_sorted = sorted(pts, key=angle, reverse=False)
    pts_sorted = pts_sorted[::-1]
    return pts_sorted

# ---------- 注记偏移，保证注记点不落在 bbox 内 ----------
def get_safe_offset_point(vx, vy, cx, cy, bbox, offset=LABEL_OFFSET):
    dx, dy = vx - cx, vy - cy
    dist = math.hypot(dx, dy)
    if dist < 1e-9:
        return vx + offset, vy
    ux, uy = dx / dist, dy / dist
    lx, ly = vx + ux * offset, vy + uy * offset
    xmin, ymin, xmax, ymax = bbox
    if xmin <= lx <= xmax and ymin <= ly <= ymax:
        lx, ly = vx - ux * offset, vy - uy * offset
    return lx, ly

# ---------- 调用 ZJZB 注记（插件命令） ----------
def run_zjzb(doc, vx, vy, lx, ly):
    cmd = f'(command "ZJZB" (list {vx} {vy} 0.0) (list {lx} {ly} 0.0))\n'
    try:
        doc.SendCommand(cmd)
        time.sleep(0.25)
        pythoncom.PumpWaitingMessages()
    except Exception as e:
        log(f"ZJZB SendCommand 错误: {e}")

# ---------- 绘制矩形 ----------
def draw_rect(doc, cx, cy, buf=BUFFER):
    p1 = (cx - buf, cy - buf)
    p3 = (cx + buf, cy + buf)
    cmd = f'(command "RECTANG" (list {p1[0]} {p1[1]}) (list {p3[0]} {p3[1]}))\n'
    try:
        doc.SendCommand(cmd)
        time.sleep(0.45)
        pythoncom.PumpWaitingMessages()
    except Exception as e:
        log(f"RECTANG SendCommand 错误: {e}")
    return p1, p3

# ---------- 触发 SAVET（只触发） ----------
def trigger_savet_command(doc, p1, p3):
    cmd = f'(command "SAVET" 2 (list {p1[0]} {p1[1]} 0.0) (list {p3[0]} {p3[1]} 0.0))\n'
    try:
        doc.SendCommand(cmd)
        pythoncom.PumpWaitingMessages()
    except Exception as e:
        log(f"SAVET SendCommand 错误: {e}")

# ---------- 后台线程：延时后检测弹窗并输入文件名（仅文件名，不带路径） ----------
class SaveDialogInputThread(threading.Thread):
    def __init__(self, filename, initial_delay=SAVE_DIALOG_INITIAL_DELAY,
                 detect_timeout=SAVE_DIALOG_DETECT_TIMEOUT,
                 template_image=SAVE_DIALOG_IMAGE,
                 fallback_pos=FALLBACK_INPUT_POS):
        super().__init__(daemon=True)
        self.filename = filename
        self.initial_delay = initial_delay
        self.detect_timeout = detect_timeout
        self.template_image = template_image
        self.fallback_pos = fallback_pos
        self.result = False

    def run(self):
        try:
            log(f"保存线程：初始等待 {self.initial_delay}s")
            time.sleep(self.initial_delay)
            start = time.time()
            found = None
            if self.template_image and os.path.exists(self.template_image):
                log(f"保存线程：使用模板检测 {self.template_image}")
                while time.time() - start < self.detect_timeout:
                    try:
                        region = pyautogui.locateOnScreen(self.template_image, confidence=0.7)
                    except TypeError:
                        region = pyautogui.locateOnScreen(self.template_image)
                    except Exception:
                        region = None
                    if region:
                        found = region
                        break
                    time.sleep(0.5)
            else:
                log("保存线程：无模板，直接使用 fallback 坐标")

            if found:
                cx = found.left + found.width/2
                cy = found.top + found.height/2
                log(f"保存线程：检测到模板区域，点击中心 ({int(cx)},{int(cy)})")
                pyautogui.click(cx, cy)
            else:
                fx, fy = self.fallback_pos
                log(f"保存线程：未检测到模板，点击 fallback {self.fallback_pos}")
                pyautogui.click(fx, fy)

            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.05)
            pyautogui.typewrite(str(self.filename), interval=0.02)
            time.sleep(0.05)
            pyautogui.press('enter')
            log(f"保存线程：已输入并回车：{self.filename}")
            self.result = True
            time.sleep(0.8)
        except Exception as e:
            log(f"保存线程异常: {e}")
            self.result = False

# ---------- 按图层删除临时对象（优先） ----------
def delete_temp_layer_entities(msp, layer_name=TEMP_LAYER):
    deleted = 0
    try:
        total = msp.Count
        for i in range(total - 1, -1, -1):
            try:
                ent = msp.Item(i)
                if getattr(ent, "Layer", "") == layer_name:
                    ent.Delete()
                    deleted += 1
                    time.sleep(0.005)
            except Exception:
                continue
    except Exception as e:
        log(f"delete_temp_layer_entities 错误: {e}")
    return deleted

# ---------- 删除新增对象（回退兼容） ----------
def delete_new_entities(msp, before_count):
    after_count = msp.Count
    deleted = 0
    for i in range(after_count-1, before_count-1, -1):
        try:
            ent = msp.Item(i)
            ent.Delete()
            deleted += 1
            time.sleep(0.01)
        except Exception:
            continue
    return deleted

# ---------- 记录当前时间 ----------
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ---------- 日志输出 ----------
def log(msg):
    line = f"[{now()}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

# ---------- 主流程 ----------
def main():
    log("脚本启动：车位批量导出 v16（角点识别优化 + 屏幕保存 + TEMP_LAYER 清理）")
    acad = get_acad_instance()
    try:
        doc = acad.Documents.Open(SRC_FILE)
    except Exception as e:
        log(f"打开 DWG 失败: {e}")
        return
    time.sleep(0.8)
    msp = doc.ModelSpace

    # 确保临时图层存在并切换到临时图层；保存原活动图层以便恢复
    orig_layer = ensure_temp_layer(doc, TEMP_LAYER)

    texts = collect_texts(msp, TEXT_LAYER)
    log(f"发现编号文本: {len(texts)}")
    if not texts:
        log("未找到编号文本，请检查图层与 DWG")
        restore_active_layer(doc, orig_layer)
        return
    if TEST_FIRST_N:
        texts = texts[:TEST_FIRST_N]

    total = len(texts)
    success = 0
    for idx, text in enumerate(texts, start=1):
        try:
            number, tx, ty = text
            log(f"[{idx}/{total}] 处理编号 {number} 文本坐标 ({tx:.3f},{ty:.3f})")

            # 只处理数字编号（避免诸如“装卸车泊位”等文字）
            if not str(number).strip().isdigit():
                log("  非数字编号，跳过")
                continue

            poly_ent, pts = find_car_polyline(msp, text)
            if not poly_ent:
                log("  未找到对应 polyline，跳过")
                continue

            # 记录原色并临时标红；并把 polyline 置前
            try:
                orig_color = getattr(poly_ent, "Color", None)
            except Exception:
                orig_color = None
            try:
                poly_ent.Color = 1
            except Exception:
                log("  无法设置 polyline 颜色")

            # 记录当前对象数（兼容性回退）
            before_count = msp.Count

            # 计算中心与 bbox
            cx = sum(p[0] for p in pts) / len(pts)
            cy = sum(p[1] for p in pts) / len(pts)
            bbox = poly_bbox(pts)

            # 获取角点并注记
            corners = get_polygon_corners(pts)
            log(f"  识别到角点: {corners}")

            # 在 TEMP_LAYER 上执行 ZJZB 注记（因为我们已经切换了活动图层到 TEMP_LAYER）
            for vx, vy in corners:
                lx, ly = get_safe_offset_point(vx, vy, cx, cy, bbox, LABEL_OFFSET)
                run_zjzb(doc, vx, vy, lx, ly)

            # 绘制矩形（将落在 TEMP_LAYER 上）
            p1, p3 = draw_rect(doc, cx, cy, BUFFER)
            log(f"  绘制矩形框 {p1} -> {p3}")

            # 启动保存线程（先等待 initial_delay）
            save_thread = SaveDialogInputThread(
                filename=number,
                initial_delay=SAVE_DIALOG_INITIAL_DELAY,
                detect_timeout=SAVE_DIALOG_DETECT_TIMEOUT,
                template_image=SAVE_DIALOG_IMAGE,
                fallback_pos=FALLBACK_INPUT_POS
            )
            save_thread.start()

            # 触发 SAVET（不等待）
            trigger_savet_command(doc, p1, p3)
            log("  已触发 SAVET 命令")

            # 等待线程完成（最大等待 initial_delay + detect_timeout + 5）
            total_wait = SAVE_DIALOG_INITIAL_DELAY + SAVE_DIALOG_DETECT_TIMEOUT + 5.0
            save_thread.join(timeout=total_wait)
            if save_thread.is_alive():
                log("  保存线程超时或未完成（继续清理）")
            else:
                log(f"  保存线程结果: {save_thread.result}")

            # 首选：删除 TEMP_LAYER 上所有对象（注记 + 矩形）
            deleted_by_layer = delete_temp_layer_entities(msp, TEMP_LAYER)
            log(f"  已删除 {deleted_by_layer} 个 TEMP_LAYER 临时对象")

            # 兼容回退：若仍有多余对象（或插件把注记放在其他图层），根据 before_count 删除新增对象
            deleted_by_count = delete_new_entities(msp, before_count)
            if deleted_by_count:
                log(f"  额外删除 {deleted_by_count} 个新增对象（计数方式）")

            # 恢复 polyline 颜色
            if orig_color is not None:
                try:
                    poly_ent.Color = orig_color
                except Exception:
                    pass

            success += 1
            time.sleep(0.3)

        except Exception as e:
            log(f"处理编号 {text[0]} 时出错: {e}")
            continue

    # 恢复活动图层到原来
    restore_active_layer(doc, orig_layer)

    try:
        doc.Close(False)
    except Exception:
        pass
    log(f"完成：成功导出 {success}/{total} 个编号")

# ========== PyQt5 界面代码 =============
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QTabWidget,
                             QVBoxLayout, QHBoxLayout, QFormLayout, QLabel,
                             QLineEdit, QPushButton, QFileDialog,
                             QDoubleSpinBox, QSpinBox, QMessageBox)
from PyQt5.QtCore import Qt, QThread

class Worker(QThread):
    """
    后台线程：接收参数并运行主流程。
    """
    def __init__(self, src_file, output_dir, text_layer, car_layer, buffer,
                 label_offset, test_first_n, parent=None):
        super().__init__(parent)
        self.src_file = src_file
        self.output_dir = output_dir
        self.text_layer = text_layer
        self.car_layer = car_layer
        self.buffer = buffer
        self.label_offset = label_offset
        self.test_first_n = test_first_n

    def run(self):
        """
        设置全局参数并调用主流程。
        """
        global SRC_FILE, OUTPUT_DIR, LOG_FILE, TEXT_LAYER, CAR_LAYER, BUFFER, LABEL_OFFSET, TEST_FIRST_N
        SRC_FILE = self.src_file
        OUTPUT_DIR = self.output_dir
        LOG_FILE = os.path.join(OUTPUT_DIR, "cad_parking_export_v16.log")
        TEXT_LAYER = self.text_layer
        CAR_LAYER = self.car_layer
        BUFFER = self.buffer
        LABEL_OFFSET = self.label_offset
        # TEST_FIRST_N: 如果为0或None则不限制
        TEST_FIRST_N = self.test_first_n if self.test_first_n != 0 else None
        try:
            os.makedirs(OUTPUT_DIR, exist_ok=True)
        except Exception as e:
            log(f"创建输出目录失败: {e}")
            return
        main()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("车位批量导出")
        self.init_ui()

    def init_ui(self):
        # 创建选项卡
        tabs = QTabWidget()
        self.setCentralWidget(tabs)

        # ----------- 第1页：剪切功能 -----------
        tab1 = QWidget()
        layout1 = QVBoxLayout()

        form_layout = QFormLayout()

        # DWG 文件路径
        self.line_dwg = QLineEdit()
        btn_dwg = QPushButton("浏览...")
        def browse_dwg():
            path, _ = QFileDialog.getOpenFileName(self, "选择 DWG 文件", "", "DWG 文件 (*.dwg);;所有文件 (*)")
            if path:
                self.line_dwg.setText(path)
        btn_dwg.clicked.connect(browse_dwg)
        h_layout1 = QHBoxLayout()
        h_layout1.addWidget(self.line_dwg)
        h_layout1.addWidget(btn_dwg)
        form_layout.addRow("DWG 文件:", h_layout1)

        # 输出文件夹
        self.line_out = QLineEdit()
        btn_out = QPushButton("浏览...")
        def browse_out():
            path = QFileDialog.getExistingDirectory(self, "选择输出目录", "")
            if path:
                self.line_out.setText(path)
        btn_out.clicked.connect(browse_out)
        h_layout2 = QHBoxLayout()
        h_layout2.addWidget(self.line_out)
        h_layout2.addWidget(btn_out)
        form_layout.addRow("输出目录:", h_layout2)

        # 文本图层
        self.line_text_layer = QLineEdit(TEXT_LAYER)
        form_layout.addRow("文本图层:", self.line_text_layer)

        # 车位多边形图层
        self.line_car_layer = QLineEdit(CAR_LAYER)
        form_layout.addRow("车位图层:", self.line_car_layer)

        # 矩形半宽 BUFFER
        self.spin_buffer = QDoubleSpinBox()
        self.spin_buffer.setRange(0.0, 1000.0)
        self.spin_buffer.setValue(BUFFER)
        self.spin_buffer.setDecimals(2)
        form_layout.addRow("矩形半宽 (BUFFER):", self.spin_buffer)

        # 注记偏移 LABEL_OFFSET
        self.spin_offset = QDoubleSpinBox()
        self.spin_offset.setRange(0.0, 1000.0)
        self.spin_offset.setValue(LABEL_OFFSET)
        self.spin_offset.setDecimals(2)
        form_layout.addRow("注记偏移 (LABEL_OFFSET):", self.spin_offset)

        # 处理编号数量限制 TEST_FIRST_N
        self.spin_limit = QSpinBox()
        self.spin_limit.setMinimum(0)
        self.spin_limit.setMaximum(1000000)
        self.spin_limit.setValue(TEST_FIRST_N if TEST_FIRST_N else 0)
        self.spin_limit.setSpecialValueText("无")
        form_layout.addRow("编号数量限制 (0 表示不限):", self.spin_limit)

        layout1.addLayout(form_layout)

        # 运行按钮
        self.btn_run = QPushButton("运行")
        self.btn_run.clicked.connect(self.run_clicked)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_run)
        btn_layout.addStretch()
        layout1.addLayout(btn_layout)

        tab1.setLayout(layout1)
        tabs.addTab(tab1, "剪切功能")

        # ----------- 第2页：功能扩展（空） -----------
        tab2 = QWidget()
        layout2 = QVBoxLayout()
        lbl2 = QLabel("功能扩展待开发")
        lbl2.setAlignment(Qt.AlignCenter)
        layout2.addWidget(lbl2)
        tab2.setLayout(layout2)
        tabs.addTab(tab2, "功能扩展")

    def run_clicked(self):
        # 获取用户输入
        src_file = self.line_dwg.text().strip()
        output_dir = self.line_out.text().strip()
        text_layer = self.line_text_layer.text().strip()
        car_layer = self.line_car_layer.text().strip()
        buffer_val = self.spin_buffer.value()
        offset_val = self.spin_offset.value()
        test_n = self.spin_limit.value()

        # 校验输入
        if not src_file:
            QMessageBox.critical(self, "错误", "请指定 DWG 文件路径！")
            return
        if not os.path.isfile(src_file):
            QMessageBox.critical(self, "错误", "DWG 文件不存在！")
            return
        if not output_dir:
            QMessageBox.critical(self, "错误", "请指定输出目录！")
            return
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法创建输出目录: {e}")
            return
        if not text_layer or not car_layer:
            QMessageBox.critical(self, "错误", "请指定文本图层和车位图层！")
            return

        # 禁用按钮并启动线程
        self.btn_run.setEnabled(False)
        self.worker = Worker(src_file, output_dir, text_layer, car_layer,
                             buffer_val, offset_val, test_n)
        self.worker.finished.connect(self.run_finished)
        self.worker.start()

    def run_finished(self):
        # 恢复按钮并提示完成
        self.btn_run.setEnabled(True)
        QMessageBox.information(self, "完成", "批量导出完成！")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(600, 400)
    window.show()
    sys.exit(app.exec_())
