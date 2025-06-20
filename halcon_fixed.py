# -*- coding: utf-8 -*-
import sys
import os
import clr
import csv
import cv2
import numpy as np
from pathlib import Path
from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import Qt

# 设置环境变量 (修复路径问题) - 在程序一开始就设置
os.environ['HALCONROOT'] = r"C:\Program Files\MVTec\HALCON-24.11-Progress-Steady"
os.environ['PATH'] = r"C:\Program Files\MVTec\HALCON-24.11-Progress-Steady\bin;" + os.environ.get('PATH', '')

# —— 一、配置 DLL 路径
HALCON_DOTNET = r"C:\Program Files\MVTec\HALCON-24.11-Progress-Steady\bin\dotnet35\halcondotnetxl.dll"
ENGINE_DLL = r"C:\Users\USERA\source\repos\TemplateEngineProj\TemplateEngineProj\bin\Debug\TemplateEngineProj.dll"

# 检查DLL文件是否存在
if not os.path.exists(HALCON_DOTNET):
    QtWidgets.QMessageBox.critical(
        None,
        "文件缺失",
        f"找不到HALCON DLL文件:\n{HALCON_DOTNET}\n请检查HALCON安装路径。"
    )
    sys.exit(1)

if not os.path.exists(ENGINE_DLL):
    QtWidgets.QMessageBox.critical(
        None,
        "文件缺失",
        f"找不到自定义引擎DLL文件:\n{ENGINE_DLL}\n请编译C#项目并生成DLL。"
    )
    sys.exit(1)

# —— 二、通过反射获取 CreateTemplate 方法
def get_create_template_method(engine_assembly):
    try:
        # 获取类型
        type_name = "TemplateEngineProj.LoadImages+TemplateEngine"
        template_engine_type = engine_assembly.GetType(type_name)

        if template_engine_type is None:
            # 尝试备用类型名
            type_name = "TemplateEngineProj.TemplateEngine"
            template_engine_type = engine_assembly.GetType(type_name)

            if template_engine_type is None:
                # 列出所有类型
                all_types = "\n".join([t.FullName for t in engine_assembly.GetTypes()])
                raise Exception(f"未找到类型: {type_name}\n程序集中的类型:\n{all_types}")

        # 获取 CreateTemplate 方法
        method = template_engine_type.GetMethod("CreateTemplate")

        if method is None:
            # 尝试备用方法名
            method = template_engine_type.GetMethod("create_template")
            if method is None:
                # 列出所有方法
                methods = "\n".join([m.Name for m in template_engine_type.GetMethods()])
                raise Exception(f"未找到 CreateTemplate 方法\n类型中的方法:\n{methods}")

        print(f"成功获取 CreateTemplate 方法")
        return method
    except Exception as e:
        QtWidgets.QMessageBox.critical(
            None,
            "反射错误",
            f"无法获取 CreateTemplate 方法: {str(e)}\n"
            f"请检查C#代码中的类名和方法名是否正确。"
        )
        return None

# —— 三、完整 PolygonLabel 实现
class PolygonLabel(QtWidgets.QLabel):
    """
    支持缩放、平移和多边形点添加的自定义 QLabel。
    """
    polygon_finished = QtCore.pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.points = []
        self.drawing = False
        self.selected_point = -1
        self.zoom = 1.0            # 缩放比例
        self._pixmap = None        # 原始 QPixmap
        self.image_rect = None     # 图像显示区域
        self.setFocusPolicy(Qt.StrongFocus)
        self.setStyleSheet("background-color: #2D2D30;")  # 深灰色背景

        # 提高框选精准度
        self.point_radius = 8  # 点半径
        self.point_hit_range = 12  # 点检测范围

    def setPixmap(self, pixmap: QtGui.QPixmap):
        # 保存原始 pixmap 并应用当前缩放
        self._pixmap = pixmap
        self.apply_zoom()

    def apply_zoom(self):
        if self._pixmap is None:
            return
        w = int(self._pixmap.width() * self.zoom)
        h = int(self._pixmap.height() * self.zoom)
        scaled = self._pixmap.scaled(w, h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        super().setPixmap(scaled)
        self.image_rect = QtCore.QRect(
            (self.width() - scaled.width()) // 2,
            (self.height() - scaled.height()) // 2,
            scaled.width(),
            scaled.height()
        )

    def mapToImage(self, pos):
        if self._pixmap is None or self.image_rect is None:
            return None
        if not self.image_rect.contains(pos):
            return None
        x_in_image = (pos.x() - self.image_rect.x()) / self.zoom
        y_in_image = (pos.y() - self.image_rect.y()) / self.zoom
        return QtCore.QPoint(int(x_in_image), int(y_in_image))

    def mapToScreen(self, pt):
        if self.image_rect is None:
            return None
        return QtCore.QPoint(
            self.image_rect.x() + int(pt.x() * self.zoom),
            self.image_rect.y() + int(pt.y() * self.zoom)
        )

    def wheelEvent(self, event: QtGui.QWheelEvent):
        delta = event.angleDelta().y()
        factor = 1.25 if delta > 0 else 0.8
        self.zoom = max(0.1, min(self.zoom * factor, 10.0))
        self.apply_zoom()
        event.accept()

    def mousePressEvent(self, event):
        if self._pixmap is None:
            return
        img_pt = self.mapToImage(event.pos())
        if img_pt is None:
            return
        if event.button() == QtCore.Qt.LeftButton:
            # 检查是否点击了现有点
            for i, pt in enumerate(self.points):
                distance = ((img_pt.x() - pt.x())**2 + (img_pt.y() - pt.y())**2)**0.5
                if distance < self.point_hit_range:
                    self.selected_point = i
                    return
            # 添加新点
            self.drawing = True
            self.points.append(img_pt)
            self.selected_point = -1
            self.update()
        elif event.button() == QtCore.Qt.RightButton:
            # 检查是否右键点击了现有点
            for i, pt in enumerate(self.points):
                distance = ((img_pt.x() - pt.x())**2 + (img_pt.y() - pt.y())**2)**0.5
                if distance < self.point_hit_range:
                    del self.points[i]
                    self.selected_point = -1
                    self.update()
                    return
            # 完成多边形绘制
            if self.drawing and len(self.points) >= 3:
                self.polygon_finished.emit([(p.x(), p.y()) for p in self.points])
                self.drawing = False
                self.update()

    def mouseDoubleClickEvent(self, event):
        if self._pixmap is None:
            return
        img_pt = self.mapToImage(event.pos())
        if img_pt and event.button() == QtCore.Qt.LeftButton:
            self.points.append(img_pt)
            self.update()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Delete, Qt.Key_Backspace):
            if 0 <= self.selected_point < len(self.points):
                del self.points[self.selected_point]
                self.selected_point = -1
                self.update()
            elif self.points:
                self.points.pop()
                self.update()
        elif event.key() == Qt.Key_Escape:
            self.selected_point = -1
            self.update()
        super().keyPressEvent(event)

    def mouseMoveEvent(self, event):
        if 0 <= self.selected_point < len(self.points):
            img_pt = self.mapToImage(event.pos())
            if img_pt:
                self.points[self.selected_point] = img_pt
                self.update()
        super().mouseMoveEvent(event)

    def paintEvent(self, event):
        super().paintEvent(event)
        if not self.points:
            return

        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)
        painter.setFont(QtGui.QFont("Arial", 10))

        # 绘制点
        for i, pt in enumerate(self.points):
            screen_pt = self.mapToScreen(pt)
            if not screen_pt:
                continue

            # 设置点颜色（选中的点为红色）
            if i == self.selected_point:
                pen = QtGui.QPen(QtCore.Qt.red)
                brush = QtGui.QBrush(QtCore.Qt.red)
            else:
                pen = QtGui.QPen(QtCore.Qt.cyan)
                brush = QtGui.QBrush(QtCore.Qt.cyan)

            pen.setWidth(2)
            painter.setPen(pen)
            painter.setBrush(brush)
            painter.drawEllipse(screen_pt, self.point_radius, self.point_radius)

            # 绘制点坐标
            painter.setPen(QtGui.QPen(QtCore.Qt.yellow))
            painter.drawText(screen_pt.x() + 10, screen_pt.y() - 8, f"P{i+1}:({pt.x()},{pt.y()})")

        # 绘制多边形连线
        if len(self.points) > 1:
            pen = QtGui.QPen(QtGui.QColor(0, 200, 255))
            pen.setWidth(2)
            painter.setPen(pen)
            path = QtGui.QPainterPath()

            # 第一个点
            start_pt = self.mapToScreen(self.points[0])
            if start_pt:
                path.moveTo(start_pt)

            # 中间点
            for pt in self.points[1:]:
                p = self.mapToScreen(pt)
                if p:
                    path.lineTo(p)

            # 闭合多边形（如果已完成绘制）
            if not self.drawing and start_pt:
                path.lineTo(start_pt)

            painter.drawPath(path)

        # 绘制操作提示
        painter.setPen(QtGui.QPen(QtCore.Qt.white))
        painter.drawText(10, 20, "左键:添加/选中  右键:删除/完成  Delete:删除点  Esc:取消选中  滚轮:缩放")
        painter.end()

# —— 四、主窗口 TemplateMaker (带延迟加载)
class TemplateMaker(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("高精度多边形标注工具 (HALCON .NET 3.5)")
        self.setMinimumSize(800, 600)
        self.setStyleSheet("""
            QWidget {
                background-color: #1E1E1E;
                color: #DCDCDC;
                font-family: Segoe UI;
                font-size: 12pt;
            }
            QPushButton {
                background-color: #0078D7;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                min-width: 100px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #106EBE;
            }
            QPushButton:pressed {
                background-color: #005A9E;
            }
            QPushButton:disabled {
                background-color: #3F3F46;
                color: #A0A0A0;
            }
            QStatusBar {
                background-color: #252526;
                color: #DCDCDC;
                border-top: 1px solid #3F3F46;
            }
        """)

        self.current_image_path = None
        self.image = None
        self.mask = None
        self.last_vis = None
        self.orig_pts = []  # 存储原始图像坐标

        # 创建图像显示区域
        self.label = PolygonLabel()
        self.label.setAlignment(Qt.AlignCenter)
        scroll = QtWidgets.QScrollArea()
        scroll.setWidget(self.label)
        scroll.setWidgetResizable(True)
        scroll.setAlignment(Qt.AlignCenter)
        scroll.setStyleSheet("background-color: #2D2D30;")

        # 创建按钮
        self.btn_load   = QtWidgets.QPushButton("导入图像")
        self.btn_pre    = QtWidgets.QPushButton("预处理图像")
        self.btn_save   = QtWidgets.QPushButton("保存模板")
        self.btn_export = QtWidgets.QPushButton("导出坐标")
        self.btn_clear  = QtWidgets.QPushButton("清除所有点")

        # 初始禁用按钮
        for btn in (self.btn_pre, self.btn_save, self.btn_export, self.btn_clear):
            btn.setEnabled(False)

        # 按钮布局
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.btn_load)
        button_layout.addWidget(self.btn_pre)
        button_layout.addWidget(self.btn_save)
        button_layout.addWidget(self.btn_export)
        button_layout.addWidget(self.btn_clear)
        button_layout.setSpacing(10)
        button_layout.setContentsMargins(5, 5, 5, 5)

        # 主布局
        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.addWidget(scroll)
        main_layout.addLayout(button_layout)
        main_layout.setStretch(0, 1)  # 图像区域占据更多空间

        # 状态栏
        self.status_bar = QtWidgets.QStatusBar()
        self.status_bar.setFont(QtGui.QFont("Arial", 10))
        main_layout.addWidget(self.status_bar)

        # 连接信号和槽
        self.btn_load.clicked.connect(self.load_image)
        self.btn_pre.clicked.connect(self.preprocess_image)
        self.btn_save.clicked.connect(self.save_template)
        self.btn_export.clicked.connect(self.export_coordinates)
        self.btn_clear.clicked.connect(self.clear_points)
        self.label.polygon_finished.connect(self.on_polygon_finished)

        # DLL 相关状态
        self.CreateTemplateMethod = None
        self.dll_loaded = False

    def ensure_dll_loaded(self):
        """确保DLL已加载"""
        if self.dll_loaded:
            return True

        try:
            # 加载 HALCON DLL
            clr.AddReference(HALCON_DOTNET)
            print(f"HALCON DLL 加载成功: {HALCON_DOTNET}")

            # 加载自定义引擎 DLL
            clr.AddReference(ENGINE_DLL)
            print(f"自定义引擎 DLL 加载成功: {ENGINE_DLL}")

            # 获取当前加载的程序集
            assemblies = clr.System.AppDomain.CurrentDomain.GetAssemblies()
            engine_assembly = next((a for a in assemblies if a.GetName().Name == "TemplateEngineProj"), None)
            if engine_assembly is None:
                engine_assembly = clr.System.Reflection.Assembly.LoadFrom(ENGINE_DLL)

            # 获取 CreateTemplate 方法
            self.CreateTemplateMethod = get_create_template_method(engine_assembly)

            if self.CreateTemplateMethod:
                self.dll_loaded = True
                return True
            else:
                return False

        except Exception as e:
            QtWidgets.QMessageBox.critical(
                self,
                "DLL 加载错误",
                f"加载DLL时发生错误:\n{str(e)}\n\n"
                "可能原因:\n"
                "1. .NET Framework 3.5未正确安装\n"
                "2. HALCON许可证问题\n"
                "3. DLL依赖项缺失"
            )
            return False

    def load_image(self):
        """加载图像文件"""
        # 首次操作时加载DLL
        if not self.ensure_dll_loaded():
            return

        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "选择图像", "",
            "图像文件 (*.png *.jpg *.jpeg *.bmp *.tif *.tiff)"
        )
        if not path:
            return

        # 使用OpenCV读取图像
        img = cv2.imread(path)
        if img is None:
            QtWidgets.QMessageBox.critical(self, "加载错误", f"无法加载图片：{path}")
            return

        self.current_image_path = path
        self.image = img
        self.last_vis = None
        self.orig_pts.clear()
        self.label.points.clear()
        self.label.selected_point = -1
        self.label.zoom = 1.0
        self.mask = None

        # 更新显示
        self.update_display(self.image)

        # 更新状态栏
        filename = Path(path).name
        self.status_bar.showMessage(f"已加载图像: {filename}  尺寸: {img.shape[1]}x{img.shape[0]}")

        # 启用相关按钮
        self.btn_pre.setEnabled(True)
        for btn in (self.btn_save, self.btn_export, self.btn_clear):
            btn.setEnabled(False)

    def preprocess_image(self):
        """预处理图像（CLAHE + 伽马校正）"""
        if self.image is None:
            return

        img = self.image.copy()

        # 转换为LAB颜色空间
        lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
        l, a, b = cv2.split(lab)

        # 应用CLAHE
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        cl = clahe.apply(l)

        # 合并通道并转换回BGR
        img = cv2.cvtColor(cv2.merge((cl, a, b)), cv2.COLOR_LAB2BGR)

        # 应用伽马校正
        inv_gamma = 1.0 / 0.8
        table = np.array([(i / 255.0) ** inv_gamma * 255 for i in range(256)], np.uint8)
        img = cv2.LUT(img, table)

        self.image = img
        self.clear_points()
        self.update_display(self.image)
        self.status_bar.showMessage("图像预处理完成 (CLAHE + 伽马校正)")

    def on_polygon_finished(self, poly):
        """多边形绘制完成时的处理"""
        if len(poly) < 3:
            QtWidgets.QMessageBox.warning(self, "无效多边形", "至少需要3个点构成多边形")
            return

        # 存储原始坐标点
        self.orig_pts = [(x, y) for x, y in poly]

        # 创建掩码
        h, w = self.image.shape[:2]
        mask = np.zeros((h, w), np.uint8)
        pts = np.array(self.orig_pts, dtype=np.int32)
        cv2.fillPoly(mask, [pts], 255)
        self.mask = mask

        # 创建带透明度的覆盖层
        overlay = self.image.copy()
        cv2.fillPoly(overlay, [pts], (255, 150, 0))  # 蓝色填充
        alpha = 0.3
        beta = 1.0 - alpha
        self.last_vis = cv2.addWeighted(self.image, beta, overlay, alpha, 0)

        # 更新显示
        self.update_display(self.last_vis)

        # 启用保存和导出按钮
        self.btn_save.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.btn_clear.setEnabled(True)

        # 更新状态栏
        self.status_bar.showMessage(f"多边形已创建，包含 {len(self.orig_pts)} 个点")

    def clear_points(self):
        """清除所有点"""
        self.label.points.clear()
        self.label.selected_point = -1
        self.label.zoom = 1.0
        self.mask = None
        self.last_vis = None
        self.orig_pts.clear()

        # 显示原始图像
        if self.image is not None:
            self.update_display(self.image)

        # 禁用相关按钮
        for btn in (self.btn_save, self.btn_export):
            btn.setEnabled(False)

        self.status_bar.showMessage("已清除所有点")

    def export_coordinates(self):
        """导出坐标到CSV或TXT文件"""
        if not self.orig_pts:
            QtWidgets.QMessageBox.warning(self, "提示", "无坐标可导出")
            return

        # 获取保存路径
        fn, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "保存坐标", "", "CSV (*.csv);;TXT (*.txt);;All Files (*)"
        )
        if not fn:
            return

        try:
            if fn.lower().endswith('.csv'):
                # 保存为CSV格式
                with open(fn, 'w', newline='', encoding='utf-8') as f:
                    w = csv.writer(f)
                    w.writerow(['序号', 'X坐标', 'Y坐标'])
                    for i, (x, y) in enumerate(self.orig_pts, 1):
                        w.writerow([i, x, y])
            else:
                # 保存为TXT格式
                with open(fn, 'w', encoding='utf-8') as f:
                    f.write('序号\tX坐标\tY坐标\n')
                    for i, (x, y) in enumerate(self.orig_pts, 1):
                        f.write(f"{i}\t{x}\t{y}\n")

            # 显示成功消息
            filename = Path(fn).name
            QtWidgets.QMessageBox.information(self, "完成", f"已保存坐标文件: {filename}")
            self.status_bar.showMessage(f"已导出坐标到 {filename}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "导出错误", f"导出坐标时出错: {str(e)}")

    def save_template(self):
        """保存HALCON模板"""
        # 确保DLL已加载
        if not self.ensure_dll_loaded() or not self.CreateTemplateMethod:
            return

        if not self.orig_pts:
            QtWidgets.QMessageBox.warning(self, "提示", "请先创建多边形")
            return

        # 获取保存路径前缀
        output_prefix, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "保存模板前缀", "", "无"
        )
        if not output_prefix:
            return

        try:
            # 硬编码的参数
            circle_row, circle_col = 172, 197
            radius_min, radius_max = 1.0, 10.0
            contour1X = [103, 289, 266, 139, 103]
            contour1Y = [ 88,  83, 253, 258,  88]
            contour2X = [593, 945, 948, 947, 956, 1310, 1320, 599, 597, 593, 593]
            contour2Y = [292, 287, 499, 517, 592,  593,  875, 882, 538, 471, 292]

            # 提取多边形坐标
            corner_rows = [pt[1] for pt in self.orig_pts]
            corner_cols = [pt[0] for pt in self.orig_pts]

            # 显示等待光标
            QtWidgets.QApplication.setOverrideCursor(Qt.WaitCursor)

            # 创建.NET数组
            contour1X_arr = clr.System.Array[clr.System.Int32](contour1X)
            contour1Y_arr = clr.System.Array[clr.System.Int32](contour1Y)
            contour2X_arr = clr.System.Array[clr.System.Int32](contour2X)
            contour2Y_arr = clr.System.Array[clr.System.Int32](contour2Y)
            corner_rows_arr = clr.System.Array[clr.System.Int32](corner_rows)
            corner_cols_arr = clr.System.Array[clr.System.Int32](corner_cols)

            # 通过反射调用C#方法
            self.CreateTemplateMethod.Invoke(None, [
                self.current_image_path,
                output_prefix,
                circle_row,
                circle_col,
                radius_min,
                radius_max,
                contour1X_arr,
                contour1Y_arr,
                contour2X_arr,
                contour2Y_arr,
                corner_rows_arr,
                corner_cols_arr
            ])

            # 恢复光标
            QtWidgets.QApplication.restoreOverrideCursor()

            # 显示成功消息
            QtWidgets.QMessageBox.information(self, "完成", "HALCON 模板已生成！")
            self.status_bar.showMessage("模板生成成功")

        except Exception as e:
            # 确保恢复光标状态
            QtWidgets.QApplication.restoreOverrideCursor()

            # 显示错误信息
            QtWidgets.QMessageBox.critical(
                self,
                "模板生成错误",
                f"生成模板时发生错误:\n{str(e)}\n\n"
                "可能原因:\n"
                "1. 图像路径无效\n"
                "2. 参数超出范围\n"
                "3. HALCON 许可证问题\n"
                "4. 内存不足"
            )
            self.status_bar.showMessage(f"模板生成失败: {str(e)}")

    def update_display(self, img):
        """更新图像显示"""
        try:
            # 将OpenCV图像转换为Qt图像
            rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
            h, w, ch = rgb.shape
            bytes_per_line = ch * w
            qimg = QtGui.QImage(rgb.data, w, h, bytes_per_line, QtGui.QImage.Format_RGB888)
            pix = QtGui.QPixmap.fromImage(qimg)

            # 更新标签
            self.label.setPixmap(pix)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "显示错误", f"更新显示时出错: {str(e)}")

# —— 五、应用程序入口
if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("Fusion")

    # 设置深色主题
    dark_palette = QtGui.QPalette()
    dark_palette.setColor(QtGui.QPalette.Window, QtGui.QColor(30, 30, 30))
    dark_palette.setColor(QtGui.QPalette.WindowText, QtGui.QColor(220, 220, 220))
    dark_palette.setColor(QtGui.QPalette.Base, QtGui.QColor(45, 45, 48))
    dark_palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(30, 30, 30))
    dark_palette.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor(220, 220, 220))
    dark_palette.setColor(QtGui.QPalette.ToolTipText, QtGui.QColor(220, 220, 220))
    dark_palette.setColor(QtGui.QPalette.Text, QtGui.QColor(220, 220, 220))
    dark_palette.setColor(QtGui.QPalette.Button, QtGui.QColor(45, 45, 48))
    dark_palette.setColor(QtGui.QPalette.ButtonText, QtGui.QColor(220, 220, 220))
    dark_palette.setColor(QtGui.QPalette.BrightText, QtGui.QColor(255, 0, 0))
    dark_palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(0, 120, 215))
    dark_palette.setColor(QtGui.QPalette.HighlightedText, QtGui.QColor(255, 255, 255))
    app.setPalette(dark_palette)

    # 创建并显示主窗口
    win = TemplateMaker()
    win.showMaximized()

    # 应用程序退出
    sys.exit(app.exec_())
