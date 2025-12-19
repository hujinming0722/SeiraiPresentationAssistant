import sys
import os
from PyQt6.QtWidgets import (QWidget, QHBoxLayout, QVBoxLayout, QButtonGroup, 
                             QLabel, QFrame, QPushButton, QGridLayout, QStackedWidget,
                             QScrollArea, QSizePolicy)
from PyQt6.QtCore import (Qt, QSize, pyqtSignal, QEvent, QRect, QPropertyAnimation, 
                          QEasingCurve, QAbstractAnimation, QTimer, QSequentialAnimationGroup, 
                          QPoint, QPointF, QThread)
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QColor, QPen, QPainterPath
from qfluentwidgets import (TransparentToolButton, ToolButton, SpinBox,
                            PrimaryPushButton, PushButton, TabWidget,
                            ToolTipFilter, ToolTipPosition, Flyout, FlyoutAnimationType,
                            Pivot, SegmentedWidget, TimePicker, Theme, isDarkTheme,
                            FluentIcon, StrongBodyLabel, TitleLabel, LargeTitleLabel,
                            BodyLabel, CaptionLabel, IndeterminateProgressRing,
                            SmoothScrollArea, FlowLayout)
from qfluentwidgets.components.material import AcrylicFlyout
from .detached_flyout import DetachedFlyoutWindow

def icon_path(name):
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    # 返回上级目录的icons文件夹路径
    return os.path.join(os.path.dirname(base_dir), "icons", name)

def get_icon(name, theme=Theme.DARK):
    path = icon_path(name)
    if not os.path.exists(path):
        return QIcon()
        
    if theme == Theme.LIGHT:
        # Read SVG and replace white with black
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            # Simple replacement of fill="white" or fill="#ffffff" or stroke="white"
            # This is a heuristic, might need adjustment
            content = content.replace('fill="white"', 'fill="#333333"')
            content = content.replace('fill="#ffffff"', 'fill="#333333"')
            content = content.replace('stroke="white"', 'stroke="#333333"')
            content = content.replace('stroke="#ffffff"', 'stroke="#333333"')
            
            pixmap = QPixmap()
            pixmap.loadFromData(content.encode('utf-8'))
            return QIcon(pixmap)
        except:
            return QIcon(path)
    else:
        return QIcon(path)





class PenSettingsFlyout(QWidget):
    color_selected = pyqtSignal(int)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(240, 200)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        title = StrongBodyLabel("墨迹颜色", self)
        layout.addWidget(title)
        
        self.grid_widget = QWidget()
        self.grid_layout = QGridLayout(self.grid_widget)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)
        self.grid_layout.setSpacing(8)
        
        colors = [
            (255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0),
            (255, 0, 255), (0, 255, 255), (0, 0, 0), (255, 255, 255),
            (255, 165, 0), (128, 0, 128)
        ]
        
        for i, rgb in enumerate(colors):
            btn = QPushButton()
            btn.setFixedSize(32, 32)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            
            color_hex = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
            ppt_color = rgb[0] + (rgb[1] << 8) + (rgb[2] << 16)
            
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {color_hex};
                    border: 2px solid rgba(128, 128, 128, 0.2);
                    border-radius: 16px;
                }}
                QPushButton:hover {{
                    border: 2px solid rgba(128, 128, 128, 0.8);
                }}
                QPushButton:pressed {{
                    border: 2px solid white;
                }}
            """)
            
            btn.clicked.connect(lambda checked, c=ppt_color: self.on_color_clicked(c))
            row = i // 5
            col = i % 5
            self.grid_layout.addWidget(btn, row, col)
            
        layout.addWidget(self.grid_widget)
        layout.addStretch()
        
    def on_color_clicked(self, color):
        self.color_selected.emit(color)


class EraserSettingsFlyout(QWidget):
    clear_all_clicked = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(220, 100)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        title = StrongBodyLabel("橡皮擦选项", self)
        layout.addWidget(title)
        
        btn = PrimaryPushButton("清除当前页笔迹", self)
        btn.setFixedSize(188, 32)
        btn.clicked.connect(self.on_clicked)
        layout.addWidget(btn)
        
    def on_clicked(self):
        self.clear_all_clicked.emit()


class SlidePreviewCard(QFrame):
    clicked = pyqtSignal(int)
    
    def __init__(self, index, image_path, parent=None):
        super().__init__(parent)
        self.index = index
        self.setFixedSize(140, 100) 
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet("""
            SlidePreviewCard {
                background-color: transparent;
                border: 1px solid rgba(0, 0, 0, 0.1);
                border-radius: 6px;
            }
            SlidePreviewCard:hover {
                background-color: rgba(0, 0, 0, 0.05);
                border: 1px solid rgba(0, 0, 0, 0.2);
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(4)
        
        self.img_label = QLabel()
        self.img_label.setFixedSize(130, 73) # 16:9 approx
        self.img_label.setStyleSheet("background-color: rgba(128, 128, 128, 0.2); border-radius: 4px;")
        self.img_label.setScaledContents(True)
        if image_path and os.path.exists(image_path):
            self.img_label.setPixmap(QPixmap(image_path))
            
        self.txt_label = CaptionLabel(f"{index}", self)
        self.txt_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addWidget(self.img_label)
        layout.addWidget(self.txt_label)
        
    def mousePressEvent(self, event):
        self.clicked.emit(self.index)


class LoadingOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PyQt6.QtWidgets import QApplication
        self.setGeometry(QApplication.primaryScreen().geometry())
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.ring = IndeterminateProgressRing(self)
        self.ring.setFixedSize(60, 60)
        
        layout.addWidget(self.ring)
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor(0, 0, 0, 80)) # Reduced opacity (more transparent)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRect(self.rect())


class SlideSelectorFlyout(QWidget):
    slide_selected = pyqtSignal(int)
    
    def __init__(self, ppt_app, parent=None):
        super().__init__(parent)
        self.ppt_app = ppt_app
        self.setFixedSize(480, 520)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        
        title = StrongBodyLabel("幻灯片预览", self)
        layout.addWidget(title)
        
        self.scroll = SmoothScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("background-color: transparent; border: none;")
        
        self.container = QWidget()
        self.container.setStyleSheet("background-color: transparent;")
        self.flow = FlowLayout(self.container)
        self.flow.setContentsMargins(0, 0, 0, 0)
        self.flow.setHorizontalSpacing(12)
        self.flow.setVerticalSpacing(12)
        
        self.scroll.setWidget(self.container)
        layout.addWidget(self.scroll)
        
        # We will load slides in background if possible, but here we just trigger load
        QTimer.singleShot(10, self.load_slides)
        
    def get_cache_dir(self, presentation_path):
        import hashlib
        import os
        try:
            path_hash = hashlib.md5(presentation_path.encode('utf-8')).hexdigest()
        except:
            path_hash = "default"
        cache_dir = os.path.join(os.environ['APPDATA'], 'PPTAssistant', 'Cache', path_hash)
        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)
        return cache_dir

    def load_slides(self):
        try:
            presentation = self.ppt_app.ActivePresentation
            slides_count = presentation.Slides.Count
            presentation_path = presentation.FullName
            cache_dir = self.get_cache_dir(presentation_path)
            
            for i in range(1, slides_count + 1):
                slide = presentation.Slides(i)
                thumb_path = os.path.join(cache_dir, f"slide_{i}.jpg")
                
                # Check if exists (assuming cache is handled by BusinessLogic or previous run)
                # If not, we might lag here exporting. 
                # Ideally BusinessLogic pre-caches.
                if not os.path.exists(thumb_path):
                    try:
                        slide.Export(thumb_path, "JPG", 320, 180) 
                    except:
                        pass
                    
                card = SlidePreviewCard(i, thumb_path)
                card.clicked.connect(self.on_card_clicked)
                self.flow.addWidget(card)
                
        except Exception as e:
            print(f"Error loading slides: {e}")
            
    def on_card_clicked(self, index):
        self.slide_selected.emit(index)




class SpotlightOverlay(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        from PyQt6.QtWidgets import QApplication
        self.setGeometry(QApplication.primaryScreen().geometry())
        
        self.selection_rect = QRect()
        self.is_selecting = False
        self.has_selection = False
        self.current_theme = Theme.DARK
        
        # Close button (context aware)
        self.btn_close = TransparentToolButton(FluentIcon.CLOSE, self)
        self.btn_close.setFixedSize(32, 32)
        self.btn_close.setIconSize(QSize(12, 12))
        self.btn_close.hide()
        self.btn_close.clicked.connect(self.close)
        
        self.set_theme(Theme.DARK)

    def set_theme(self, theme):
        self.current_theme = theme
        self.update()
        if theme == Theme.LIGHT:
            # Light mode style for close button (when visible on top of overlay)
            # Since overlay is dark (dimmed screen), close button should probably remain light/white for visibility?
            # Or if it's on top of content...
            # The overlay background is black with alpha. So white icon is best.
            pass
        
        # We'll keep the close button style consistent: white on red or just white?
        # Standard WinUI close is usually red on hover.
        # TransparentToolButton handles this.
        # But since we are on a dark overlay (0,0,0,180), we should force dark theme style for the button
        # so it has white icon.
        self.btn_close.setIcon(FluentIcon.CLOSE)
        self.btn_close.setStyleSheet("""
            TransparentToolButton {
                background-color: #cc0000;
                border: none;
                border-radius: 4px;
                color: white;
            }
            TransparentToolButton:hover {
                background-color: #e60000;
            }
            TransparentToolButton:pressed {
                background-color: #b30000;
            }
        """)

    def mousePressEvent(self, event):
        from PyQt6.QtGui import QMouseEvent
        from PyQt6.QtCore import Qt
        if event.button() == Qt.MouseButton.RightButton:
            self.close()
        elif event.button() == Qt.MouseButton.LeftButton:
            self.selection_rect.setTopLeft(event.pos())
            self.selection_rect.setBottomRight(event.pos())
            self.is_selecting = True
            self.has_selection = False
            self.btn_close.hide()
            self.update()
            
    def mouseMoveEvent(self, event):
        if self.is_selecting:
            self.selection_rect.setBottomRight(event.pos())
            self.update()
            
    def mouseReleaseEvent(self, event):
        from PyQt6.QtCore import Qt, QPoint
        if self.is_selecting:
            self.is_selecting = False
            self.has_selection = True
            # Show close button near top-right of selection
            normalized_rect = self.selection_rect.normalized()
            self.btn_close.move(normalized_rect.topRight() + QPoint(10, -15))
            self.btn_close.show()
            self.update()
            
    def paintEvent(self, event):
        from PyQt6.QtGui import QPainter, QColor, QPen
        from PyQt6.QtCore import Qt
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Fill screen with semi-transparent color based on theme
        if self.current_theme == Theme.LIGHT:
            painter.setBrush(QColor(255, 255, 255, 180))
        else:
            painter.setBrush(QColor(0, 0, 0, 180))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRect(self.rect())
        
        if self.has_selection or self.is_selecting:
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_Clear)
            painter.setBrush(Qt.GlobalColor.transparent)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(self.selection_rect, 10, 10)
            
            # Draw border
            painter.setCompositionMode(QPainter.CompositionMode.CompositionMode_SourceOver)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            pen = QPen(QColor("#00cc7a"))
            pen.setWidth(3)
            painter.setPen(pen)
            painter.drawRoundedRect(self.selection_rect, 10, 10)


class PageNavWidget(QWidget):
    request_slide_jump = pyqtSignal(int)
    prev_clicked = pyqtSignal()
    next_clicked = pyqtSignal()
    
    def __init__(self, parent=None, is_right=False):
        super().__init__(parent)
        self.ppt_app = None 
        self.is_right = is_right
        self.current_theme = Theme.DARK
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        
        inner_layout = QHBoxLayout(self.container)
        inner_layout.setContentsMargins(8, 6, 8, 6) 
        inner_layout.setSpacing(15) 
        
        # Ensure consistent height
        self.container.setMinimumHeight(52)
        self.btn_prev = TransparentToolButton(parent=self)
        self.btn_prev.setFixedSize(36, 36) 
        self.btn_prev.setIconSize(QSize(18, 18))
        self.btn_prev.setToolTip("上一页")
        self.btn_prev.installEventFilter(ToolTipFilter(self.btn_prev, 1000, ToolTipPosition.TOP))
        self.btn_prev.clicked.connect(self.prev_clicked.emit)
        
        self.btn_next = TransparentToolButton(parent=self)
        self.btn_next.setFixedSize(36, 36) 
        self.btn_next.setIconSize(QSize(18, 18))
        self.btn_next.setToolTip("下一页")
        self.btn_next.installEventFilter(ToolTipFilter(self.btn_next, 1000, ToolTipPosition.TOP))
        self.btn_next.clicked.connect(self.next_clicked.emit)

        # Page Info Clickable Area
        self.page_info_widget = QWidget()
        self.page_info_widget.setCursor(Qt.CursorShape.PointingHandCursor)
        self.page_info_widget.installEventFilter(self)
        
        from PyQt6.QtWidgets import QVBoxLayout
        info_layout = QVBoxLayout(self.page_info_widget)
        info_layout.setContentsMargins(10, 0, 10, 0)
        info_layout.setSpacing(2)
        info_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        self.lbl_page_num = QLabel("1/--")
        self.lbl_page_text = QLabel("页码")
        
        info_layout.addWidget(self.lbl_page_num, 0, Qt.AlignmentFlag.AlignCenter)
        info_layout.addWidget(self.lbl_page_text, 0, Qt.AlignmentFlag.AlignCenter)
        
        inner_layout.addWidget(self.btn_prev)
        
        self.line1 = QFrame()
        self.line1.setFrameShape(QFrame.Shape.VLine)
        inner_layout.addWidget(self.line1)
        
        inner_layout.addWidget(self.page_info_widget)
        
        self.line2 = QFrame()
        self.line2.setFrameShape(QFrame.Shape.VLine)
        inner_layout.addWidget(self.line2)
        
        inner_layout.addWidget(self.btn_next)
        
        layout.addWidget(self.container)
        self.setLayout(layout)

        self.setup_click_feedback(self.btn_prev, QSize(18, 18))
        self.setup_click_feedback(self.btn_next, QSize(18, 18))
        
        self.set_theme(Theme.AUTO)
        
    def set_theme(self, theme):
        if theme == Theme.AUTO:
            import qfluentwidgets
            theme = qfluentwidgets.theme()
            
        self.current_theme = theme
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
            subtext_color = "#666666"
            line_color = "rgba(0, 0, 0, 0.1)"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            subtext_color = "#aaaaaa"
            line_color = "rgba(255, 255, 255, 0.2)"
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-bottom: 1px solid {border_color};
                border-radius: 12px;
            }}
            QLabel {{
                font-family: "Segoe UI", "Microsoft YaHei";
                color: {text_color};
            }}
        """)
        
        self.lbl_page_num.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {text_color};")
        self.lbl_page_text.setStyleSheet(f"font-size: 12px; color: {subtext_color};")
        self.line1.setStyleSheet(f"color: {line_color};")
        self.line2.setStyleSheet(f"color: {line_color};")
        
        self.btn_prev.setIcon(get_icon("Previous.svg", theme))
        self.btn_next.setIcon(get_icon("Next.svg", theme))
        
        self.style_nav_btn(self.btn_prev, theme)
        self.style_nav_btn(self.btn_next, theme)

    def style_nav_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(0, 0, 0, 0.05)"
            pressed_bg = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
        else:
            hover_bg = "rgba(255, 255, 255, 0.1)"
            pressed_bg = "rgba(255, 255, 255, 0.2)"
            text_color = "white"
            
        btn.setStyleSheet(f"""
            TransparentToolButton {{
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: {text_color};
            }}
            TransparentToolButton:hover {{
                background-color: {hover_bg};
            }}
            TransparentToolButton:pressed {{
                background-color: {pressed_bg};
            }}
        """)
    
    def setup_click_feedback(self, btn, base_size):
        anim = QPropertyAnimation(btn, b"iconSize", self)
        anim.setDuration(120)
        anim.setEasingCurve(QEasingCurve.Type.InOutQuad)

        def on_pressed():
            if anim.state() == QAbstractAnimation.State.Running:
                anim.stop()
            shrink = QSize(int(base_size.width() * 0.85), int(base_size.height() * 0.85))
            anim.setStartValue(base_size)
            anim.setKeyValueAt(0.5, shrink)
            anim.setEndValue(base_size)
            anim.start()

        btn.pressed.connect(on_pressed)
    
    def eventFilter(self, obj, event):
        if obj == self.page_info_widget:
            if event.type() == QEvent.Type.MouseButtonDblClick:
                return super().eventFilter(obj, event)
            elif event.type() == QEvent.Type.MouseButtonRelease:
                self.show_slide_selector()
                return True
        return super().eventFilter(obj, event)

    def show_slide_selector(self):
        if not self.ppt_app:
            return
            
        view = SlideSelectorFlyout(self.ppt_app)
        view.slide_selected.connect(self.request_slide_jump.emit)
        
        # Ensure view has a background
        if self.current_theme == Theme.LIGHT:
            bg_color = "#f3f3f3"
        else:
            bg_color = "#202020"
        view.setStyleSheet(f"SlideSelectorFlyout {{ background-color: {bg_color}; border-radius: 8px; }}")

        win = DetachedFlyoutWindow(view, self)
        view.slide_selected.connect(win.close)
        win.show_at(self.page_info_widget)

    def update_page(self, current, total):
        self.lbl_page_num.setText(f"{current}/{total}")
    
    def apply_settings(self):
        self.btn_prev.setToolTip("上一页")
        self.btn_next.setToolTip("下一页")
        self.lbl_page_text.setText("页码")
        self.style_nav_btn(self.btn_prev, self.current_theme)
        self.style_nav_btn(self.btn_next, self.current_theme)


class ToolBarWidget(QWidget):
    request_spotlight = pyqtSignal()
    request_pointer_mode = pyqtSignal(int)
    request_pen_color = pyqtSignal(int)
    request_clear_ink = pyqtSignal()
    request_exit = pyqtSignal()
    request_timer = pyqtSignal()
    request_board_in_board = pyqtSignal()
    
    def __init__(self):
        super().__init__()
        self.current_theme = Theme.DARK
        self.was_checked = False
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.container = QWidget()
        self.container.setObjectName("Container")
        self.container.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        container_layout = QHBoxLayout(self.container)
        container_layout.setContentsMargins(8, 6, 8, 6) 
        container_layout.setSpacing(12) 
        
        # Ensure consistent height
        self.container.setMinimumHeight(56)
        
        self.group = QButtonGroup(self)
        self.group.setExclusive(True)
        
        self.btn_arrow = self.create_tool_btn("选择", "Mouse.svg")
        self.btn_arrow.clicked.connect(lambda: self.request_pointer_mode.emit(1))
        
        self.btn_pen = self.create_tool_btn("笔", "Pen.svg")
        self.btn_pen.clicked.connect(lambda: self.request_pointer_mode.emit(2))
        
        self.btn_eraser = self.create_tool_btn("橡皮", "Eraser.svg")
        self.btn_eraser.clicked.connect(lambda: self.request_pointer_mode.emit(5))
        
        self.btn_clear = self.create_action_btn("一键清除", "Clear.svg")
        self.btn_clear.clicked.connect(self.request_clear_ink.emit)
        
        self.group.addButton(self.btn_arrow)
        self.group.addButton(self.btn_pen)
        self.group.addButton(self.btn_eraser)
        
        self.btn_spotlight = self.create_action_btn("聚焦", "Select.svg")
        self.btn_spotlight.clicked.connect(self.request_spotlight.emit)
        
        self.btn_board = self.create_action_btn("板中板", "board-in-board.svg")
        self.btn_board.clicked.connect(self.request_board_in_board.emit)
        
        self.btn_timer = self.create_action_btn("计时器", "timer.svg")
        self.btn_timer.clicked.connect(self.request_timer.emit)

        self.btn_exit = self.create_action_btn("结束放映", "Minimaze.svg")
        self.btn_exit.clicked.connect(self.request_exit.emit)
        
        container_layout.addWidget(self.btn_arrow)
        container_layout.addWidget(self.btn_pen)
        container_layout.addWidget(self.btn_eraser)
        container_layout.addWidget(self.btn_clear)
        
        self.line1 = QFrame()
        self.line1.setFrameShape(QFrame.Shape.VLine)
        container_layout.addWidget(self.line1)
        
        container_layout.addWidget(self.btn_spotlight)
        container_layout.addWidget(self.btn_board)

        self.line2 = QFrame()
        self.line2.setFrameShape(QFrame.Shape.VLine)
        container_layout.addWidget(self.line2)
        
        container_layout.addWidget(self.btn_timer)
        container_layout.addWidget(self.btn_exit)
        
        layout.addWidget(self.container)
        self.setLayout(layout)
        
        self.btn_arrow.setChecked(True)
        
        # Install event filter to detect second click for expansion
        self.btn_pen.installEventFilter(self)
        self.btn_eraser.installEventFilter(self)

        self.setup_click_feedback(self.btn_arrow, QSize(20, 20))
        self.setup_click_feedback(self.btn_pen, QSize(20, 20))
        self.setup_click_feedback(self.btn_eraser, QSize(20, 20))
        self.setup_click_feedback(self.btn_clear, QSize(20, 20))
        self.setup_click_feedback(self.btn_spotlight, QSize(20, 20))
        self.setup_click_feedback(self.btn_timer, QSize(20, 20))
        self.setup_click_feedback(self.btn_exit, QSize(20, 20))
        
        self.set_theme(Theme.AUTO)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.MouseButtonPress:
            if event.button() == Qt.MouseButton.RightButton:
                if obj == self.btn_pen:
                    self.show_pen_settings()
                    return True
                elif obj == self.btn_eraser:
                    self.show_eraser_settings()
                    return True
            elif event.button() == Qt.MouseButton.LeftButton:
                # If clicking Pen button while it is already checked -> Show settings
                if obj == self.btn_pen and self.btn_pen.isChecked():
                    self.show_pen_settings()
                    return True

        elif event.type() == QEvent.Type.MouseButtonRelease:
            pass
                    
        return super().eventFilter(obj, event)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)

    def show_pen_settings(self):
        view = PenSettingsFlyout(self)
        view.color_selected.connect(self.request_pen_color.emit)
        
        view.setStyleSheet(f"background-color: {self.get_flyout_bg_color()}; border-radius: 8px;")
        
        win = DetachedFlyoutWindow(view, self)
        view.color_selected.connect(win.close)
        win.show_at(self.btn_pen)

    def show_eraser_settings(self):
        view = EraserSettingsFlyout(self)
        view.clear_all_clicked.connect(self.request_clear_ink.emit)
        
        view.setStyleSheet(f"background-color: {self.get_flyout_bg_color()}; border-radius: 8px;")
        
        win = DetachedFlyoutWindow(view, self)
        view.clear_all_clicked.connect(win.close)
        win.show_at(self.btn_eraser)
        
    def get_flyout_bg_color(self):
        if self.current_theme == Theme.LIGHT:
            return "#f3f3f3"
        else:
            return "#202020"

    def set_theme(self, theme):
        if theme == Theme.AUTO:
            import qfluentwidgets
            theme = qfluentwidgets.theme()
            
        self.current_theme = theme
        
        # Update container style
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
            line_color = "rgba(0, 0, 0, 0.1)"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            line_color = "rgba(255, 255, 255, 0.2)"
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-bottom: 1px solid {border_color};
                border-radius: 12px;
            }}
        """)
        
        self.line1.setStyleSheet(f"color: {line_color};")
        self.line2.setStyleSheet(f"color: {line_color};")
        
        # Update icons and button styles
        for btn, icon_name in [
            (self.btn_arrow, "Mouse.svg"),
            (self.btn_pen, "Pen.svg"),
            (self.btn_eraser, "Eraser.svg"),
            (self.btn_clear, "Clear.svg"),
            (self.btn_spotlight, "Select.svg"),
            (self.btn_board, "board-in-board.svg"),
            (self.btn_timer, "timer.svg"),
            (self.btn_exit, "Minimaze.svg")
        ]:
            btn.setIcon(get_icon(icon_name, theme))
            self.style_tool_btn(btn, theme) if btn.isCheckable() else self.style_action_btn(btn, theme)

    def create_tool_btn(self, text, icon_name):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(get_icon(icon_name, self.current_theme))
        btn.setFixedSize(40, 40) 
        btn.setIconSize(QSize(20, 20))
        btn.setCheckable(True)
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        # Style will be set in set_theme
        return btn
        
    def create_action_btn(self, text, icon_name):
        btn = TransparentToolButton(parent=self)
        btn.setIcon(get_icon(icon_name, self.current_theme))
        btn.setFixedSize(40, 40)
        btn.setIconSize(QSize(20, 20))
        btn.setToolTip(text)
        btn.installEventFilter(ToolTipFilter(btn, 1000, ToolTipPosition.TOP))
        # Style will be set in set_theme
        return btn
    
    def style_tool_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(0, 0, 0, 0.05)"
            checked_bg = "rgba(0, 0, 0, 0.05)"
            text_color = "#333333"
            border_bottom = "#00cc7a"
        else:
            hover_bg = "rgba(255, 255, 255, 0.1)"
            checked_bg = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            border_bottom = "#00cc7a"
            
        btn.setStyleSheet(f"""
            TransparentToolButton {{
                border-radius: 6px;
                border: none;
                background-color: transparent;
                color: {text_color};
                margin-bottom: 2px;
            }}
            TransparentToolButton:hover {{
                background-color: {hover_bg};
            }}
            TransparentToolButton:checked {{
                background-color: {checked_bg};
                color: {text_color};
                border-bottom: 3px solid {border_bottom};
                border-bottom-left-radius: 2px;
                border-bottom-right-radius: 2px;
            }}
            TransparentToolButton:checked:hover {{
                background-color: {hover_bg};
            }}
        """)
    
    def style_action_btn(self, btn, theme):
        if theme == Theme.LIGHT:
            hover_bg = "rgba(0, 0, 0, 0.05)"
            pressed_bg = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
        else:
            hover_bg = "rgba(255, 255, 255, 0.1)"
            pressed_bg = "rgba(255, 255, 255, 0.2)"
            text_color = "white"
            
        # Special case for exit button hover color if needed, but keeping it simple for now
        # If it's exit button, we might want red hover. 
        # But the previous code had specific style for btn_exit. 
        # Let's handle btn_exit specifically in set_theme loop or check here.
        
        if btn == self.btn_exit:
             btn.setStyleSheet(f"""
                TransparentToolButton {{
                    border-radius: 6px;
                    border: none;
                    background-color: transparent;
                    color: {text_color};
                }}
                TransparentToolButton:hover {{
                    background-color: rgba(255, 50, 50, 0.3);
                }}
                TransparentToolButton:pressed {{
                    background-color: rgba(255, 50, 50, 0.5);
                }}
            """)
        else:
            btn.setStyleSheet(f"""
                TransparentToolButton {{
                    border-radius: 6px;
                    border: none;
                    background-color: transparent;
                    color: {text_color};
                }}
                TransparentToolButton:hover {{
                    background-color: {hover_bg};
                }}
                TransparentToolButton:pressed {{
                    background-color: {pressed_bg};
                }}
            """)

    def setup_click_feedback(self, btn, base_size):
        anim = QPropertyAnimation(btn, b"iconSize", self)
        anim.setDuration(120)
        anim.setEasingCurve(QEasingCurve.Type.InOutQuad)

        def on_pressed():
            if anim.state() == QAbstractAnimation.State.Running:
                anim.stop()
            shrink = QSize(int(base_size.width() * 0.85), int(base_size.height() * 0.85))
            anim.setStartValue(base_size)
            anim.setKeyValueAt(0.5, shrink)
            anim.setEndValue(base_size)
            anim.start()

        btn.pressed.connect(on_pressed)


class TimerWindow(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_theme = Theme.DARK
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.resize(340, 280)
        
        self.up_seconds = 0
        self.up_running = False
        self.down_total_seconds = 0
        self.down_remaining = 0
        self.down_running = False
        self.sound_effect = None
        self.drag_pos = None
        
        # Main Container
        self.container = QWidget(self)
        self.container.setObjectName("Container")
        
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.addWidget(self.container)
        
        self.layout = QVBoxLayout(self.container)
        self.layout.setContentsMargins(20, 20, 20, 20)
        self.layout.setSpacing(15)
        
        # Header
        header_layout = QHBoxLayout()
        self.title_label = TitleLabel("计时工具", self)
        self.close_btn = TransparentToolButton(FluentIcon.CLOSE, self)
        self.close_btn.setFixedSize(32, 32)
        self.close_btn.setIconSize(QSize(12, 12))
        self.close_btn.clicked.connect(self.close)
        header_layout.addWidget(self.title_label)
        header_layout.addStretch()
        header_layout.addWidget(self.close_btn)
        self.layout.addLayout(header_layout)
        
        # Segmented Widget for switching modes
        self.pivot = SegmentedWidget(self)
        self.pivot.addItem("up", "正计时")
        self.pivot.addItem("down", "倒计时")
        self.pivot.currentItemChanged.connect(self.on_pivot_changed)
        self.layout.addWidget(self.pivot)
        
        # Content Stack
        self.stack = QStackedWidget(self)
        self.up_page = QWidget()
        self.down_page = QWidget()
        self.completed_page = QWidget()
        self.setup_up_page()
        self.setup_down_page()
        self.setup_completed_page()
        self.stack.addWidget(self.up_page)
        self.stack.addWidget(self.down_page)
        self.stack.addWidget(self.completed_page)
        self.layout.addWidget(self.stack)
        
        self.init_timers()
        self.init_sound()
        self.set_theme(Theme.DARK)
        
        self.pivot.setCurrentItem("up")

    def on_pivot_changed(self, route_key):
        if route_key == "up":
            self.stack.setCurrentWidget(self.up_page)
        else:
            self.stack.setCurrentWidget(self.down_page)

    def setup_up_page(self):
        layout = QVBoxLayout(self.up_page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        self.up_label = QLabel("00:00")
        self.up_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addStretch()
        layout.addWidget(self.up_label)
        layout.addStretch()
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        self.up_start_btn = PrimaryPushButton("开始", self.up_page)
        self.up_start_btn.setFixedWidth(100)
        self.up_reset_btn = PushButton("重置", self.up_page)
        self.up_reset_btn.setFixedWidth(100)
        
        self.up_start_btn.clicked.connect(self.toggle_up)
        self.up_reset_btn.clicked.connect(self.reset_up)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.up_start_btn)
        btn_layout.addWidget(self.up_reset_btn)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)
        layout.addSpacing(10)

    def setup_completed_page(self):
        layout = QVBoxLayout(self.completed_page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        self.completed_label = QLabel("倒计时已结束")
        self.completed_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        layout.addStretch()
        layout.addWidget(self.completed_label)
        layout.addSpacing(20)
        
        self.back_btn = PrimaryPushButton("返回", self.completed_page)
        self.back_btn.setFixedWidth(120)
        self.back_btn.clicked.connect(self.on_completed_back)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(self.back_btn)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)
        layout.addStretch()

    def on_completed_back(self):
        self.reset_down()
        self.stack.setCurrentWidget(self.down_page)

    def setup_down_page(self):
        layout = QVBoxLayout(self.down_page)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(20)
        
        # Time input area
        self.input_widget = QWidget()
        input_layout = QHBoxLayout(self.input_widget)
        input_layout.setContentsMargins(0, 0, 0, 0)
        input_layout.setSpacing(10)
        
        self.down_min_spin = SpinBox()
        self.down_min_spin.setRange(0, 999)
        self.down_min_spin.setSuffix(" 分")
        self.down_min_spin.setFixedWidth(110)
        self.down_min_spin.setValue(5) # Default 5 mins
        
        self.down_sec_spin = SpinBox()
        self.down_sec_spin.setRange(0, 59)
        self.down_sec_spin.setSuffix(" 秒")
        self.down_sec_spin.setFixedWidth(110)
        
        input_layout.addStretch()
        input_layout.addWidget(self.down_min_spin)
        input_layout.addWidget(self.down_sec_spin)
        input_layout.addStretch()
        
        self.down_label = QLabel("00:00")
        self.down_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.down_label.hide()
        
        layout.addStretch()
        layout.addWidget(self.input_widget)
        layout.addWidget(self.down_label)
        layout.addStretch()
        
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(15)
        self.down_start_btn = PrimaryPushButton("开始", self.down_page)
        self.down_start_btn.setFixedWidth(100)
        self.down_reset_btn = PushButton("重置", self.down_page)
        self.down_reset_btn.setFixedWidth(100)
        
        self.down_start_btn.clicked.connect(self.toggle_down)
        self.down_reset_btn.clicked.connect(self.reset_down)
        
        btn_layout.addStretch()
        btn_layout.addWidget(self.down_start_btn)
        btn_layout.addWidget(self.down_reset_btn)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)
        layout.addSpacing(10)

    def set_theme(self, theme):
        self.current_theme = theme
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 248)"
            border_color = "rgba(0, 0, 0, 0.1)"
            text_color = "#333333"
            self.title_label.setTextColor("#333333", "#333333")
        else:
            bg_color = "rgba(32, 32, 32, 248)"
            border_color = "rgba(255, 255, 255, 0.1)"
            text_color = "white"
            self.title_label.setTextColor("white", "white")
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 12px;
            }}
        """)
        
        font_style = f"font-size: 56px; font-weight: bold; color: {text_color}; font-family: 'Segoe UI', 'Microsoft YaHei';"
        self.up_label.setStyleSheet(font_style)
        self.down_label.setStyleSheet(font_style)
        
        completed_style = f"font-size: 24px; font-weight: bold; color: {text_color}; font-family: 'Segoe UI', 'Microsoft YaHei';"
        if hasattr(self, 'completed_label'):
            self.completed_label.setStyleSheet(completed_style)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton and self.drag_pos:
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()
            
    def mouseReleaseEvent(self, event):
        self.drag_pos = None

    def shake_window(self):
        self.shake_anim = QSequentialAnimationGroup(self)
        original_pos = self.pos()
        offset = 10
        
        for _ in range(2):
            anim1 = QPropertyAnimation(self, b"pos")
            anim1.setDuration(50)
            anim1.setStartValue(original_pos)
            anim1.setEndValue(original_pos + QPoint(offset, 0))
            anim1.setEasingCurve(QEasingCurve.Type.InOutQuad)
            
            anim2 = QPropertyAnimation(self, b"pos")
            anim2.setDuration(50)
            anim2.setStartValue(original_pos + QPoint(offset, 0))
            anim2.setEndValue(original_pos - QPoint(offset, 0))
            anim2.setEasingCurve(QEasingCurve.Type.InOutQuad)
            
            anim3 = QPropertyAnimation(self, b"pos")
            anim3.setDuration(50)
            anim3.setStartValue(original_pos - QPoint(offset, 0))
            anim3.setEndValue(original_pos)
            anim3.setEasingCurve(QEasingCurve.Type.InOutQuad)
            
            self.shake_anim.addAnimation(anim1)
            self.shake_anim.addAnimation(anim2)
            self.shake_anim.addAnimation(anim3)
            
        self.shake_anim.start()

    def init_timers(self):
        self.up_timer = QTimer(self)
        self.up_timer.setInterval(1000)
        self.up_timer.timeout.connect(self.update_up)
        self.down_timer = QTimer(self)
        self.down_timer.setInterval(1000)
        self.down_timer.timeout.connect(self.update_down)

    def init_sound(self):
        try:
            from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
            from PyQt6.QtCore import QUrl
        except Exception:
            self.player = None
            return
            
        self.player = QMediaPlayer(self)
        self.audio_output = QAudioOutput(self)
        self.player.setAudioOutput(self.audio_output)
        
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base_dir, "timer_ring.ogg")
        
        if os.path.exists(path):
            self.player.setSource(QUrl.fromLocalFile(path))
            self.audio_output.setVolume(1.0)

    def play_ring(self):
        if hasattr(self, 'player') and self.player:
            self.player.stop()
            self.player.play()

    def format_time(self, seconds):
        h = seconds // 3600
        m = (seconds % 3600) // 60
        s = seconds % 60
        if h > 0:
            return f"{h:02d}:{m:02d}:{s:02d}"
        return f"{m:02d}:{s:02d}"

    def toggle_up(self):
        if not self.up_running:
            self.up_timer.start()
            self.up_running = True
            self.up_start_btn.setText("暂停")
        else:
            self.up_timer.stop()
            self.up_running = False
            self.up_start_btn.setText("开始")

    def reset_up(self):
        self.up_timer.stop()
        self.up_running = False
        self.up_seconds = 0
        self.up_label.setText("00:00")
        self.up_start_btn.setText("开始")

    def update_up(self):
        self.up_seconds += 1
        self.up_label.setText(self.format_time(self.up_seconds))

    def toggle_down(self):
        if not self.down_running:
            if self.down_remaining == 0:
                minutes = self.down_min_spin.value()
                seconds = self.down_sec_spin.value()
                self.down_total_seconds = minutes * 60 + seconds
                self.down_remaining = self.down_total_seconds
            
            if self.down_remaining > 0:
                self.down_timer.start()
                self.down_running = True
                self.down_start_btn.setText("暂停")
                self.input_widget.hide()
                self.down_label.show()
        else:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("继续")

    def reset_down(self):
        self.down_timer.stop()
        self.down_running = False
        self.down_remaining = 0
        self.down_total_seconds = 0
        self.down_label.setText("00:00")
        self.down_start_btn.setText("开始")
        self.input_widget.show()
        self.down_label.hide()

    def update_down(self):
        if self.down_remaining > 0:
            self.down_remaining -= 1
            self.down_label.setText(self.format_time(self.down_remaining))
            
            if self.down_remaining == 0:
                self.down_timer.stop()
                self.down_running = False
                self.down_start_btn.setText("开始")
                self.play_ring()
                self.stack.setCurrentWidget(self.completed_page)


class DrawingCanvas(QWidget):
    """专用的绘画画布组件"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.strokes = []
        self.current_path = None
        self.drawing = False
        self.pen_color = QColor(255, 0, 0)  # 默认红色
        self.pen_width = 3
        
        # 设置最小大小
        self.setMinimumSize(350, 250)
        
    def set_strokes(self, strokes):
        """设置笔画列表"""
        self.strokes = strokes
        
    def set_pen_color(self, color):
        """设置画笔颜色"""
        self.pen_color = color
        self.update()
        
    def set_pen_width(self, width):
        """设置画笔宽度"""
        self.pen_width = width
        self.update()
        
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drawing = True
            self.current_path = QPainterPath()
            self.current_path.moveTo(QPointF(event.pos()))
            self.update()
            
    def mouseMoveEvent(self, event):
        if self.drawing and event.buttons() == Qt.MouseButton.LeftButton:
            if self.current_path:
                self.current_path.lineTo(QPointF(event.pos()))
                self.update()
            
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.drawing:
            self.drawing = False
            if self.current_path:
                self.strokes.append(self.current_path)
                self.current_path = None
            self.update()
            
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # 绘制背景
        painter.fillRect(self.rect(), Qt.GlobalColor.white)
        
        # 绘制历史笔画
        for path in self.strokes:
            painter.setPen(QPen(self.pen_color, self.pen_width, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            painter.drawPath(path)
        
        # 绘制当前笔画
        if self.drawing and self.current_path:
            painter.setPen(QPen(self.pen_color, self.pen_width, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin))
            painter.drawPath(self.current_path)
            
    def clear(self):
        """清除画布"""
        self.strokes.clear()
        self.current_path = None
        self.drawing = False
        self.update()


class BoardInBoardWindow(QWidget):
    cleared = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_theme = Theme.DARK
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setMinimumSize(400, 300)
        self.resize(600, 400)
        
        self.drag_pos = None
        self.resizing = False
        self.resize_dir = None
        self.lastMousePos = None
        
        self.pen_color = QColor(255, 0, 0)  # 默认红色
        self.pen_width = 3
        
        self.drawing = False
        self.last_point = QPoint()
        self.strokes = []  # 存储绘画笔画
        self.current_path = None
        
        # 绑定绘画事件到canvas_bg
        if hasattr(self, 'canvas_bg'):
            self.canvas_bg.set_strokes(self.strokes)
            self.canvas_bg.set_pen_color(self.pen_color)
            self.canvas_bg.set_pen_width(self.pen_width)
        
        self.init_ui()
        
    def init_ui(self):
        # 主容器
        self.container = QWidget(self)
        self.container.setObjectName("Container")
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.container)
        
        # 标题栏
        self.title_bar = QWidget(self.container)
        self.title_bar.setObjectName("TitleBar")
        self.title_bar.setFixedHeight(40)
        self.title_bar.setCursor(Qt.CursorShape.SizeAllCursor)
        
        title_layout = QHBoxLayout(self.title_bar)
        
        # 标题
        self.title_label = StrongBodyLabel("板中板", self.title_bar)
        title_layout.addWidget(self.title_label)
        
        title_layout.addStretch()
        
        # 清除按钮
        self.clear_btn = TransparentToolButton(parent=self.title_bar)
        self.clear_btn.setFixedSize(30, 30)
        self.clear_btn.setIconSize(QSize(16, 16))
        self.clear_btn.setToolTip("清除")
        self.clear_btn.clicked.connect(self.clear_canvas)
        
        # 关闭按钮
        self.close_btn = TransparentToolButton(FluentIcon.CLOSE, self.title_bar)
        self.close_btn.setFixedSize(30, 30)
        self.close_btn.setIconSize(QSize(12, 12))
        self.close_btn.clicked.connect(self.close)
        
        title_layout.addWidget(self.clear_btn)
        title_layout.addWidget(self.close_btn)
        
        # 绘画区域 - 直接使用DrawingCanvas作为主绘画区域
        self.canvas_bg = DrawingCanvas(self)
        self.canvas_bg.setObjectName("CanvasBg")
        self.canvas_bg.setCursor(Qt.CursorShape.CrossCursor)
        self.canvas_bg.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        # 设置画布布局
        canvas_layout = QVBoxLayout(self.container)
        canvas_layout.setContentsMargins(0, 0, 0, 0)
        canvas_layout.setSpacing(0)
        canvas_layout.addWidget(self.title_bar)
        canvas_layout.addWidget(self.canvas_bg, 1)  # 让画布占据剩余空间
        
        # 绑定绘画数据
        self.canvas_bg.set_strokes(self.strokes)
        self.canvas_bg.set_pen_color(self.pen_color)
        self.canvas_bg.set_pen_width(self.pen_width)
        
        # 初始化主题
        self.set_theme(Theme.AUTO)
        
        # 只为窗口和标题栏安装事件过滤器
        self.installEventFilter(self)
        self.title_bar.installEventFilter(self)
        
    def set_theme(self, theme):
        if theme == Theme.AUTO:
            import qfluentwidgets
            theme = qfluentwidgets.theme()
            
        self.current_theme = theme
        
        if theme == Theme.LIGHT:
            bg_color = "rgba(255, 255, 255, 240)"
            border_color = "rgba(0, 0, 0, 0.2)"
            text_color = "#333333"
            canvas_bg = "white"
            title_bg = "rgba(248, 248, 248, 200)"
        else:
            bg_color = "rgba(30, 30, 30, 240)"
            border_color = "rgba(255, 255, 255, 0.2)"
            text_color = "white"
            canvas_bg = "rgba(20, 20, 20, 255)"
            title_bg = "rgba(40, 40, 40, 200)"
            
        self.container.setStyleSheet(f"""
            QWidget#Container {{
                background-color: {bg_color};
                border: 1px solid {border_color};
                border-radius: 8px;
            }}
            QWidget#TitleBar {{
                background-color: {title_bg};
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                border-bottom: 1px solid {border_color};
            }}
            QWidget#Canvas {{
                background-color: {canvas_bg};
                border-bottom-left-radius: 8px;
                border-bottom-right-radius: 8px;
            }}
        """)
        
        self.title_label.setStyleSheet(f"color: {text_color}; font-weight: bold;")
        
        # 设置按钮图标
        self.clear_btn.setIcon(get_icon("Clear.svg", theme))
        
    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.MouseButtonPress:
            if event.button() == Qt.MouseButton.LeftButton:
                mouse_pos = event.globalPosition().toPoint()
                
                # 只在标题栏区域处理拖拽
                if obj == self.title_bar:
                    if self.is_on_title_bar(mouse_pos):
                        self.drag_pos = mouse_pos - self.frameGeometry().topLeft()
                        return True
                    else:
                        # 检查是否在边缘进行大小调整
                        resize_dir = self.get_resize_direction(mouse_pos)
                        if resize_dir:
                            self.resizing = True
                            self.resize_dir = resize_dir
                            self.lastMousePos = mouse_pos
                            return True
                # 画布区域的鼠标事件由mousePressEvent处理
                        
        elif event.type() == QEvent.Type.MouseMove:
            if self.drag_pos:
                if event.buttons() == Qt.MouseButton.LeftButton:
                    self.move(event.globalPosition().toPoint() - self.drag_pos)
                    return True
            elif self.resizing and self.lastMousePos:
                if event.buttons() == Qt.MouseButton.LeftButton:
                    self.resize_window(event.globalPosition().toPoint())
                    return True
            else:
                # 更新鼠标形状
                mouse_pos = event.globalPosition().toPoint()
                resize_dir = self.get_resize_direction(mouse_pos)
                if resize_dir:
                    if 'left' in resize_dir or 'right' in resize_dir:
                        self.setCursor(Qt.CursorShape.SizeHorCursor)
                    elif 'top' in resize_dir or 'bottom' in resize_dir:
                        self.setCursor(Qt.CursorShape.SizeVerCursor)
                    elif 'corner' in resize_dir:
                        self.setCursor(Qt.CursorShape.SizeFDiagCursor)
                    else:
                        self.setCursor(Qt.CursorShape.SizeAllCursor)
                else:
                    self.setCursor(Qt.CursorShape.SizeAllCursor)
                    
        elif event.type() == QEvent.Type.MouseButtonRelease:
            self.drag_pos = None
            self.resizing = False
            self.resize_dir = None
            self.lastMousePos = None
            self.setCursor(Qt.CursorShape.SizeAllCursor)
                
        return super().eventFilter(obj, event)
    
    def is_on_title_bar(self, pos):
        title_geometry = self.title_bar.geometry()
        return title_geometry.contains(self.mapFromGlobal(pos))
    
    def get_resize_direction(self, pos):
        """检测鼠标在哪个边缘，返回调整大小的方向"""
        frame_geometry = self.frameGeometry()
        mouse_pos = self.mapFromGlobal(pos)
        
        edge_threshold = 5
        
        # 检查各个边缘
        if mouse_pos.y() <= edge_threshold:
            return 'top'
        elif mouse_pos.y() >= frame_geometry.height() - edge_threshold:
            return 'bottom'
        elif mouse_pos.x() <= edge_threshold:
            return 'left'
        elif mouse_pos.x() >= frame_geometry.width() - edge_threshold:
            return 'right'
        elif mouse_pos.x() <= edge_threshold and mouse_pos.y() <= edge_threshold:
            return 'top-left'
        elif mouse_pos.x() >= frame_geometry.width() - edge_threshold and mouse_pos.y() <= edge_threshold:
            return 'top-right'
        elif mouse_pos.x() <= edge_threshold and mouse_pos.y() >= frame_geometry.height() - edge_threshold:
            return 'bottom-left'
        elif mouse_pos.x() >= frame_geometry.width() - edge_threshold and mouse_pos.y() >= frame_geometry.height() - edge_threshold:
            return 'bottom-right'
            
        return None
    
    def resize_window(self, mouse_pos):
        """根据方向调整窗口大小"""
        if not self.resize_dir:
            return
            
        frame_geometry = self.frameGeometry()
        current_pos = self.pos()
        current_size = self.size()
        
        delta = mouse_pos - self.lastMousePos
        self.lastMousePos = mouse_pos
        
        new_geometry = QRect(frame_geometry)
        
        if 'top' in self.resize_dir:
            new_geometry.setTop(new_geometry.top() + delta.y())
            new_geometry.setHeight(current_size.height() - delta.y())
            
        if 'bottom' in self.resize_dir:
            new_geometry.setHeight(current_size.height() + delta.y())
            
        if 'left' in self.resize_dir:
            new_geometry.setLeft(new_geometry.left() + delta.x())
            new_geometry.setWidth(current_size.width() - delta.x())
            
        if 'right' in self.resize_dir:
            new_geometry.setWidth(current_size.width() + delta.x())
            
        # 确保最小大小
        min_size = self.minimumSize()
        if new_geometry.width() < min_size.width():
            if 'left' in self.resize_dir:
                new_geometry.setLeft(new_geometry.right() - min_size.width() + 1)
            new_geometry.setWidth(min_size.width())
            
        if new_geometry.height() < min_size.height():
            if 'top' in self.resize_dir:
                new_geometry.setTop(new_geometry.bottom() - min_size.height() + 1)
            new_geometry.setHeight(min_size.height())
            
        self.setGeometry(new_geometry)
    

        
    def clear_canvas(self):
        self.strokes.clear()
        self.current_path = None
        if hasattr(self, 'canvas_bg'):
            self.canvas_bg.clear()
        self.cleared.emit()
        
    def show_at(self, widget):
        """在指定控件附近显示窗口"""
        if widget:
            screen_geo = widget.screen().availableGeometry()
            widget_geo = widget.geometry()
            
            # 计算显示位置（在控件右下方）
            x = widget_geo.right() + 10
            y = widget_geo.top()
            
            # 确保不超出屏幕边界
            if x + self.width() > screen_geo.right():
                x = widget_geo.left() - self.width() - 10
            if y + self.height() > screen_geo.bottom():
                y = screen_geo.bottom() - self.height()
                
            self.move(x, y)
        
        self.show()
        self.activateWindow()
        self.raise_()

    def reset_up(self):
        self.up_timer.stop()
        self.up_running = False
        self.up_seconds = 0
        self.up_start_btn.setText("开始")
        self.up_label.setText(self.format_time(self.up_seconds))

    def update_up(self):
        self.up_seconds += 1
        self.up_label.setText(self.format_time(self.up_seconds))

    def toggle_down(self):
        if not self.down_running:
            # Check if we are resuming or starting new
            if self.down_remaining <= 0 and self.down_total_seconds == 0:
                # Start new
                minutes = self.down_min_spin.value()
                seconds = self.down_sec_spin.value()
                total = minutes * 60 + seconds
                if total <= 0:
                    return
                self.down_total_seconds = total
                self.down_remaining = total
            
            # Switch to label view
            self.input_widget.hide()
            self.down_label.show()
            self.down_label.setText(self.format_time(self.down_remaining))
            
            self.down_timer.start()
            self.down_running = True
            self.down_start_btn.setText("暂停")
        else:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("开始")

    def reset_down(self):
        self.down_timer.stop()
        self.down_running = False
        self.down_remaining = 0
        self.down_total_seconds = 0
        self.down_start_btn.setText("开始")
        
        # Show input again
        self.down_label.hide()
        self.input_widget.show()

    def update_down(self):
        if self.down_remaining > 0:
            self.down_remaining -= 1
            self.down_label.setText(self.format_time(self.down_remaining))
        
        if self.down_remaining <= 0 and self.down_running:
            self.down_timer.stop()
            self.down_running = False
            self.down_start_btn.setText("开始")
            
            self.stack.setCurrentWidget(self.completed_page)
            self.play_ring()
            self.shake_window()
