"""
自定义日历组件 - 重写版本
"""

import tkinter as tk
from tkinter import ttk
import datetime


class Calendar:
    """日历组件"""
    
    def __init__(self, parent, on_date_select=None, colors=None, on_week_select=None, on_month_select=None):
        self.parent = parent
        self.on_date_select = on_date_select
        self.on_week_select = on_week_select
        self.on_month_select = on_month_select
        self.selected_date = None
        self.selected_week = None
        self.selected_month = None
        self.highlighted_dates = set()
        self.current_year = datetime.datetime.now().year
        self.current_month = datetime.datetime.now().month
        
        # 配色方案
        if colors:
            self.colors = colors
        else:
            self.colors = {
                'bg': '#F5F7FA',
                'white': '#FFFFFF',
                'accent': '#3498DB',
                'success': '#27AE60',
                'warning': '#F39C12',
                'danger': '#E74C3C',
                'light_gray': '#ECF0F1',
                'fg': '#2C3E50',
                'week_bg': '#E8F6F3',
                'month_bg': '#FFF3E0'
            }
        
        # 动画相关
        self.animating = False
        self.day_buttons = {}
        
        self.create_widgets()
    
    def create_widgets(self):
        """创建日历组件"""
        # 月份导航
        nav_frame = tk.Frame(self.parent, bg=self.colors['bg'])
        nav_frame.pack(fill=tk.X, padx=2, pady=2)
        
        self.prev_btn = tk.Button(nav_frame, text="<", 
                                  command=self.prev_month,
                                  bg=self.colors['light_gray'], fg=self.colors['fg'],
                                  font=('Microsoft YaHei', 8, 'bold'), relief='flat',
                                  padx=4, pady=2, cursor='hand2', width=2)
        self.prev_btn.pack(side=tk.LEFT)
        
        self.month_label = tk.Label(nav_frame, text="", 
                                   font=('Microsoft YaHei', 9, 'bold'),
                                   bg=self.colors['bg'], fg=self.colors['fg'])
        self.month_label.pack(side=tk.LEFT, expand=True, padx=2)
        
        self.next_btn = tk.Button(nav_frame, text=">", 
                                  command=self.next_month,
                                  bg=self.colors['light_gray'], fg=self.colors['fg'],
                                  font=('Microsoft YaHei', 8, 'bold'), relief='flat',
                                  padx=4, pady=2, cursor='hand2', width=2)
        self.next_btn.pack(side=tk.RIGHT)
        
        # 添加按钮悬停效果
        self._add_button_hover_effect(self.prev_btn, self.colors['light_gray'], '#D5DBDB')
        self._add_button_hover_effect(self.next_btn, self.colors['light_gray'], '#D5DBDB')
        
        # 星期标题
        week_frame = tk.Frame(self.parent, bg=self.colors['bg'])
        week_frame.pack(fill=tk.X, padx=2, pady=(0, 1))
        
        weekdays = ["日", "一", "二", "三", "四", "五", "六"]
        for i, day in enumerate(weekdays):
            label = tk.Label(week_frame, text=day, 
                           font=('Microsoft YaHei', 6, 'bold'),
                           bg=self.colors['bg'], fg=self.colors['fg'])
            label.grid(row=0, column=i, padx=0, pady=0, sticky='nsew')
        
        for i in range(7):
            week_frame.grid_columnconfigure(i, weight=1)
        
        # 日历主体
        self.calendar_frame = tk.Frame(self.parent, bg=self.colors['bg'])
        self.calendar_frame.pack(fill=tk.BOTH, expand=True, padx=2, pady=(0, 2))
        
        self.day_buttons = {}
        self.update_calendar()
    
    def update_calendar(self):
        """更新日历显示"""
        # 清空现有按钮
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()
        
        self.day_buttons = {}
        
        # 更新月份标签
        month_names = ["1月", "2月", "3月", "4月", "5月", "6月",
                      "7月", "8月", "9月", "10月", "11月", "12月"]
        self.month_label.config(text=f"{self.current_year}年 {month_names[self.current_month-1]}")
        
        # 获取该月第一天是星期几
        first_day = datetime.datetime(self.current_year, self.current_month, 1)
        start_weekday = first_day.weekday() + 1  # 0=周一, 6=周日, 转换为0=周日, 6=周六
        
        # 获取该月总天数
        if self.current_month == 12:
            next_month = datetime.datetime(self.current_year + 1, 1, 1)
        else:
            next_month = datetime.datetime(self.current_year, self.current_month + 1, 1)
        total_days = (next_month - first_day).days
        
        # 创建日历按钮
        day = 1
        for row in range(6):
            for col in range(7):
                if row == 0 and col < start_weekday:
                    # 空白占位
                    label = tk.Label(self.calendar_frame, text="",
                                   bg=self.colors['bg'])
                    label.grid(row=row, column=col, padx=0, pady=0, sticky='nsew')
                    continue
                if day > total_days:
                    break
                
                date_str = f"{self.current_year}-{self.current_month:02d}-{day:02d}"
                
                # 检查是否是高亮日期
                bg_color = self.colors['white']
                fg_color = self.colors['fg']
                font_weight = 'normal'
                
                if date_str in self.highlighted_dates:
                    bg_color = self.colors['warning']  # 金色
                    fg_color = self.colors['white']
                    font_weight = 'bold'
                
                # 检查是否是选中日期
                if date_str == self.selected_date:
                    bg_color = self.colors['success']  # 绿色
                    fg_color = self.colors['white']
                    font_weight = 'bold'
                
                # 检查是否在选中的周中
                if self.selected_week:
                    week_start, week_end = self.selected_week
                    if week_start <= date_str <= week_end:
                        if date_str != self.selected_date:  # 不是选中的日期
                            bg_color = self.colors['week_bg']
                            fg_color = self.colors['fg']
                
                # 检查是否在选中的月中
                if self.selected_month:
                    if date_str != self.selected_date:  # 不是选中的日期
                        bg_color = self.colors['month_bg']
                        fg_color = self.colors['fg']
                
                btn = tk.Button(self.calendar_frame, text=str(day), 
                               bg=bg_color, fg=fg_color,
                               font=('Microsoft YaHei', 7, font_weight),
                               relief='flat', cursor='hand2',
                               padx=0, pady=0,
                               command=lambda d=date_str: self.select_date(d))
                btn.grid(row=row, column=col, padx=0, pady=0, sticky='nsew')
                self.calendar_frame.grid_columnconfigure(col, weight=1, minsize=35)
                self.calendar_frame.grid_rowconfigure(row, weight=1, minsize=25)
                
                # 存储按钮引用
                self.day_buttons[date_str] = btn
                
                # 添加按钮悬停效果
                self._add_button_hover_effect(btn, bg_color, self._darken_color(bg_color))
                
                day += 1
        
        # 配置行权重
        for row in range(6):
            self.calendar_frame.grid_rowconfigure(row, weight=1)
    
    def _darken_color(self, color, factor=0.8):
        """使颜色变暗"""
        if color.startswith('#'):
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            r = int(r * factor)
            g = int(g * factor)
            b = int(b * factor)
            return f'#{r:02x}{g:02x}{b:02x}'
        return color
    
    def _add_button_hover_effect(self, button, normal_color, hover_color):
        """为按钮添加悬停效果"""
        def on_enter(event):
            button.configure(bg=hover_color)
        
        def on_leave(event):
            button.configure(bg=normal_color)
        
        button.bind('<Enter>', on_enter)
        button.bind('<Leave>', on_leave)
    
    def prev_month(self):
        """上个月"""
        if self.animating:
            return
        
        self.current_month -= 1
        if self.current_month < 1:
            self.current_month = 12
            self.current_year -= 1
        self.update_calendar()
    
    def next_month(self):
        """下个月"""
        if self.animating:
            return
        
        self.current_month += 1
        if self.current_month > 12:
            self.current_month = 1
            self.current_year += 1
        self.update_calendar()
    
    def select_date(self, date_str: str):
        """选择日期"""
        if self.animating:
            return
        
        # 清除周和月选择
        self.selected_week = None
        self.selected_month = None
        
        old_selected = self.selected_date
        self.selected_date = date_str
        
        # 更新显示
        self.update_calendar()
        
        if self.on_date_select:
            self.on_date_select(date_str)
    
    def select_week(self, week_start: str, week_end: str):
        """选择一周"""
        if self.animating:
            return
        
        # 清除日期和月选择
        self.selected_date = None
        self.selected_month = None
        
        self.selected_week = (week_start, week_end)
        
        # 更新显示
        self.update_calendar()
        
        if self.on_week_select:
            self.on_week_select(week_start, week_end)
    
    def select_month(self, year: int, month: int):
        """选择一个月"""
        if self.animating:
            return
        
        # 清除日期和周选择
        self.selected_date = None
        self.selected_week = None
        
        self.selected_month = (year, month)
        
        # 更新显示
        self.update_calendar()
        
        if self.on_month_select:
            self.on_month_select(year, month)
    
    def highlight_dates(self, dates: list):
        """高亮显示日期"""
        self.highlighted_dates = set(dates)
        self.update_calendar()
    
    def set_selected_date(self, date_str: str):
        """设置选中日期"""
        self.selected_date = date_str
        self.selected_week = None
        self.selected_month = None
        try:
            date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
            self.current_year = date.year
            self.current_month = date.month
        except:
            pass
        self.update_calendar()
    
    def get_selected_date(self) -> str:
        """获取选中日期"""
        return self.selected_date
    
    def get_selected_week(self) -> tuple:
        """获取选中的周"""
        return self.selected_week
    
    def get_selected_month(self) -> tuple:
        """获取选中的月"""
        return self.selected_month


# 兼容性别名
CalendarWidget = Calendar