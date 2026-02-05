# 🎓 班级请假记录系统

<div align="center">

![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Version](https://img.shields.io/badge/Version-1.0.1(2)-orange.svg)

**一个让班主任不再因为统计请假而掉头发的神器**

[功能特点](#-核心功能) • [快速开始](#-快速开始) • [使用指南](#-使用指南) • [更新日志](#-更新日志)

</div>

---

## 📌 项目简介

这是一个用 Python + tkinter 开发的班级请假记录管理系统。

简单来说,就是帮你把那些乱七八糟的请假记录整理得井井有条,让你从"谁请假了?请几天?"的崩溃中解脱出来。

**适用人群**:
- 🏫 班主任
- 👨‍🏫 辅导员
- 😵‍💫 任何需要管理学生请假记录的苦命人

**为什么选择本系统?**
- ✅ 不用Excel表格手动统计,告别Ctrl+F的噩梦
- ✅ 不用担心数据丢失,自动保存,月初不清零
- ✅ 不用担心手滑误删,操作简单,容错率高
- ✅ 不用担心格式混乱,一键导出精美Excel

---

## 🎯 版本信息

| 项目 | 内容 |
|:----:|:----:|
| **当前版本** | v1.0.1(2) |
| **更新日期** | 2025-02-05 |
| **作者** | 112114141 |
| **开发语言** | ![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg) |
| **GUI框架** | tkinter |
| **数据格式** | JSON |
| **许可证** | MIT |

---

## ✨ 核心功能

### 👥 学生名单管理
> 再也不用为找不到学生名字而翻遍全班名单了

| 功能 | 说明 | 效率提升 |
|:----:|:----:|:--------:|
| ➕ 添加学生 | 一个一个加,慢工出细活 | ⭐⭐ |
| 📥 批量导入 | 复制粘贴,效率起飞 | ⭐⭐⭐⭐⭐ |
| ❌ 删除学生 | 支持多选,告别误删的噩梦 | ⭐⭐⭐ |
| 🔤 智能排序 | 按拼音自动排序,找学生不费眼 | ⭐⭐⭐⭐ |

### 📝 请假录入
> 告别弹窗地狱,点击即选,丝般顺滑

- 📅 **日期选择**: 左侧日历,想选哪天选哪天
- ⚡ **快速录入**: 点击"全天"或"半天"列,一键搞定
- 🛡️ **误操作保护**: 点击姓名列不会弹出对话框,避免手滑
- ⏰ **历史修改**: 想改哪天改哪天,时光倒流不是梦
- 🗑️ **清空保存**: 允许保存空记录,方便纠错

### ⚠️ 常请假名单
> 自动识别"请假达人",重点关注

- 📊 **自动统计**: 5天内请假≥3次的学生自动上榜
- 🔗 **双向联动**: 两个名单同步更新,一个都不能少

### 📅 日历功能
> 三种视图,满足你的所有需求

| 视图 | 功能 | 适用场景 |
|:----:|:----:|:--------:|
| 📅 日视图 | 单日查看和录入 | 日常记录 |
| 📆 周视图 | 一周的请假情况尽收眼底 | 周报统计 |
| 🗓️ 月视图 | 整月的请假记录一目了然 | 月度总结 |

- 🌟 **高亮显示**: 有记录的日期金光闪闪
- ⬅️➡️ **月份导航**: 上个月、下个月,想看哪月看哪月

### 📊 数据统计与导出
> 统计数据,一键导出,老板看了都说好

<details>
<summary><b>📈 四种统计类型</b></summary>

| 类型 | 说明 | 使用场景 |
|:----:|:----:|:--------:|
| 📅 当前日期 | 今天谁请假了? | 日常查看 |
| 📅 本周 | 这一周谁最"忙"? | 周报准备 |
| 📅 本月 | 一个月的请假大盘点 | 月度总结 |
| 📅 自定义 | 你说了算 | 特殊需求 |

</details>

<details>
<summary><b>🎨 颜色区分</b></summary>

| 类型 | 颜色 | 含义 |
|:----:|:----:|:----:|
| 🔵 工作日 | 蓝色 | 正常上学日 |
| 🟡 周六 | 黄色 | 还要上学 😭 |
| 🔴 周日 | 红色 | 终于休息了 😊 |

</details>

- 👥 **学生选择**: 全部学生 or 单个学生,灵活切换
- 🔄 **自动刷新**: 切换选项卡自动更新,省心省力
- 📥 **Excel导出**: 一键导出,格式精美

### 💾 数据持久化
> 数据安全,永不丢失

- 💾 **JSON存储**: 轻量级,易备份
- ♾️ **长期保存**: 月初不清零,数据永流传
- 🆕 **自动创建**: 首次运行自动生成数据文件

### ⚙️ 设置功能
> 个性化配置,随心所欲

- 🔄 **开机自启**: 可选择开机是否自动启动Web服务器
- ⏰ **备份频率**: 自定义自动备份频率(1-7天)
- 🗑️ **保留数量**: 保留备份文件数量(默认3个)
- 💾 **立即备份**: 手动创建备份,文件名格式"手动备份-[日期-时间]"
- 📥 **备份导入**: 一键恢复历史备份,数据不丢失
- 📁 **数据管理**: 数据文件统一放在data文件夹,方便管理

### 🎨 UI设计
> 告别Windows 98时代的审美

- 🎨 **现代配色**: 舒适护眼
- 📐 **宽松布局**: 不拥挤,不压抑
- ✨ **动画效果**: 启动淡入、保存成功提示,仪式感拉满
- 🖱️ **鼠标滚轮**: 所有可滚动组件都支持滚轮

---

## 📁 文件说明

```
class-leave-record-system/
├── 📄 班级请假记录系统.py    # 主程序文件
├── 📄 tkintercalendar.py      # 自定义日历组件
├── 📄 requirements.txt       # Python依赖包列表
├── 📄 README.md              # 本文档
├── 📄 .gitignore            # Git忽略配置
└── 📁 数据文件(运行时生成):
    └── data/                # 数据文件夹
        ├── 📄 students.json     # 学生名单数据
        └── 📄 leave_records.json # 请假记录数据
```

| 文件名 | 说明 | 是否必须 |
|:------:|:----:|:--------:|
| `班级请假记录系统.py` | 主程序文件 | ✅ 必须 |
| `tkintercalendar.py` | 日历组件 | ✅ 必须 |
| `requirements.txt` | Python依赖包列表 | ✅ 必须 |
| `students.json` | 学生名单数据(data文件夹) | ❌ 自动生成 |
| `leave_records.json` | 请假记录数据(data文件夹) | ❌ 自动生成 |
| `README.md` | 本文档 | ❌ 可选 |
| `.gitignore` | Git忽略配置 | ❌ 可选 |

---

## 🚀 快速开始

### 环境要求

- ![Windows](https://img.shields.io/badge/OS-Windows-blue.svg) Windows操作系统
- ![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg) Python 3.7 或更高版本

### 安装步骤

<details>
<summary><b>📥 方法一: 克隆项目</b></summary>

```bash
git clone https://github.com/yourusername/class-leave-record-system.git
cd class-leave-record-system
```

</details>

<details>
<summary><b>📥 方法二: 下载ZIP</b></summary>

1. 访问 [GitHub Releases](https://github.com/yourusername/class-leave-record-system/releases)
2. 下载最新版本的ZIP文件
3. 解压到任意目录

</details>

### 安装依赖

```bash
pip install -r requirements.txt
```

<details>
<summary><b>📋 依赖包列表</b></summary>

| 包名 | 版本 | 用途 |
|:----:|:----:|:----:|
| openpyxl | 3.1.2 | Excel文件处理 |
| pyinstaller | 6.18.0 | 打包成exe |
| pypinyin | 0.51.0 | 中文拼音排序 |

</details>

### 运行程序

```bash
python 班级请假记录系统.py
```

### 打包成exe(可选)

如果你想分享给没有安装Python的同事:

```bash
pyinstaller --onefile --noconsole --name "班级请假记录系统" "班级请假记录系统.py" "tkintercalendar.py"
```

打包完成后,exe文件在 `dist` 文件夹中。

---

## 🛠️ 技术细节

### 数据结构

**data/students.json**
```json
[
  "张三",
  "李四",
  "王五"
]
```

**data/leave_records.json**
```json
{
  "2026-02-04": {
    "张三": {"type": "full"},
    "李四": {"type": "half"}
  },
  "2026-02-05": {
    "王五": {"type": "full"}
  }
}
```

### 关键技术点

| 技术 | 说明 | 优势 |
|:----:|:----:|:----:|
| 🔒 线程安全 | 使用锁机制保护数据写入 | 防止并发问题 |
| 💾 原子性操作 | 使用临时文件+原子替换 | 保证数据完整性 |
| ⚡ 防抖优化 | 日历更新防抖 | 避免频繁刷新 |
| 🎨 Canvas绘制 | 使用Canvas绘制统计表格 | 支持动态行高 |
| ✨ 动画效果 | 启动淡入、保存成功提示 | 仪式感拉满 |

---

## 🌟 更新日志

<details>
<summary><b>v1.0.1(2) (2025-02-05)</b></summary>

### 🔧 优化
- 优化导入备份弹窗,去掉"取消"按钮,简化操作流程
- 用户可通过点击对话框关闭按钮(X)来关闭弹窗

</details>

<details>
<summary><b>v1.0.1 (2025-02-04)</b></summary>

### ✨ 新增
- 添加设置选项卡,支持个性化配置
- 添加手动备份功能,可随时创建备份
- 添加备份导入/删除功能,可恢复或删除历史备份
- 添加开机自启Web服务器选项
- 添加自动备份频率设置(1-7天)
- 添加保留备份文件数量设置(默认3个)
- 使用pypinyin实现完整拼音排序,支持所有汉字
- 数据文件统一放在data文件夹,方便管理
- 备份文件名格式优化为"手动备份-[日期-时间]"
- 添加自定义月历图标(带MTWTFSS星期字母)
- 优化UI布局,提升用户体验
- 添加动画效果(启动淡入、保存成功提示)
- 支持鼠标滚轮滚动
- 添加教程选项卡

### 🔧 优化
- 美化设置界面,添加标题栏和说明文字
- 修复emoji和文字对齐问题(统一使用Segoe UI Symbol字体)
- 优化数据文件管理,使用data文件夹集中存储
- 调整常请假名单规则为5天内≥3次(更严格)
- 优化统计数据刷新性能
- 优化日历高亮更新(防抖)
- 优化窗口大小改变时的表格刷新

### 🐛 修复
- 修复学生排序功能,使用pypinyin替代手动拼音映射
- 修复备份文件路径问题
- 修复备份文件名冒号问题(Windows文件系统不支持冒号)
- 修复备份功能数据检查问题
- 修复若干已知问题
- 修复Excel导出时的格式问题

</details>

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request!

<details>
<summary><b>📝 贡献指南</b></summary>

1. Fork 本仓库
2. 创建你的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交你的修改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启一个 Pull Request

</details>

---

## 📄 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](LICENSE) 文件

---

## 📮 联系方式

- 👤 **作者**: 112114141
- 📧 **邮箱**: your.email@example.com
- 🌐 **项目地址**: [GitHub](https://github.com/yourusername/class-leave-record-system)
- 🐛 **问题反馈**: [Issues](https://github.com/yourusername/class-leave-record-system/issues)

---

<div align="center">

**如果这个项目对你有帮助,请给个⭐Star支持一下!**

Made with ❤️ by 112114141

</div>

---

