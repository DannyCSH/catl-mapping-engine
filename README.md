# CATL 四层映射引擎 · Product Site v2

## 产品定位

这是 HKU × CATL 联合研究项目的**产品原型系统**，用于将宁德时代的电池数据，通过四层映射逻辑，转换为符合欧盟《电池法》(EU Battery Regulation 2023/1542) 和 CBAM 口径要求的结构化输出包。

**这不是：**
- 法定提交系统
- 正式 TÜV / NB / CAB 审计报告
- 已接入真实 Tractus-X / Catena-X 连接器的平台
- 已接入企业真实数据库的系统

**这是：**
- 输入数据组织工具
- Excel 四层映射重算触发器
- EU / CBAM / TÜV Audit-Prep 输出包生成器
- 研究与竞赛演示原型

---

## 技术栈

- **后端：** Python 3 + Flask
- **前端：** Vanilla HTML/CSS/JS（无框架依赖）
- **计算引擎：** Excel（通过 win32com COM 接口重算）
- **数据格式：** CSV（值版）、Excel Workbook（模板）

---

## 本地运行

### 前置条件

- Windows（因为需要 Excel COM）
- Python 3.9+
- Microsoft Excel（已安装）
- pip

### 步骤

```powershell
# 1. 进入项目目录
cd D:\VibeCoding\codex\reports\hku_catl_standards_mapping\product_site_v2

# 2. 安装依赖
pip install -r requirements.txt

# 3. 启动服务
python app.py
```

然后浏览器打开：**http://127.0.0.1:5000**

---

## 页面说明

| 路由 | 页面 | 说明 |
|------|------|------|
| `/` | 首页 | 产品官网风格，展示系统价值、边界声明、四层映射简介 |
| `/intake` | 数据填写工作台 | 按分类填写 CATL 数据，支持草稿保存 |
| `/outputs` | 输出中心总览 | 三个输出包入口卡片 + 快速预览标签页 |
| `/outputs/eu/<session_id>` | EU Battery Regulation Pack | EU 输出详细视图，含证据/核验/权限/缺口子页签 |
| `/outputs/cbam/<session_id>` | CBAM Integrated Pack | CBAM 影子输出详细视图 |
| `/outputs/tuv/<session_id>` | TÜV Audit-Prep Pack | 审计准备视图，含 manifest 生成 |

---

## API 路由

| 方法 | 路由 | 说明 |
|------|------|------|
| GET | `/api/schema/intake` | 获取表单 schema（来自 `01_手动输入.csv`）|
| POST | `/api/session` | 创建新 session（复制 master workbook 到 runtime）|
| GET | `/api/session/<id>` | 获取 session 状态 |
| POST | `/api/session/<id>/save` | 保存表单草稿 |
| POST | `/api/session/<id>/calculate` | 写 Excel → COM 重算 → 导出 CSV |
| GET | `/api/session/<id>/outputs/<target>` | 获取输出数据（eu / cbam / tuv / evidence 等）|
| GET | `/api/session/<id>/output_summary` | 输出统计摘要 |
| GET | `/api/session/<id>/export/<target>` | 导出 ZIP 包（eu / cbam / tuv）|
| GET | `/api/health` | 健康检查 |

---

## 工作流程

1. **选择模板模式**（Generic Template 或 Shenxing Case Demo）
   - 系统在 `runtime/<session_id>/` 下复制 master workbook
   - 表单预填通用模板默认值

2. **填写 / 修改数据**
   - 86+ 个字段按 8 个分类分组展示
   - 所有修改自动保存为 `runtime/<session_id>/draft.json`

3. **触发计算**
   - 后端将输入值写入 `01_手动输入` sheet 的 G 列
   - 通过 Excel COM 执行 `CalculateFullRebuild()`
   - 将 11 张关键 sheet 导出为 CSV 到 session 目录

4. **查看输出**
   - EU Pack、CBAM Pack、TÜV Audit-Prep Pack
   - 每包含：输出表 + 证据映射 + 核验映射 + 权限映射 + 缺口清单

5. **导出整包**
   - ZIP 格式，包含多个 CSV + manifest.json + README.txt
   - 每个 session 的导出包保存在 `exports/` 目录

---

## 目录结构

```
product_site_v2/
├── app.py                    # Flask 主程序（含所有 API）
├── requirements.txt
├── README.md
│
├── templates/                # HTML 页面
│   ├── index.html            # 首页
│   ├── intake.html           # 数据填写工作台
│   ├── outputs.html          # 输出中心总览
│   ├── output_eu.html        # EU 输出详情页
│   ├── output_cbam.html      # CBAM 输出详情页
│   └── output_tuv.html       # TÜV Audit-Prep 详情页
│
├── static/
│   ├── css/style.css         # 样式（UTF-8）
│   └── js/main.js            # 前端逻辑（UTF-8）
│
├── runtime/                  # 运行时 session 数据（每次计算新建子目录）
│   └── <session_id>/
│       ├── workbook.xlsx     # session 副本
│       ├── draft.json        # 保存的输入值
│       ├── 02_欧盟电池法输出.csv
│       ├── 03_CBAM输出.csv
│       └── ...（共 11 张导出 CSV）
│
└── exports/                  # 导出包 ZIP
    └── CATL_EU_pack_<id>_<ts>.zip
```

---

## 基线数据文件

| 文件 | 用途 |
|------|------|
| `catl_four_layer_mapping_engine_generic_template_2026-04-07.xlsx` | 计算引擎 master template |
| `catl_four_layer_mapping_engine_shenxing_case_snapshot_2026-04-07.xlsx` | 案例演示模式 |
| `review_csv_generic_template_v1_values/` | 表单 schema 和值版 CSV |

路径：`D:\VibeCoding\codex\reports\hku_catl_standards_mapping\`

---

## 已知限制

1. **Excel COM 依赖：** `calculate` 接口需要 Windows + 已安装 Excel。
   - 如果 COM 不可用，后端返回 `success_com_fallback`，使用静态 CSV 快照作为输出。
   - 请查看 `/api/health` 的 `com_available` 字段确认。

2. **中文编码：** 所有文件使用 UTF-8 编码。读取 Excel 导出的 CSV 时，
   后端会依次尝试 `utf-8-sig` → `gbk` → `latin1`。

3. **Session 管理：** session ID 通过 URL 传递，无超时清理机制。
   - Session 数据保存在 `runtime/` 目录，重启 Flask 后清除。

4. **非生产系统：** 本系统为研究原型，不具备生产级安全性、并发控制或备份机制。

---

## 数据保存路径约定

所有运行时数据（session、exports、logs）均保存在 D 盘项目目录下：

- `runtime/` — 每次 session 的 workbook 副本和导出 CSV
- `exports/` — 每次导出的 ZIP 包
- `logs/` — 本地运行日志

**不写入：**
- `C:\Users\huawei\.claude\` 及任何 C 盘用户目录
- 系统 temp 目录作为长期存储

---

## 🚀 Railway 云端部署（无需 Excel COM，纯 Python 计算）

### 部署步骤

1. **推送到 GitHub**
   ```bash
   cd deploy/catl_mapping_engine
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/YOUR_USERNAME/catl-mapping-engine.git
   git push -u origin main
   ```

2. **接入 Railway**
   - 访问 [railway.app](https://railway.app)，用 GitHub 账号登录
   - 点击 **New Project** → **Deploy from GitHub Repo**
   - 选择刚才创建的仓库
   - Railway 会自动读取 `railway.json` 并部署

3. **获取公开链接**
   - 部署完成后，Railway 会分配一个公开 URL（如 `https://catl-mapping.up.railway.app`）
   - 把这个链接发给老师/同学/CATL 员工即可公开访问

### 重要说明

- 云端版本**不依赖 Excel COM**，使用纯 Python 计算引擎
- Session 数据保存在 Railway 分配的临时磁盘上（重启会丢失）
- 如果需要持久化，需要额外配置 PostgreSQL 或对象存储
