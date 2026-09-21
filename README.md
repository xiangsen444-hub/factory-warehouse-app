# 工厂仓库管理系统

使用 Python、Streamlit 和 SQLite 编写的轻量仓库管理系统，用于登记出入库、查询库存、管理库位和打印标签。此仓库提供可本地运行的源码和匿名演示数据，不包含实际仓库数据库、客户照片、图纸或备份。

## 功能

- **库存查询与库位视图**：搜索物料、客户和位置，查看库存与低库存提示，以及物料照片和 PDF 图纸；按库区查看格子占用情况。
- **出入库与扫码**：单笔入库、领用或报废出库，扫码上架，批量扫码出入库；出库时检查可用库存。
- **出库提报与处理**：无需编辑密码即可提交领料申请，由已解锁的负责人确认后扣减库存；支持提交问题反馈。
- **物料与辅料管理**：维护物料、料号别名、客户和库位，导入 Excel / CSV，处理盘点差异；单独管理辅料、配件及其出入库记录。
- **流水与导出**：查询出入库明细，导出 CSV / Excel，支持带图片的表格导出。
- **二维码与标签**：生成物料二维码、位置标签，批量预览和打印；物料二维码可打开对应详情页。
- **离线问答**：根据库存数据和预设规则回答查询问题，不调用外部 AI 服务。
- **备份与访问统计**：每日自动备份数据库、滚动备份附件，也可手动备份；访问统计需要解锁后查看。

## 文件说明

| 文件 | 用途 |
| --- | --- |
| `app.py` | 仓库系统主程序 |
| `requirements.txt` | 仓库系统依赖 |
| `练习场启动.bat` | Windows 仓库系统启动脚本 |
| `label_tool.py` | 独立标签打印工具，不读写仓库数据库 |
| `label_tool_requirements.txt` | 独立标签工具依赖 |
| `start_label_tool.bat` | Windows 独立标签工具启动脚本 |
| `stock_api.py` | 可选的本机库存查询、工单扣料接口 |

## 本地启动

需要 Python 3.11 或更新版本。先在项目目录创建虚拟环境并安装依赖；依赖版本已在 `requirements.txt` 中固定。

### Windows PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
$env:WAREHOUSE_EDIT_PASSWORD = "替换为你自己的编辑密码"
.\.venv\Scripts\python.exe -m streamlit run app.py --server.port 8502
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
export WAREHOUSE_EDIT_PASSWORD='替换为你自己的编辑密码'
python -m streamlit run app.py --server.port 8502
```

浏览器打开 `http://localhost:8502`。程序首次启动会自动创建 `warehouse.db`，在空库中写入匿名演示物料和出入库流水，无需下载数据库。

### 编辑密码

设置环境变量 `WAREHOUSE_EDIT_PASSWORD` 后，可以在页面输入该密码解锁入库、出库、物料管理、提报处理等功能。也可以在本机 `.streamlit/secrets.toml` 中配置同名设置：

```toml
WAREHOUSE_EDIT_PASSWORD = "替换为你自己的编辑密码"
```

未配置密码时，编辑功能保持锁定。库存查询、流水查询和出库提报等页面仍可使用。密码或 secrets 文件不应提交到 Git。

### 局域网扫码访问

需要让同一局域网中的手机访问时，启动参数增加 `--server.address 0.0.0.0`，然后在“批量打印二维码”页设置手机能访问的系统地址，例如 `http://你的电脑局域网IP:8502`。物料二维码使用这个地址；地址变化后，需要重新生成并打印标签。

## 独立标签打印工具

该工具可以导入 Excel / CSV 或手工输入标签内容，生成文字和可选二维码，不依赖仓库系统，也不会修改仓库库存。在已启用的虚拟环境中运行下面的命令；Windows 若未激活虚拟环境，将本节及 API 启动命令中的 `python` 替换为 `.\.venv\Scripts\python.exe`：

```bash
python -m pip install -r label_tool_requirements.txt
python -m streamlit run label_tool.py --server.port 8503
```

Windows 也可运行 `start_label_tool.bat`。直接向系统打印机发送任务需要 Windows、`pywin32` 和可用的打印机驱动；其他系统可使用页面提供的标签预览、下载功能。中文标签需要安装可用的中文字体。

## 可选库存 API

先至少运行一次主程序，完成数据库建表，再在项目目录运行：

```bash
python stock_api.py
```

接口默认只监听 `127.0.0.1:8510`，与主程序共用同目录的数据库；可通过环境变量 `STOCK_API_PORT` 修改端口。它用于本机程序对接，没有鉴权功能，不作为公网服务入口。首次启动还会添加一条用于生产扣料演示的原料及期初库存。

| 请求 | 用途 |
| --- | --- |
| `GET /health` | 健康检查 |
| `GET /stock/lookup?q=查询内容` | 按物料编号、料号、二维码网址等查询物料和库存 |
| `GET /stock/by-task?id=工单编号` | 查询某个工单的出库记录 |
| `POST /stock/issue` | 传入 `material_id`、`qty` 及可选的 `task_id`、`stage_id`、`operator`、`note`，登记工单扣料 |

库存接口也支持对应的 `/api/stock/...` 路径。主程序和 API 应在同一项目副本中运行，避免读取不同的数据库。

## 本地数据

运行后生成的数据保存在项目目录：

- `warehouse.db`：物料、库存流水、设置和访问记录。
- `uploads/`：产品照片、图纸等附件。
- `backups/`：数据库备份和附件的滚动备份。
- `label_tool_settings.json`：独立标签工具的本机设置。

这些文件和目录已列入 `.gitignore`。迁移实际使用环境时，应单独备份并转移数据库与附件；从 GitHub 下载源码不会包含实际库存。
