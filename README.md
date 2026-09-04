# 食堂平台导入导出工具

一个基于 C# WinForms 的桌面小工具，用于在 **SQL Server（用友 U8 业务库）**、**MySQL（食堂平台库）** 和 **Excel 文件** 之间进行双向数据交换，服务于食堂采购 / 食安数据的采集与上报。

- 程序集名称：`execl导入导出程序`
- 解决方案：`out.sln`
- 项目文件：`out_execl.csproj`
- 目标框架：.NET Framework 4.7.2
- 主窗体标题：`食堂平台导入导出工具`

---

## 一、功能特性

| 功能 | 数据方向 | 说明 |
| --- | --- | --- |
| 导出成 Excel | SQL Server → Excel | 按时间范围从用友 U8 库查询采购、入库、出库、验收、食材等业务数据，写入一个多工作表的 `.xlsx` 文件 |
| Excel 导入 | Excel → MySQL | 读取 Excel 文件，逐工作表（Sheet）以 `INSERT ... ON DUPLICATE KEY UPDATE` 方式写入 MySQL |

### 导出功能（生成 10 个工作表）

| 工作表名 | 数据内容 |
| --- | --- |
| `gxcgzz_cghzdddxx` | 采购汇总单订单信息 |
| `gxcgzz_cghzdddxxxx` | 采购汇总单订单详情信息 |
| `gxcgzz_rkxx` | 入库信息 |
| `gxcgzz_rkxqxx` | 入库详情信息 |
| `gxcgzz_ckxx` | 出库信息 |
| `gxcgzz_ckxqxx` | 出库详情信息 |
| `gxcgzz_cgddysxx` | 采购订单验收信息 |
| `gxcgzz_cgddysxqxx` | 采购订单验收详细信息 |
| `gxcgzz_scxx` | 食材信息 |
| `gxcgzz_scjgxx` | 食材价格信息 |

> 数据均来自用友 U8 的表（如 `PU_PurchaseOrder`、`PU_PurchaseArrival`、`ST_RDRecord`、`AA_Inventory` 等），并过滤掉分类代码以 `14/99/98/00` 开头的记录。

### 导入功能

- 支持 `.xlsx`、`.xls`、`.txt` 文件（由 Spire.XLS 解析）。
- 遍历 Excel 中的**每一个工作表**，工作表名即目标 MySQL 表名。
- 自动根据列类型构建参数化 `INSERT` 语句，遇到主键/唯一键冲突时执行 `ON DUPLICATE KEY UPDATE` 更新。

---

## 二、环境要求

| 项 | 要求 |
| --- | --- |
| 操作系统 | Windows（WinForms 桌面应用） |
| 运行时 | .NET Framework 4.7.2 |
| 开发环境 | Visual Studio 2017 及以上（项目为旧式 `csproj` 格式） |
| 数据库 | SQL Server（用友 U8 数据源）、MySQL 5.x / 8.x |
| 依赖库 | NuGet 包（详见下） |

### 主要 NuGet 依赖

| 包 | 用途 |
| --- | --- |
| `NPOI` (2.7.3) | 生成 `.xlsx` Excel 文件 |
| `Spire.XLS` (14.1.0) | 读取 / 解析 Excel 文件 |
| `MySql.Data` (9.3.0) | 连接 MySQL 数据库 |
| `Newtonsoft.Json` (13.0.3) | JSON 处理 |
| `DevExpress.*` (23.2) | 已引用（当前代码未直接使用） |
| 其它（BouncyCastle、Google.Protobuf、K4os、MathNet、SixLabors 等） | 多为 Spire.XLS 等包的传递依赖 |

---

## 三、架构说明

项目采用简单的三层结构：

```
表示层（UI）        Form1.cs / Form1.Designer.cs
                      └─ 用户交互：选择时间、选择文件路径、触发导入/导出

业务逻辑层          Form1.cs 中的事件处理方法
                      ├─ button1_Click：SQL Server 查询 → 组装工作簿 → 导出 Excel
                      └─ button7_Click：读取 Excel → 构建 SQL → 写入 MySQL

数据访问层          DBHelper.cs      （SQL Server 帮助类，namespace: out_exel）
                    DbHelperMySQL.cs （MySQL 帮助类，namespace: out_execl）

文件处理层          NPOI（写）、Spire.XLS（读）
```

### 数据流

**导出流程：**

```
SQL Server（用友 U8）
   │  按时间范围执行多条 SELECT 查询
   ▼
DataTable（DBHelper.ExecuterQuery）
   │  逐行写入
   ▼
NPOI XSSFWorkbook（多工作表）
   │  workbook.Write
   ▼
本地 .xlsx 文件
```

**导入流程：**

```
本地 Excel 文件（.xlsx/.xls/.txt）
   │  Spire.XLS workbook.LoadFromFile
   ▼
逐个 Worksheet → DataTable
   │  构建 INSERT ... ON DUPLICATE KEY UPDATE
   ▼
MySQL（yongyou_zjedu_st_2025）
```

### 主要类说明

| 文件 | 类 / 命名空间 | 职责 |
| --- | --- | --- |
| `Program.cs` | `Program` | 程序入口，启动主窗体 |
| `Form1.cs` | `Form1` | 主窗体逻辑，含导入、导出两个核心方法及数据库连接字符串 |
| `Form1.Designer.cs` | `Form1` | 窗体控件布局（设计器生成） |
| `DBHelper.cs` | `DBHelper` (out_exel) | SQL Server 数据访问：查询、增删改、事务、DataTable 转 JSON |
| `DbHelperMySQL.cs` | `DbHelperMySQL` (out_execl) | MySQL 数据访问：查询、执行、事务、取最大 ID 等 |

---

## 四、使用说明

### 4.1 运行

1. 用 Visual Studio 打开 `out.sln`，还原 NuGet 包后编译运行；
2. 或直接运行 `bin\Debug\execl导入导出程序.exe`（需保证依赖 DLL 齐全）。
3. 启动前请确保能访问下方「数据库配置」中对应的 SQL Server 与 MySQL。

### 4.2 导出数据到 Excel

1. 在左侧「导出」区域，通过 `开始时间`、`结束时间` 选择时间范围；
2. 点击 **「导出成 excel」** 按钮；
3. 在弹出的保存对话框中指定输出文件名与保存位置（支持 `.xlsx`、`.xls`、`.txt` 等）；
4. 程序自动执行 10 组查询并写入对应工作表，完成后弹窗提示生成的文件路径及无数据的表格数量。

### 4.3 导入 Excel 到 MySQL

1. 在右侧「导入」区域，点击 **「excel 导入」** 按钮；
2. 选择要导入的 Excel 文件（`.xlsx` / `.xls` / `.txt`）；
3. 程序读取文件后，将**每个工作表按工作表名**写入 MySQL 对应表，完成后弹窗提示「导入成功」。

> 注意：导入前请确保 MySQL 中已存在与工作表同名的表结构，且字段与 Excel 列对应，否则会报「导入失败」。

---

## 五、数据库配置

数据库连接字符串以**常量形式硬编码**在 `Form1.cs` 文件顶部：

| 变量 | 数据库 | 用途 |
| --- | --- | --- |
| `conn_mes` | SQL Server `10.10.210.130` / 库 `UFTData545228_900000` | 导出数据源（用友 U8） |
| `constring` | MySQL `172.17.41.196` / 库 `yongyou_zjedu_st_2025` | 导入目标库 |
| `constring2` | SQL Server `127.0.0.1` / 库 `text` | 备用连接（当前未使用） |

需要切换环境时，直接修改 `Form1.cs` 中对应常量即可。

---

## 六、注意事项

- **安全提示**：连接字符串中包含明文数据库账号密码，请勿将配置文件对外公开；建议改为读取 `App.config` 的 `AppSettings` 或使用加密配置。
- **导入的表名即工作表名**：若 MySQL 中没有同名表，导入会失败，需预先建表。
- **类型映射有限**：`GetMySqlDbType` 仅覆盖 `int / string / DateTime / decimal / bool / long / float / double / byte[]`，其余类型默认按 `VarChar` 处理，复杂类型可能需要补充。
- **固定等待**：`button7_Click` 中存在 `Thread.Sleep(5000)` 固定延时，属遗留逻辑。
- **导出 SQL 使用字符串拼接**：时间参数通过 `string.Format` 拼接进 SQL，存在潜在 SQL 注入风险（本工具为内部固定查询，风险可控），建议后续改为参数化查询。
