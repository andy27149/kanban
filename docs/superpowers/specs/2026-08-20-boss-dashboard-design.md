# 老板看板（Boss Dashboard）设计文档

日期：2026-08-20

## 背景与目标

数据维护人手头有若干张 Excel 表格，需要从中提取指定的行列数据，生成一个包含 KPI 卡片、趋势图、占比/排名图表、明细表格的综合看板页面，供老板查看公司经营情况。

约束条件：
- 数据维护人不具备安装 Python / 配置开发环境的能力，因此不能依赖"本地运行脚本"的方案。
- 未来部署环境尚未最终确定（可能继续本机使用，也可能挂到公司服务器），架构需要让这个决定可以推迟，且现在做的东西以后不用推倒重来。
- 数据每次更新是"手动定期更新"（维护人隔一段时间替换/更新一次 Excel），暂不需要自动定时拉取。
- v1 只需要展示最新一次快照，不做历史趋势，但架构上不能堵死未来加历史趋势的路。
- Excel 的工作表结构（sheet 名、列的位置）是否长期稳定尚不确定，取数逻辑需要有一定容错能力。

## 整体架构

用一个轻量级 Web 应用（Python/Flask）同时承担"接收 Excel 上传"和"展示看板"两个职责，本地可直接运行，未来迁移到服务器时代码不需要大改。

```
[维护人的电脑]                    [Web 服务器（Flask）]                     [老板的浏览器]
     |                                      |                                    |
  打开 /upload（需密码）  ---------->   校验密码                                   |
  选择本地 Excel 并上传                     |                                    |
                                    保存上传文件到 uploads/                        |
                                    按 config.yaml 解析 Excel                     |
                                    (pandas / openpyxl)                          |
                                            |                                    |
                                    生成 data.json（覆盖式，只存最新快照）           |
                                            |                                    |
                                            |  <----- 直接打开 /dashboard（只读，无需密码）
                                    返回看板页面 HTML                              |
                                    页面 JS 请求 /api/data 读取 data.json 渲染    |
                                    KPI 卡片 / 趋势图 / 占比图 / 明细表格           |
```

### 组件划分

1. **Flask Web 应用**（`app.py` 等）
   - `GET/POST /upload`：需要密码保护。GET 返回上传表单页；POST 接收 Excel 文件，保存后触发解析流程，返回处理结果（成功/失败提示）。
   - `GET /dashboard`：只读页面，无需密码，任何拿到链接的人都可访问。渲染看板骨架 HTML，实际数据由前端 JS 异步请求获取。
   - `GET /api/data`：返回 `data.json` 的内容（JSON），供看板页面渲染使用。
   - 密码校验：简单的共享密码（存于服务器环境变量/配置文件），登录后设置 session，仅保护 `/upload` 相关路由；`/dashboard` 和 `/api/data` 不受此限制。

   **上传处理细节：**
   - 表单支持一次选择单个或多个 Excel 文件（`<input type="file" multiple>`）。
   - "上传完成"的判定不需要额外机制：HTTP 请求本身就是完成信号——Werkzeug/Flask 会先把整个 `multipart/form-data` 请求体接收完整，才会调用 `/upload` 的视图函数，因此只要该函数开始执行，就已保证本次提交的所有文件都完整到达服务器。前端只需在请求发出后显示"上传中"，收到响应后展示成功/失败结果即可，无需轮询或监听文件夹。
   - 存储规则：每个上传的文件按其原始文件名保存到固定目录 `uploads/`（文件名需与 `config.yaml` 中的 `source_file` 字段对应）。若 `uploads/` 中已存在同名文件，直接覆盖；不存在则新建。
   - 触发解析：本次请求中的所有文件保存完毕后，解析器读取**整个 `uploads/` 目录**（而非仅本次上传的文件）并结合 `config.yaml` 重新生成 `data.json`——这样即使一次只更新了其中一张表，其余表仍使用之前保存的版本，`data.json` 始终是完整快照。

2. **配置映射文件 `config.yaml`**
   - 描述看板上每一个指标/图表的数据来源：对应哪个 Excel 文件、哪个 sheet、以及取数方式。
   - KPI 和表格支持以下取数方式，应对"表结构是否稳定不确定"以及"部分指标需要跨指标运算/聚合"的问题：
     - **固定范围**（`fixed_range`，仅 KPI）：直接指定单元格范围，如 `Sheet1!B2:B10`。
     - **按表头匹配**（`header_match`，KPI/图表/表格通用）：指定表头文字（如"销售额"），程序自动定位该列，不依赖固定列号，能容忍列顺序变化。
     - **计算型**（`computed`，仅 KPI）：不直接读 Excel，而是对已经取到的其他 KPI 值做四则运算，用于"差值/占比"这类衍生指标（如"库存账实差异 = 库存余额 − 港口实际库存"）。通过 `from`/`minus` 等字段引用其他 KPI 的 `key`；若被引用的 KPI 取数失败，计算型 KPI 也标记为取数失败并给出说明。
     - **分组汇总**（`group_by_sum`，仅表格）：按指定表头分组（如"收货人"），对另一个数值表头求和（如"目的港称重"），并按汇总值降序排列，用于生成排名/占比类视图。
   - 新增/修改指标只需要改这个配置文件，不需要碰解析代码。
   - **表格多视图切换**：多个 `tables` 配置项共享同一个 `view_group` 时，前端会将它们渲染成同一张卡片，通过 tab 在多个视图之间切换（例如"明细列表" / "按收货人排名"），默认展示 `view_group` 中第一个列出的视图。每个视图项用 `view_label` 作为 tab 上显示的文字。

   示例结构：
   ```yaml
   kpis:
     - key: total_revenue
       label: "总营收"
       source_file: "经营数据.xlsx"
       sheet: "汇总"
       mode: fixed_range      # fixed_range | header_match | computed
       range: "B2"
     - key: total_profit
       label: "总利润"
       source_file: "经营数据.xlsx"
       sheet: "汇总"
       mode: header_match
       header: "利润"
     - key: stock_discrepancy
       label: "库存账实差异"
       mode: computed
       operation: subtract     # 目前只支持 subtract
       from: stock_balance     # 被引用的 KPI key（另一个 kpis 条目的 key）
       minus: port_actual_stock

   charts:
     - key: monthly_sales_trend
       type: line
       title: "月度销售趋势"
       source_file: "销售明细.xlsx"
       sheet: "月度"
       mode: header_match
       x_header: "月份"
       y_header: "销售额"

   tables:
     - key: sales_detail
       title: "销售明细"
       source_file: "销售明细.xlsx"
       sheet: "明细"
       mode: header_match
       columns: ["日期", "客户", "产品", "金额"]
     - key: outbound_list
       title: "出库明细"
       view_group: outbound     # 与下一项共享 view_group，前端渲染为一张卡片 + tab 切换
       view_label: "明细列表"
       source_file: "出库台账.xlsx"
       sheet: "出库明细"
       mode: header_match
       columns: ["完货时间", "船名", "航次", "装货港称重", "目的港称重", "发货人", "收货人"]
     - key: outbound_by_consignee
       title: "出库明细"
       view_group: outbound
       view_label: "按收货人排名"
       source_file: "出库台账.xlsx"
       sheet: "出库明细"
       mode: group_by_sum
       group_by_header: "收货人"
       sum_header: "目的港称重"
   ```

3. **Excel 解析模块**（`extractor.py` 等）
   - 输入：`config.yaml` + 上传的 Excel 文件。
   - 用 `pandas` / `openpyxl` 按配置里每一项的 `mode` 分别取数（固定范围 / 按表头匹配 / 分组汇总）；KPI 的 `computed` 模式不读 Excel，而是在其余 KPI 都取数完成后，按 `key` 引用已取到的值做运算。
   - 输出：一份标准化的 `data.json`，结构与前端渲染模块一一对应（`kpis` / `charts` / `tables` 三个顶层字段，各自是取数结果的数组）；共享同一 `view_group` 的多个表格条目各自作为独立的数组元素输出，前端负责按 `view_group` 归并展示。
   - 任一项取数失败（如找不到指定表头、`computed` 引用的 KPI 不存在）时，不应导致整个流程崩溃，应记录清晰的错误信息，其余项正常生成，方便维护人定位是哪个 Excel 出了问题。

4. **看板前端页面**（静态 HTML + JS + CSS）
   - 图表库使用 **ECharts**（对中文标签支持好，覆盖折线图/柱状图/饼图等常见图表类型）。
   - 页面加载时请求 `/api/data`，拿到 JSON 后依次渲染：KPI 卡片区、趋势图表区、占比/排名图表区、明细表格区（表格支持基本的排序）。关键的衍生型 KPI（如库存账实差异）与其他 KPI 一样出现在首屏 KPI 卡片区，不需要额外的独立区域。
   - 表格区渲染时按 `view_group` 归并：同组的多个表格条目合并为一张卡片，卡片内用 tab 控件在各视图间切换，默认显示 `view_group` 中第一个条目对应的视图；没有 `view_group` 的表格条目照常各自单独渲染一张卡片。
   - 页面本身不做特殊登录态处理（只读、无密码）。
   - 视觉与多端适配要求见下方独立章节。

5. **数据落地：`data.json` 单文件**
   - 不引入数据库，服务器本地一份 JSON 文件即可，每次上传解析后整体覆盖。
   - 预留扩展性：未来若要支持历史趋势，可将结构从"单快照"演进为"按日期存多份快照的数组"，前端读取逻辑改动可控，无需推翻整体设计。

## 视觉设计与多端适配要求

看板是给老板看的展示型页面，视觉呈现和数据本身同等重要。

- **视觉风格**：整体走简洁、专业的商务风格——清晰的层次感、克制的配色、留白得当，避免信息拥挤。KPI 数字、图表标题、图例等排版需要一眼能看懂重点，而不是罗列所有数字。具体的配色方案、字体、卡片样式在实施阶段用 `frontend-design` 技能产出，本文档只定基调和约束。
- **多端适配**：老板可能用手机、平板、PC 浏览器三种设备查看，页面必须是响应式的，而不是"电脑上做的页面缩小了看"：
  - 布局采用 CSS Grid/Flexbox + 媒体查询断点，大屏（PC）多列网格展示 KPI/图表/表格，小屏（手机）自动收窄为单列堆叠。
  - 图表组件（ECharts）需要监听容器尺寸变化并自适应重绘，不能在小屏上被压扁变形。
  - 明细表格在小屏上列数多容易溢出，需要有对应方案（如横向可滚动、或在窄屏下切换为卡片式逐行展示），不能出现需要横向拖动整个页面的情况。
- 这部分是纯前端展示层的要求，不影响后端架构（Flask 提供数据接口不变），实施计划阶段会把它拆成具体的页面开发任务。

## 数据更新流程

1. 维护人打开 `/upload`，输入密码进入上传页。
2. 选择本地一个或多个 Excel 文件，点击上传。
3. 服务器保存文件、按 `config.yaml` 解析、生成/覆盖 `data.json`，页面提示成功或具体报错（如"找不到表头'销售额'"）。
4. 老板（或任何人）打开 `/dashboard`，页面自动加载最新的 `data.json` 并展示。

## 技术栈

- 后端：Python + Flask
- Excel 解析：pandas + openpyxl
- 前端：静态 HTML/CSS/JS + ECharts（通过本地引入或 CDN）
- 数据存储：本地 JSON 文件（`data.json`），无需数据库
- 部署：本地可直接 `python app.py` 运行访问；未来迁移服务器时无需改动核心代码，只需调整运行环境

## v1 范围之外（明确不做）

- 历史趋势对比（同比/环比跨期数据）——架构预留扩展空间，但本次不实现。
- 在线文档（飞书/腾讯文档等）数据源接入——本次采用网页上传方案代替。
- 多用户账号体系——上传页只用单一共享密码，不做用户管理。
- 数据自动定时拉取/刷新——数据更新完全由维护人手动触发上传。

## 开放风险 / 后续观察点

- Excel 工作表结构若发生较大变化（如整体换了模板），`header_match` 模式能缓解但不能完全免疫，需要维护人在改动 Excel 结构时同步检查 `config.yaml`。
- 部署环境仍未最终确定；本设计已确保该决定可以推迟，不影响当前实现。
