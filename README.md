# 采购发票对账工具

Excel 台账 × 发票 PDF，按发票号精确匹配，产出 outcome.xlsx。

## 安装

    pip install -r requirements.txt

## 用法

界面版（推荐给非技术同事）：

    streamlit run app.py

命令行批量：

    python reconcile.py "Neocytogen Procurement Tracking List (2026).xlsx" "./2026 Jan/" "outcome.xlsx"

## 文件

| 文件 | 作用 |
|---|---|
| `excel_loader.py` | Excel 台账清洗与标准化 |
| `invoice_extractor.py` | PDF 发票字段抽取 |
| `reconcile.py` | 匹配主逻辑 + Excel 输出 |
| `app.py` | Streamlit 界面（仅界面层，换 Zoho 时替换这一层） |

## 关键设计

- **发票号只做精确匹配。** 同供应商的发票号可能只差几位（97809367 / 97809395），
  模糊匹配一定配错。相似度只用于给人看的候选提示。
- **金额只在发票合计层面校验**，容差 0.05。各家发票行级 GST 舍入方式不同。
- **Excel 直读，不经过模型。** 数字必须确定性读取。
- **状态为 Ordered / Pending 的行不算 mismatch**，单独进 pending sheet。
- **跨年订单**：PO 号为上一年度的（如 PONCG2025xxxxx），Note 会提示核对上一年台账。
- `reconcile()` 返回纯数据结构，`write_output()` 是独立的输出层。
  迁移到 Zoho 时只替换输出层和界面层，匹配逻辑不动。
