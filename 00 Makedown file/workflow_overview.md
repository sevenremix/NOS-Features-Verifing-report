# NOS Feature & Case Toolchain Workflow

本项目通过 `main.py` 驱动一个多步骤的自动化流水线，将原始功能列表与测试报告进行关联，最终生成可视化的合并报告。

## 流程概览

```mermaid
graph TD
    %% 开始节点
    Start([开始运行 main.py]) --> Step1[运行 Transform_simple_copy.py]

    %% 第一步：预处理
    subgraph "1. 数据预处理 (Preprocessing)"
        Step1 --> In1[/"输入: NCS520 and UAR600D Feature Compare-1226.xlsx"/]
        In1 --> Proc1["提取 UAR600D-10XA Feature 页签<br/>进行格式化归一化"]
        Proc1 --> Out1[/"输出: Formatted_Feature_Compare.xlsx"/]
    end

    %% 第二步：校验
    subgraph "2. 数据校验 (Verification)"
        Out1 --> Step2[运行 check_consistency.py]
        Step2 --> Comp["对比原始文件与格式化文件"]
        Comp -- "检测不一致" --> Warn["输出警告/错误信息"]
        Comp -- "数据一致" --> Step3
    end

    %% 第三步：合并
    subgraph "3. 报告合并 (Merging)"
        Step3[运行 Feature_Case_Merging.py]
        Step3 --> In2[/"输入: Formatted_Feature_Compare.xlsx"/]
        Step3 --> In3[/"输入: 最新测试报告 (UAR600D-10XA...xlsx)"/]
        In2 & In3 --> Proc3["关联 Feature 与 Test Case<br/>计算通过率与统计数据"]
        Proc3 --> Out2[/"最终输出: Feature_TestCase_YYYYMMDD.xlsx"/]
    end

    %% 结束节点
    Out2 --> End([流程完成])

    %% 样式美化
    style Start fill:#f9f,stroke:#333,stroke-width:2px
    style End fill:#f9f,stroke:#333,stroke-width:2px
    style Out2 fill:#ff9,stroke:#333,stroke-width:4px
```

## 详细步骤说明

### 1. 驱动程序 (`main.py`)
- **作用**: 流水线控制器。
- **逻辑**: 使用 `subprocess` 按序调用各个 Python 脚本，并监控执行状态。如果某步失败，流程将立即中止。

### 2. 预处理 (`Transform_simple_copy.py`)
- **输入**: `NCS520 and UAR600D Feature Compare-1226.xlsx`
- **处理**:
    - 锁定 `UAR600D-10XA Feature` 工作表。
    - 对 `描述` 列进行向下填充（ffill），确保每个功能点都有所属分类。
    - 应用特定的样式（边框、加粗、颜色）。
- **输出**: `Formatted_Feature_Compare.xlsx`

### 3. 一致性检查 (`check_consistency.py`)
- **作用**: 风险控制。
- **逻辑**: 对比原始文件和生成的格式化文件中的行数、`规格`、`UT NOS` 等关键字段，确保在转换过程中没有丢失或改动数据。

### 4. 核心合并 (`Feature_Case_Merging.py`)
- **输入**:
    - `Formatted_Feature_Compare.xlsx` (功能清单)
    - 最新日期命名的测试报告 (测试数据)
- **处理**:
    - **鲁棒读取**: 直接解析 XLSX 的 XML 结构，绕过复杂的公式和格式干扰。
    - **匹配算法**: 按照功能名称将测试用例关联到对应的 Feature。
    - **统计计算**: 自动计算每个 Section 以及全局的测试通过率 (Pass Rate) 和 CLI 覆盖率。
    - **可视化导出**: 生成带有分级折叠（Outline）、背景颜色标识、冻结窗格的 Excel 报告。
- **输出**: `Feature_TestCase_YYYYMMDD.xlsx`
