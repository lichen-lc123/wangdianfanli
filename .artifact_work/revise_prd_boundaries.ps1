$ErrorActionPreference = 'Stop'

$businessPath = '03-产品设计/业务PRD/政策配置业务PRD.md'
$interactionPath = '03-产品设计/交互PRD/政策配置列表交互PRD.md'

$business = Get-Content -LiteralPath $businessPath -Raw -Encoding UTF8
$business = $business.Replace('**版本：V0.2**', '**版本：V0.3**')
$business = [regex]::Replace($business, '\*\*交付定位：.*?\*\*', '**交付定位：面向研发与编码智能体提供完整业务口径和验收标准；不规定接口、数据库、框架或代码结构。**', 1)
$business = $business.Replace('页面查询与分页必须由服务端执行，前端不得一次性加载全量政策后再筛选或分页。', '必须支持几十万条政策数据下的分页查询，页面只加载当前页所需数据；具体实现方式由技术设计确定。')
$businessBoundary = @'
## 8. 开发适配边界

| 边界项 | 要求 |
|---|---|
| 开发依据 | 研发人员或编码智能体必须实现本 PRD 已确认的业务规则、页面结果和验收标准 |
| 工程适配 | 开发前先读取现有工程，沿用项目已有的架构、接口规范、组件体系、权限和异常处理机制 |
| 数据规模 | 必须满足几十万条政策数据的查询和分页使用要求；具体前后端、存储与性能方案由技术设计确定 |
| 文档边界 | 本 PRD 不规定接口路径与方法、请求响应结构、错误码、数据库表与索引、技术栈、组件名称、状态变量或代码目录 |
| 缺口处理 | 现有工程能力不足或实现口径仍不明确时，先形成技术设计或待确认问题，不在产品 PRD 中补成技术事实 |
| 范围控制 | 不得根据本 PRD 自行扩展编辑、导入、删除、复制上月等本期范围外流程 |

'@
$business = [regex]::Replace($business, '(?s)## 8\. 全栈实现契约.*?(?=## 9\. 受影响的下游文档)', $businessBoundary)
$business = $business.Replace('| V0.2 | 2026-09-19 | 增加全栈接口、数据契约、错误码、性能与测试要求 |', "| V0.2 | 2026-09-19 | 增加面向研发的交付要求 |`r`n| V0.3 | 2026-09-19 | 移除具体技术选型，明确产品 PRD 与技术设计边界 |")
Set-Content -LiteralPath $businessPath -Value $business -Encoding UTF8

$interaction = Get-Content -LiteralPath $interactionPath -Raw -Encoding UTF8
$interaction = $interaction.Replace('**版本：V0.2**', '**版本：V0.3**')
$interaction = [regex]::Replace($interaction, '\*\*交付定位：.*?\*\*', '**交付定位：面向研发与编码智能体提供用户可见的页面行为和验收标准；不规定组件实现、接口形式或状态管理方案。**', 1)
$interaction = $interaction.Replace('`SingleCalendarDateRange`', '单日历日期区间选择器')
$interaction = $interaction.Replace('`SearchableMultiSelect`', '可搜索多选框')
$interaction = $interaction.Replace('`MultiSelect`', '多选框')
$interactionBoundary = @'
## 10. 开发适配边界

| 边界项 | 要求 |
|---|---|
| 实现目标 | 研发人员或编码智能体按本文实现页面结构、可见交互、校验提示、页面状态和验收标准 |
| 工程适配 | 开发前先读取现有工程，复用项目已有的页面框架、组件体系、接口规范、权限和异常处理方式 |
| 数据加载 | 页面只展示当前页数据，并能在几十万条政策数据下稳定查询和翻页；具体实现方案由技术设计确定 |
| 技术自由度 | 本文不规定组件名称、状态变量、接口路径与方法、请求响应结构、错误码、缓存策略或代码组织方式 |
| 数据来源 | 原型中的示例数据仅用于展示；正式环境接入项目已确定的数据来源 |
| 缺口处理 | 现有工程缺少必要能力时，先输出技术设计或待确认问题，不得把假设写成已确认产品规则 |

'@
$interaction = [regex]::Replace($interaction, '(?s)## 10\. 前端实现契约.*?(?=## 修订记录)', $interactionBoundary)
$interaction = $interaction.Replace('| V0.2 | 2026-09-19 | 增加全栈接口绑定、前端状态模型、性能与开发完成定义 |', "| V0.2 | 2026-09-19 | 增加面向研发的交付要求 |`r`n| V0.3 | 2026-09-19 | 移除具体技术选型，明确交互 PRD 与技术设计边界 |")
Set-Content -LiteralPath $interactionPath -Value $interaction -Encoding UTF8
