# 人工模板规则提炼与第二轮 AI 复判

## 模板可提炼的规则

1. 只共享行业通用词不能判为关联
- 如 `security`、`technology`、`telecom`、`services`、`group`、`international`、`eletronicos` 这类词，不足以证明是同一客户体系。

2. 核心品牌词一致时，可以接受更长的法人全称
- 如 `Tepillé -> TEPILLE SPA`
- 如 `G4S SECURE MONITORING AND RESPONSE PERU S.A.C -> G4s Peru S.A.C.`
- 如 `FG GROUP IT SAC / Fg Group IT / Grupo FG -> FG GROUP IT S.A.C. - GRUPO FG`

3. 法人后缀、国家后缀、标点和大小写差异通常可忽略
- 如 `SPA`、`S.A.C.`、`CIA.`、`LTDA.`、`E.I.R.L.`、`Peru`

4. 过短缩写不能单独作为强关联证据
- 如 `3S -> 3S STANDARD Sharing Software` 不能仅凭前缀重合就确认

5. 邮箱域名可以作为强辅助证据
- 模板中 `@tepille.cl`、`@grupofg.pe`、`@pe.g4s.com` 明显支持人工确认
- 当前 128 条复核数据里没有邮箱字段，所以本轮无法使用这条规则

## 第二轮 AI 复判范围

- 对第一轮 AI 结果中 `否` 和 `存疑` 的 102 条记录再检查一次
- 重点关注是否存在：
  - 品牌简称到完整法人名
  - 品牌词 + 城市名/地区名
  - 集团简称到标准账号名

## 第二轮复判后建议调整的记录

### 建议改为关联

1. row_id `55`
- 证书公司：`Vital`
- 当前映射：`Vital Tech International`
- 结论：`建议改为 是`
- 类型：`品牌简称 -> 完整法人名`
- 置信度：`0.78`
- 原因：`Vital` 是显著核心品牌词，`Tech International` 更像补充描述；同属 `Pakistan Group`，与模板里的 `Grupo FG -> FG GROUP IT ...` 属同类口径。

2. row_id `57`
- 证书公司：`Vital Karachi`
- 当前映射：`Vital Tech International`
- 结论：`建议改为 是`
- 类型：`品牌词 + 城市分支 -> 标准法人名`
- 置信度：`0.82`
- 原因：`Vital` 为核心品牌词，`Karachi` 更像本地分支或城市标记；同属 `Pakistan Group`，比第一轮更符合人工模板体现的“简称/本地实体 -> 标准账号”口径。

## 第二轮复判后仍不建议调整的情况

- 仅共享通用行业词
- 仅共享地区词或国家词
- 核心主体名称明显不同
- 只有很短缩写相同，但缺少独特品牌词支撑

## 结论

- 从这份拉美人工模板中，真正新增且可迁移的规则主要是：
  - `品牌简称 -> 完整法人名`
  - `品牌词 + 城市/地区标记 -> 标准法人名`
- 用这些新规则回看当前 128 条结果，明确值得翻案的主要是 `2` 条：
  - `row_id 55`
  - `row_id 57`
- 其余第一轮判 `否` 的记录，大多仍然只是共享行业通用词，不足以认定关联
