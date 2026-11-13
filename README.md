# PPT原稿接续与逐页精修 Skill

**版本：1.0.0｜整理日期：2026-11-13**

保留原稿、接续版本，逐页精修已有演示文稿。适用于职业规划、食品创新、创新创业及其他演示项目。

## 包内有什么

- `SKILL.md`：触发条件、执行顺序、硬约束及按需阅读入口。
- `references/`：完整流程、三条工作线、设计与证据规则、项目适配、质量门槛、回归测试。
- `templates/`：项目合同、逐页映射、要求记录、素材与数据记录、智能体任务单、交接与验收模板。
- `scripts/`：原稿保护/工作区初始化脚本、只读PPTX预检脚本。
- `tests/`：脚本自测与对话场景验收。

## 如何使用

在能够加载自定义Skill的工作环境中，交给该环境按其支持的方式加载此文件夹。不同宿主的导入方式可能不同；**提供此包不等于已经安装到你的账号，也不会自动恢复旧智能体或同步到其他设备。**

在新对话中也可以上传本压缩包和待修改PPT，明确要求先读取`SKILL.md`，再按它执行。仅上传文件本身不代表宿主一定会自动加载Skill。

可直接发送：

> 使用《PPT原稿接续与逐页精修》Skill修改这套PPT。先读取SKILL.md，以我确认的可编辑版本为母版，以最新图片版核对页序；按本轮逐页要求修改，不从头制作，不使用AI生图。保留已确认字体、照片与艺术字，检查后同步交付可编辑版和50MB以内的图片版。能调用子智能体时按包内三条工作线分工，不能调用就由主执行者执行并说明实际状态。

## 默认锁定的经验

普通标题沿用清晰的微软雅黑加粗；封面与尾页艺术字独立保留。改字体不回退文案，改标题不叠加标签，改照片不顺带更换其他素材。页序按内容匹配，同一可编辑主稿导出两个版本。所有图片证据用真实素材，数据不以缺失为由编造；估算与实测分开呈现。

项目明确要求优先，按作用范围更新合同，不从其他项目继承页数、标识、配色或能力模型。

## 三条工作线

按需分工为数据核验、内容审校和视觉编辑，详见 [分工说明](references/agent-roles.md)。

## 可选脚本

需要Python 3.9或更高版本，脚本仅使用标准库，不安装网络依赖，不调用OCR、图像生成或外部服务。

```bash
# 复制原稿并建立隔离工作区。目标目录必须不存在。
python3 scripts/init_revision.py \
  --source "/path/to/source_editable.pptx" \
  --order-reference "/path/to/order_reference.pptx" \
  --project-dir "/path/to/new_revision_workspace"

# 结构检查，不改PPT。
python3 scripts/pptx_preflight.py "/path/to/final_editable.pptx" \
  --report "/path/to/qc/editable.json"

# 图片版检查；50MB为50,000,000字节。
python3 scripts/pptx_preflight.py "/path/to/final_images.pptx" \
  --report "/path/to/qc/images.json" --max-mb 50 \
  --compare "/path/to/final_editable.pptx"

# 自测只使用临时构造的结构样本，不修改用户PPT。
python3 -m unittest discover -s tests -v
```

`pptx_preflight.py`核对压缩包结构、可见页顺序、文本/图片对象、声明字体、基础越界候选、内部文件引用、字节数及两版页数/页面尺寸。它**不渲染、不判定事实正确、不检查字体安装，也不能自动认定所有重叠、图像清晰度和可编辑性合格**。详细能力及边界见`references/quality-gates.md`。


## 开源许可证

本仓库按 [MIT 许可证](LICENSE) 开源。
