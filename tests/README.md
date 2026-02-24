# tests 目录说明

本文档对应当前仓库里的测试文件结构（`tests/`）。

## 运行方式

```bash
# 默认推荐：执行 tests/ 下全部测试（默认排除 ppt_com）
python -m pytest tests -q

# 查看详细输出
python -m pytest tests -v

# 仅跑慢测
python -m pytest tests -v -m slow

# 排除慢测（显式保留排除 ppt_com）
python -m pytest tests -v -m "not slow and not ppt_com"

# 仅跑 PPT COM 集成测试（需 Windows + PowerPoint + pywin32）
python -m pytest tests/test_ppt_legacy.py -v -m ppt_com
```

说明：
- `pyproject.toml` 中默认 `addopts` 包含 `-m "not ppt_com"`，所以常规运行不会执行 COM 集成测试。
- 一旦你手动传 `-m ...`，会覆盖默认 marker 过滤；建议总是显式写上 `and not ppt_com`。
- 为适配当前环境，pytest 的 `tmpdir` 插件已禁用，统一使用 `tests/conftest.py` 中的临时目录夹具。

## 测试文件与覆盖点

| 文件 | 说明 |
|---|---|
| `tests/test_smoke.py` | 样本 PPTX 冒烟测试（结构完整性、格式变体、确定性） |
| `tests/test_extraction.py` | 关键内容提取断言（标题、列表、表格、公式、图片、中文文本） |
| `tests/test_notes_extraction.py` | 备注区列表（无序/有序）提取行为 |
| `tests/test_ordered_list_start_at.py` | 有序列表 `startAt` 编号行为 |
| `tests/test_wmf_com_fallback.py` | WMF COM 兜底使用绝对路径的单测 |
| `tests/test_ppt_routing.py` | `.ppt/.pptx` 路由和配置构建单测 |
| `tests/test_ppt_legacy.py` | `.ppt` COM 路由集成测试（`@pytest.mark.ppt_com`） |
| `tests/test_smoke_history_compare.py` | 历史冒烟输出对比（按 mtime 比较最新与上一次） |

## 测试资源

`tests/conftest.py` 默认使用以下样本：

- `test_pptx/6.人工智能前沿应用场景2025.pptx`
- `test_pptx/3 深度学习概览2025.pptx`

其中慢测会进行真实转换，运行时间和机器性能、磁盘性能相关。
若本地不存在 `test_pptx/` 目录或样本文件，依赖样本的慢测会失败。

## 历史输出对比说明

`tests/test_smoke_history_compare.py`：
- 会在 `test_output/` 中按文件修改时间选择“最新”和“上一次”输出进行比较。
- 默认非严格模式：有差异时写 diff 并 `skip`。
- 严格模式：设置 `PPTX2MD_SMOKE_HISTORY_STRICT=1` 后，有差异直接 `fail`。

## 常见问题

- `ppt_com` 测试被跳过：通常是缺少 Windows/PowerPoint/pywin32。
- 慢测耗时长：可先跑 `-m "not slow and not ppt_com"` 做快速回归。
- 慢测直接报 `Test sample not found`：请先准备 `test_pptx/*.pptx` 样本。
- 临时目录堆积：测试临时文件会保留在 `tmp_test_artifacts/`，可按需手动清理。
