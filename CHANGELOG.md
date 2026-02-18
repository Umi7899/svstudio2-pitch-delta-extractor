# Changelog

## v1.0.0

- 新增脚本 `ExtractRenderedPitchToDelta.js`
- 支持将 AI 渲染音高提取到 `pitchDelta`
- 支持选中优先、边界外扩、采样档位
- 修复断开音符串联问题（休止区零锚点 + 音符边界锚点）
- 无可用渲染音高时自动零填充，不报错退出
