# Extract Rendered Pitch to Pitch Delta

Synthesizer V Studio 2 Pro 脚本。  
功能：将当前编辑组中 AI 自动渲染得到的音高曲线提取到 `Pitch Deviation (pitchDelta)` 参数中。

## 功能特性

- 选中音符优先（无选中时回退到当前组全部音符）
- 覆盖写入目标区间的 `pitchDelta`
- 自动处理断开音符之间的休止区：休止区强制锚定为 `0`
- 对每个音符补边界锚点，避免前后音符曲线串联导致整体偏移
- 当目标区间无可用渲染音高（如歌词为空、试用声库限制）时自动写 `0`，不中断脚本

## 安装

1. 将 `ExtractRenderedPitchToDelta.js` 放到：
   `Synthesizer V Studio 2/scripts/pitch/`
2. 重启 Synthesizer V Studio 2。

## 使用

1. 在编辑器中打开目标组（并可选中部分音符）。
2. 运行脚本：`Extract Rendered Pitch to Pitch Delta`。
3. 在弹窗中设置：
   - Scope（范围）
   - Sampling Profile（采样档位）
   - Padding Mode（边界策略）
4. 确认后脚本会提取并写入 `pitchDelta`。

## 参数说明（默认）

- Scope: `Selected Notes First (fallback to all notes)`
- Sampling Profile: `Balanced (5ms, low simplify)`
- Padding Mode: `Auto extend by 1/16 quarter`
- Pure AI baseline: 强制开启（先清空目标区间旧 `pitchDelta`）

## 注意事项

- 本脚本不会切换声库，只负责提取曲线。
- 若同一个 `NoteGroup` 在多个位置复用，脚本会提示并在确认后继续（会影响该组所有引用）。

## 版本

- `v1.0.0` 首次发布。
