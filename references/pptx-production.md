# PPTX 素材与渲染兼容

## 正文中文字体

使用 python-pptx 时，`run.font.name` 通常只写入 Latin 字体槽。需要明确指定中文字体时同时设置 East Asian 字体；字体名按目标机器已安装字体选择。

```python
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn

run.font.name = chosen_font
rpr = run._r.get_or_add_rPr()
ea = rpr.find(qn('a:ea'))
if ea is None:
    ea = OxmlElement('a:ea')
    rpr.append(ea)
ea.set('typeface', chosen_font)
```

如果文本在 XML 中存在，而渲染变成方框、缺字或不可见，先检查渲染器能否访问中文字体。不要仅靠增大文字框重复试错。

macOS 打包的 LibreOffice/Poppler 有时不能读取默认 Fontconfig 路径。可在工作区或临时目录建立配置，指向实际存在的系统字体目录和可写缓存，通过 `FONTCONFIG_FILE` 传给渲染进程。不安装/更改用户系统字体，也不写受限缓存目录：

```xml
<?xml version="1.0"?>
<fontconfig>
  <dir>/System/Library/Fonts</dir>
  <dir>/System/Library/Fonts/Supplemental</dir>
  <cachedir>/private/tmp/task-fontcache</cachedir>
</fontconfig>
```

这些路径只适用于相应 macOS 环境，其他系统应发现实际字体路径。渲染成功只能证明该渲染器效果，不能替代 Office 实机验收。

## 不继承意外的主题阴影

python-pptx 的矩形/连接线可能包含 `p:style/a:effectRef`，继承主题阴影。参考页要求扁平效果时，清除该引用或将效果索引设为 0，再实际渲染检查。只设线条为无填充不能清除阴影。

```python
style = shape._element.find(qn('p:style'))
if style is not None:
    effect = style.find(qn('a:effectRef'))
    if effect is not None:
        effect.set('idx', '0')
```

不要无条件删除需要保留的主题字体/填充配置。若本页所有样式都已明确指定，移除 `p:style` 也是可选做法。

## 裁切不等于清洁

- PowerPoint 的 `crop_left/right/top/bottom` 改变可见区域，嵌入原图仍完整存在。
- 可用裁切控制合法构图或复用左右景段，不能依靠裁切隐藏文字残片。
- 交付前提取嵌入原图全幅查看；不要只看 PPTX 的可见窗口。提取目录也适合检查重复图片、透明边缘和实际尺寸。
- 图片内本来应保留的文字（真实 UI、票据证据、作为独立素材的艺术标题）与误混入的正文区分处理。
- 生成透明图片后检查实际 alpha；黑色预览底不等于黑色图像底，棋盘纹也不等于真实透明。

## 来源和缺口

优先从用户 PPTX/DOCX 的媒体目录、图片原件提取高清素材。参考截图只有低分辨率时，可以保留其可用局部，或明确告知艺术素材已修复重建；不要把像素插值结果说成原生高分辨率。

若真实平台截图尚待用户补充，在交付 PPTX 中使用原生占位框与可理解的名字，例如 `待截图_评论互动`，记录所需截图、放置位置和裁切比例。用户提供截图后替换该对象，不生成假截图。

制作清单可使用这些字段：

`页码 / 对象名 / 类型（正文、几何、艺术字、图片、真实截图、占位） / 来源 / 处理（复用、裁切、修复、生成、待补） / 可编辑边界 / 待核验项`
