# BioDraw（PowerPoint 科研绘图插件）

由 CaptainMus 开发的一款用于科研绘图的 PowerPoint 插件，欢迎使用！

项目地址：https://github.com/CaptainMusX/BioDraw

## 视频教程

- 【BioDraw v1.0.0 功能演示】  
  https://www.bilibili.com/video/BV1tac6z2EJ3/?share_source=copy_web&vd_source=89f6f6134a704c2c421287c90d4a21a1

## 基础操作（含新增功能）

### 1) 素材库目录设置

- 在 `BioDraw -> 关于 -> 素材库` 中选择素材根目录。
- 素材库支持一级、二级目录结构：`类别 -> 子类 -> 素材文件`。
- 若某个类别下没有子目录，素材会直接显示该类别目录内的文件。

### 2) 素材浏览与插入

- 在“素材库”分组中选择 `类别`、`子类`，即可浏览对应素材。
- 预览区支持分页浏览，悬浮提示显示素材名称。
- 左键单击预览素材：插入到当前幻灯片。

### 3) 预览区快捷操作（新增）

- `Ctrl + 左键` 单击预览素材：删除该素材文件。
- `Alt + 左键` 单击预览素材：重命名该素材文件。
- 删除或重命名后会自动刷新预览列表与搜索缓存。

### 4) 类别/子类/搜索按钮（新增）

- “类别”按钮：当输入框名称不存在时，左键可创建一级目录；当目录已存在时，`Ctrl + 左键` 可删除该目录。
- “子类”按钮：当输入框名称不存在时，左键可创建二级目录；当目录已存在时，`Ctrl + 左键` 可删除该目录。
- “搜索”按钮：执行与搜索框回车相同的搜索操作。

### 5) 右键添加到素材库（含新增行为）

- 在 PPT 中选中对象后右键，可使用“添加到 BioDraw 素材库”。
- 支持常见对象类型：图片、原生形状、组合对象、多对象选择等。
- 多选对象右键添加时，会按对象逐个保存，不会合并成单个素材。
- 保存位置优先为当前预览所在二级目录；若不存在二级目录，则保存到当前一级目录。

### 6) 图色替换

- 选中图片后，设置“原色/新色”，点击“图片换色”执行替换。
- 支持“填充替换”与“透明替换”两种模式。
- `Fuzz` 控制容差，支持滑动与精确输入。
- 支持取色辅助，图色替换引擎由 ImageMagick 提供。

### 7) 预设管理

- 支持保存预设（原色、新色、模式、Fuzz、位置）。
- 支持设置默认预设、调整排序、删除预设。
- 即使不保存预设，也会记忆当前输入值。

## 安装前准备

请先确认以下环境：

- Windows 系统
- Microsoft PowerPoint（建议 Office 2016 及以上）
- .NET Framework 4.7.2（安装包会尝试自动安装）
- Visual Studio 2010 Tools for Office Runtime（VSTO Runtime，安装包会尝试自动安装）
- ImageMagick（图色替换功能必需）

---

## 第一步：安装 ImageMagick（必须）

BioDraw 的“图色替换”功能会调用 `magick` 命令，所以你必须先安装 ImageMagick，并确保命令可在终端直接运行。

### 1. 下载

1. 打开 ImageMagick 官网下载页：  
   https://imagemagick.org/script/download.php#windows
2. 选择 Windows 版本（通常选 64-bit 动态版即可）。

### 2. 安装时的关键选项

安装向导中请确保：

- 勾选将 ImageMagick 安装目录加入系统 PATH（Add application directory to your system path）
- 安装完成后重启一次 PowerPoint（或重启电脑）

### 3. 验证是否安装成功

打开 PowerShell，执行：

```powershell
magick -version
```

如果能看到版本信息，表示安装成功。  
如果提示“找不到 magick 命令”，请把 ImageMagick 安装目录加入系统环境变量 `Path`，然后重新打开 PowerShell 再试。

---

## 第二步：安装 BioDraw 插件

发布的安装包中，通常包含：

- `setup.exe`（推荐双击这个）
- `BioDraw.vsto`
- `Application Files` 目录

### 安装步骤

1. 双击 `setup.exe`
2. 按向导完成安装
3. 打开 PowerPoint，检查顶部是否出现 `BioDraw` 选项卡

如果出现安全提示，请选择信任发布者后继续安装。

---

## 第三步：首次使用建议

### 1. 设置素材库路径

在 `BioDraw -> 关于 -> 素材库` 中选择你的素材根目录，素材库支持一级与二级目录结构。

### 2. 验证“图色替换”

1. 在 PPT 中选中一张图片
2. 在 `图色替换` 区域输入原色/新色（或点“取色”）
3. 点击“换色”

如新色为空，会按透明替换处理。

---

## 常见问题排查

### 1) 图色替换失败，提示无法启动 ImageMagick

原因：系统找不到 `magick` 命令。  
处理：

- 先执行 `magick -version` 验证
- 若失败，补充 PATH 后重启 PowerPoint

### 2) 安装后 PowerPoint 没看到 BioDraw

处理：

- 打开 PowerPoint -> 文件 -> 选项 -> 加载项
- 底部“管理 COM 加载项” -> 转到
- 确认 `BioDraw` 已勾选启用

### 3) 被 Office 禁用

处理：

- 文件 -> 选项 -> 加载项 -> 已禁用项目
- 将 BioDraw 恢复启用

---

## 卸载方式

- Windows 设置 -> 应用 -> 已安装应用（或控制面板 -> 程序和功能）
- 找到 `BioDraw` 后卸载
