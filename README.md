# BioDraw（PowerPoint 科研绘图插件）

由 CaptainMus 开发的一款用于科研绘图的 PowerPoint 插件，欢迎使用！

项目地址：https://github.com/CaptainMusX/BioDraw

## 功能简介

- 素材库浏览与快速插入
- 图色替换（基于 ImageMagick）
- 颜色预设管理（含位置、Fuzz、透明替换）

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

在 `BioDraw -> 关于 -> 素材库` 中选择你的素材根目录，素材库最多支持二级嵌套文件夹。

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

---

## 开发者发布说明（简版）

1. 在 Visual Studio 使用 `Release | Any CPU`
2. 右键项目 -> 发布，版本设为 `1.0.0.0`
3. 生成 `setup.exe` 与 `BioDraw.vsto`
4. 推送代码并打 tag：`v1.0.0`
5. 在 GitHub Release 上传安装包资产
