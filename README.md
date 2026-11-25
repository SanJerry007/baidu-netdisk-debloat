# baidu-netdisk-debloat

一个脚本给 Windows 版百度网盘 8.2.1.102 关掉强制升级、断开日志和崩溃上传、拿掉 Office 插件，网盘照常用。 Debloat Baidu Netdisk 8.2.1.102 on Windows: no forced updates, no log upload, no Office add-in, no binary patching.

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/Platform-Windows%2010%20%2F%2011-blue?style=flat)](#用法)
[![PowerShell 5.1+](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=flat)](#用法)
[![Tested: Baidu Netdisk 8.2.1.102](https://img.shields.io/badge/Tested-Baidu%20Netdisk%208.2.1.102-green?style=flat)](docs/how-it-works.md)
[![No binary patching](https://img.shields.io/badge/Binary%20Patching-None-green?style=flat)](docs/how-it-works.md)
[![Auto re-apply](https://img.shields.io/badge/Auto%20Re--Apply-Logon%20%2B%20Every%206h-green?style=flat)](#装成自动模式)
[![Release](https://img.shields.io/github/v/release/SanJerry007/baidu-netdisk-debloat?style=flat)](https://github.com/SanJerry007/baidu-netdisk-debloat/releases)
[![Languages](https://img.shields.io/badge/Languages-%E4%B8%AD%E6%96%87%20%2F%20EN-blue?style=flat)](#english-summary)
[![Last commit](https://img.shields.io/github/last-commit/SanJerry007/baidu-netdisk-debloat?style=flat)](https://github.com/SanJerry007/baidu-netdisk-debloat/commits)
[![Stars](https://img.shields.io/github/stars/SanJerry007/baidu-netdisk-debloat?style=flat)](https://github.com/SanJerry007/baidu-netdisk-debloat/stargazers)

[中文](#baidu-netdisk-debloat) | [English](#english-summary)

---

给 Windows 版百度网盘做减法。默认只做几件改动小、证据足的事：让客户端更新程序和传输内核更新程序起不来，挡住崩溃转储和日志上传，拿掉会把自己注册进 Word、Excel、PowerPoint 的 Office 插件。右键菜单、同步盘图标、BT 入站端口、推送通知、浏览器扩展、打开方式、语音和 OCR 模块，各有一个开关，要不要动由你决定。

不改任何二进制。默认动作把更新程序和 Office 插件 DLL 换成同名空目录，`AutoUpdate` 目录对当前用户拒绝写入，给上传程序加出站防火墙规则；可选项另外会把几个文件和目录换成占位、加防火墙规则，或者改注册表、系统策略和服务启动类型。每一步先记日志再动手，除了 `-RemoveStagedUpdate` 删掉的更新包，都能撤销。原理、启动链和测试记录见 [docs/how-it-works.md](docs/how-it-works.md)。

只支持 8.2.1.102 这一个版本，其他版本脚本会拒绝执行，见[版本支持](#版本支持)。主窗口里的广告和 AI 入口去不掉，原因见[去不掉的东西](#去不掉的东西)。

## 会去掉什么

| 功能 | 怎么处理 | 默认还是选项 |
|---|---|---|
| 客户端自动升级：每次启动约 30 秒后先更新自己，再下载约 800 MB 的强制更新包，下次启动就装上 | `AutoUpdate\Autoupdate.exe` 换成同名空目录，`AutoUpdate` 目录对当前用户拒绝写入 | 默认 |
| 传输内核更新程序 `kernelUpdate.exe`，常驻 | 换成同名空目录 | 默认 |
| 崩溃转储上传 `crashpad_handler.exe`，带账号标识 | 出站防火墙规则 | 默认 |
| 日志上传 `HelpUtility.exe`：启动约 10 分钟后上传客户端日志和登录日志，登录令牌明文写在它的命令行里 | 出站防火墙规则 | 默认 |
| Office 插件 `YunOfficeAddin.dll`、`YunOfficeAddin64.dll`：会把自己注册进 Word、Excel、PowerPoint，每次启动 Office 都加载 | 两个 DLL 换成同名空目录 | 默认 |
| 已经下载好、等着安装的强制更新包 | 删掉，留一个只读空文件 | `-RemoveStagedUpdate` |
| 语音插件 `module\SpeechSDK` | 目录换成只读空文件，原目录先备份 | `-RemoveSpeechSdk` |
| 本地 OCR 模型 `module\AiEngine` | 目录换成只读空文件，原目录先备份 | `-RemoveAiEngine` |
| SYSTEM 服务 BaiduNetdiskUtility | 启动类型改为禁用 | `-DisableUtilityService` |
| 崩溃报告程序 `BugReport.exe` | 出站防火墙规则 | `-BlockBugReport` |
| 诊断更新程序 `autoDiagnoseUpdate.exe` | 换成同名空目录 | `-BlockDiagnoseUpdater` |
| Office 插件的第二道防线 | Office 策略把三个百度插件设为始终禁用 | `-BlockOfficeAddinPolicy` |
| 同步盘的 7 个图标覆盖，抢在别的同步软件前面占掉覆盖槽位 | 删除覆盖标识，CLSID 加入资源管理器黑名单 | `-RemoveOverlays` |
| 内置 BT/P2P 引擎的入站端口（TCP/UDP 51413、UDP 8568），厂商的防火墙规则在公用网络上也全部放行 | 给 4 个程序加入站阻止规则 | `-BlockInboundP2P` |
| 推送和通知通道 BaiduNetdiskService（商店应用） | 关掉它的系统通知，占住它的安装包 | `-MuteNotifyService` |
| 附带的 AI 包 `sysres\BaiduNetdiskAI.msix` | 换成同名空目录，没法再注册 | `-RemoveAiMsix` |
| 右键菜单里的「baidunetdisk」 | 隐藏经典菜单项，处理程序加入黑名单 | `-HideContextMenu` |
| 26 种音视频和 PDF 类型「打开方式」里的百度网盘，「默认应用」里的网盘看图 | 删除对应的注册表项 | `-CleanOpenWith` |
| 附带的 Chrome、Edge 扩展（能读 cookies、监听网络请求、作用于所有网站） | 浏览器的扩展黑名单策略 | `-BlockBrowserExtension` |

## 会保留什么

- 网盘本身：登录、文件列表、上传、下载、同步。主程序 `BaiduNetdisk.exe`、界面 `BaiduNetdiskUnite.exe`（Electron）和传输内核（`baidunetdiskhost.exe` 里加载的 `kernel.dll`）一概不碰。
- `YunDetectService.exe`：在网页上点下载时，由它把任务交给客户端。它只监听本机 127.0.0.1:10000。
- 右键上传和同步盘：默认保留。用不到的话加 `-HideContextMenu`、`-RemoveOverlays`。
- 看图和视频预览（ImageViewer、VastPlayer）。
- 内置 BT 引擎：它就在传输内核里，拿掉会连下载一起坏，只能用 `-BlockInboundP2P` 挡住外面连进来的流量。
- 本地 `users\` 目录里的设置和登录状态。账号数据本来就在服务器上。

设置里的「开启自动升级」不用管。关掉它没有用，更新程序一分钟内就会把它改回去。脚本的办法是让更新程序本身起不来。

代价：客户端停在 8.2.1.102，它自带的 Chromium（`BaiduNetdiskUnite.exe` 8.2.1.352）和 BT 库（Transmission 4.0.6）都收不到安全更新。服务器现在还接受 8.2.1.102，哪天不接受了，先 `-Uninstall` 让它升级。

## 去不掉的东西

**主窗口里的广告和推销。** 广告位、会员推销、玩游戏换加速、看广告换次数、领券、托盘推销文字、AI 侧栏和 AI 搜索栏，都是界面（`app.asar`）按 `pan.baidu.com` 接口返回的内容画出来的。本地没有开关，设置文件又是加密的。脚本不改 `app.asar`，所以它们都还在。网络层也挡不干净：界面里除了 `pan.baidu.com` 和百度的 CDN，找不到单独的广告域名，屏蔽它们会连网盘一起断。

**核心进程里跑的东西。** 下面这些就在主程序或传输内核里，没有能单独拦的地方：
- 启动时读一次剪贴板，取出搜索广告的点击编号（`semclickid=`）上报，再把这一段从剪贴板里删掉写回。
- 上传测速，一次约 115 MB，配置的间隔是 12 小时。
- 加密的遥测队列和内核统计。
- BT 引擎向路由器申请 UPnP、NAT-PMP 端口映射。`-BlockInboundP2P` 挡得住连进来的流量，挡不住这个出站请求，只有路由器能拒绝。
- `%USERPROFILE%\Documents\BaiduNetdiskTmp` 里不断堆积的临时文件（主要是 `*.tmp.node`，两天能攒到五六十 MB）。

**只能在客户端里自己改的设置。** 这些设置存在加密文件里或者账号里，脚本改不到。路径按 8.2.1.102 的界面整理：

| 设置 | 位置 | 建议 |
|---|---|---|
| 开机自动启动 | 设置 > 启动设置 > 开机自动启动 | 关。打开同步盘或自动备份时它会被顺手打开，之后再看一眼 |
| 在我的电脑中显示网盘 | 设置 > 启动设置 > 在我的电脑中显示网盘 | 关 |
| 启动后关联BT种子文件 | 设置 > 启动设置 > 启动后关联BT种子文件 | 关 |
| 显示悬浮窗 | 设置 > 启动设置 > 显示悬浮窗 | 关 |
| 截图快捷键（全局热键） | 设置 > 启动设置 > 截图快捷键 | 关 |
| 设置「智能看图」为系统图片默认查看方式 | 设置 > 启动设置；每次更新后的引导页里也默认勾着 | 取消勾选 |
| 快捷上传插件、网盘运行时启动插件 | 设置 > 快捷上传 | 关 |
| 自动识别剪切板里的链接 | 设置 > 通知设置 > 其他提示 | 关 |
| 订阅更新弹窗提醒 | 设置 > 通知设置 > 其他提示 | 关 |
| 商品推荐 | 设置 > 隐私设置 > 商品推荐 | 关 |
| 智能化功能授权 | 设置 > 隐私设置 > 智能化功能授权 | 关 |
| 智能问答（用 AI 总结文件内容） | 设置 > 智能问答 | 关 |
| 云一朵（AI 助手） | 设置 > 云一朵设置 | 关 |
| 首页推荐卡片 | 设置 > 展示设置 | 少留几类 |
| 系统兼容模式 | 设置 > 系统设置 > 系统兼容模式 | 别勾，它会改系统电源计划里的处理器加速模式 |
| 下载加速 | 传输列表 > 下载加速 | 别开，它会把你的上传带宽借给别人（P2P 上传共享） |

## 用法

需要 Windows PowerShell 5.1 或更新的版本，执行时要有管理员权限。从 [Releases](https://github.com/SanJerry007/baidu-netdisk-debloat/releases) 下载 `BaiduNetdiskDebloat.ps1`，或者克隆本仓库。

动手前先在托盘图标上右键，退出百度网盘。

```powershell
# 演练：只列出会改什么，不需要管理员权限
powershell -ExecutionPolicy Bypass -File .\BaiduNetdiskDebloat.ps1

# 执行一次（在管理员 PowerShell 里）
powershell -ExecutionPolicy Bypass -File .\BaiduNetdiskDebloat.ps1 -Apply

# 执行，并在每次登录和每 6 小时自动重做（在管理员 PowerShell 里）
powershell -ExecutionPolicy Bypass -File .\BaiduNetdiskDebloat.ps1 -Install

# 选项直接加在后面
powershell -ExecutionPolicy Bypass -File .\BaiduNetdiskDebloat.ps1 -Install -RemoveStagedUpdate -RemoveSpeechSdk -RemoveAiEngine
```

演练会打印找到的安装目录和找法、通过了哪项厂商校验、完整的改动计划，最后列出没选的选项、各自会启用哪些动作、证据等级。`-Apply` 照着这份计划执行。脚本可以反复跑，已经处理过的项显示 `SKIP`。

退出码：`0` 正常；`1` 有步骤失败，或者 `-Uninstall` 没能全部还原；`2` 开始前就拒绝了（同时给了几个模式、没有管理员权限、没装百度网盘、厂商校验没过、版本不支持），什么都没改。

建议的顺序：先演练；再 `-Apply`；启动客户端，确认能登录、能看到文件列表，上传和下载各试一次；最后用同样的选项跑 `-Install`。

客户端如果已经提示过更新，强制更新包可能已经下好，下次启动就会装上。先跑完脚本再启动客户端；想顺手删掉这个包，加 `-RemoveStagedUpdate`。

### 选项

全部默认关闭，按需加在命令后面：

- `-RemoveStagedUpdate`：删掉已经下好、等着安装的强制更新包 `AutoUpdate\Download\MainApp`（约 800 MB），留一个只读空文件。`-Uninstall` 会删掉这个占位文件，但更新包不备份，更新程序回来后重新下载。
- `-RemoveSpeechSdk`：语音插件目录 `module\SpeechSDK`（插件 1004）换成只读空文件，原目录先备份。
- `-RemoveAiEngine`：本地 OCR 目录 `module\AiEngine`（onnxruntime 和 OCR 模型）换成只读空文件，原目录先备份。
- `-DisableUtilityService`：禁用 SYSTEM 服务 BaiduNetdiskUtility。它的程序文件在用户可写的目录里，普通用户就能启动它。禁用后上传下载是否完全正常还没测过。
- `-BlockBugReport`：给崩溃报告程序 `BugReport.exe` 加出站阻止规则。
- `-BlockDiagnoseUpdater`：`autoDiagnoseUpdate.exe` 换成同名空目录。没见过它启动，属于预防。
- `-BlockOfficeAddinPolicy`：在 Word、Excel、PowerPoint 的用户策略 `Resiliency\AddinList` 里把三个百度插件设为始终禁用（Office 16.0），别处的副本把自己注册进去也加载不了。
- `-RemoveOverlays`：删掉同步盘的 7 个图标覆盖标识（64 位和 32 位注册表视图），CLSID 加入黑名单，把覆盖槽位还给其他同步软件。百度同步盘的状态图标会消失。重启资源管理器或重新登录后生效。
- `-BlockInboundP2P`：给 `BaiduNetdisk.exe`、`BaiduNetdiskRender.exe`、`baidunetdiskhost.exe`、`BaiduNetdiskUnite.exe` 加入站阻止规则，盖过厂商的全放行规则，网络上的人连不进内置 BT/P2P 引擎。出站传输和本机回环不受影响；对 P2P 加速下载的速度有没有影响还没测过。
- `-MuteNotifyService`：关掉商店应用 BaiduNetdiskService（推送和通知通道）的系统通知，并占住安装目录里它的安装包 `BaiduNetdiskService.Msixbundle`。传输完成的通知可能也一起没了。已经装上的应用不删，想删去 Windows 的「设置 > 应用」。
- `-RemoveAiMsix`：附带的 AI 包 `sysres\BaiduNetdiskAI.msix`（看图程序调用 Windows 本地 AI 模型用的身份）换成同名空目录，没法再注册。
- `-HideContextMenu`：隐藏右键菜单里的「baidunetdisk」。经典菜单项加 `LegacyDisable`，处理程序 `{499C8928-90D7-411D-8DBD-E35B92529123}` 加入黑名单；Windows 11 顶层菜单里那一项能不能一起藏掉还没测过。右键上传也就没了。
- `-CleanOpenWith`：从 26 种音视频和 PDF 类型的「打开方式」里删掉百度网盘，从「默认应用」里删掉网盘看图。现有的默认程序不变。
- `-BlockBrowserExtension`：用 Chrome、Edge 的 `ExtensionInstallBlocklist` 机器策略封掉附带扩展 `lfkkoabmfeaefiddnpbjeiofkmkpogim`，它的侧载注册表项在每次运行时发现就删（装成自动模式后是登录时和每 6 小时）。策略存在期间，两个浏览器都会显示「由你的组织管理」。

### 其他参数

- `-InstallRoot <路径>`：不自动查找，直接用这个安装目录。它照样要通过文件和签名校验。
- `-Force`：版本不在支持列表里也照样执行。没人测过，不建议用。
- `-IgnoreUninspectable <进程名>`：检查模块占用时，跳过这些无法检查的进程。只管这一次运行。
- `-Auto`：计划任务专用，不用手动跑。

## 装成自动模式

`-Install` 先执行一遍。全部成功才继续：把脚本和这次用的选项复制到 `%ProgramData%\AppDebloat\baidunetdisk`（只有管理员能改），再注册计划任务 `BaiduNetdiskDebloat-Watch`，每次登录和每 6 小时以你的身份、管理员权限、无窗口地跑一次 `-Auto`。有任何一步失败就不注册，并说明原因。

`-Auto` 只重做拦截类的动作。什么都没变时全是 `SKIP`；占位被删了、拒绝写入项被去掉了、服务被改回来了，它就补回去。最近一次运行的完整输出在 `auto.log`（每次覆盖），每次运行追加一行摘要到 `history.log`，都在 `%ProgramData%\AppDebloat\baidunetdisk`。看结果以 `history.log` 里的 `failure(s)` 数为准，任务计划程序里的「上次运行结果」不反映失败。

发现装的不是 8.2.1.102 时，`-Auto` 什么都不改，在 `history.log` 里记一行就退出。

想换选项，带上新的选项重新跑一次 `-Install`。想撤掉一个已经生效的选项，先 `-Uninstall`，再按新的选项 `-Install`，两步之间别启动客户端。

脚本不会从网上自动更新自己，让一个管理员权限的任务执行网上下载的代码风险太大。新版本发布后，下载下来重新跑一次 `-Install` 即可。

## 撤销

```powershell
# 删掉计划任务，按日志倒序还原所有改动（在管理员 PowerShell 里）
powershell -ExecutionPolicy Bypass -File .\BaiduNetdiskDebloat.ps1 -Uninstall
```

- 只还原那些现在还是脚本写的样子的项。被别的东西改过的标成 `CONFLICT`，不动。
- 有 `WARN`、`CONFLICT` 或失败时，日志和状态目录都留着，退出码 `1`，处理完再跑一次。全部还原后，状态目录会被删掉。
- `-RemoveStagedUpdate` 留下的只读空文件会被删掉，但更新包没有备份，还原不了。撤销结束时它会列在「只能覆盖重装找回」的文件里；不用重装，更新程序回来后会重新下载。
- 还原后启动客户端，更新程序会照常升级。也可以直接用官方安装包覆盖安装。

## 安全

- 不改任何二进制，不碰登录、会员、限速和本地设置。不用 IFEO、hosts、Defender，也不改全机的网络策略。
- 文件改动只在百度网盘自己的安装目录里。注册表只动百度自己的项，外加资源管理器黑名单（`Shell Extensions\Blocked`）以及 Office、浏览器策略里指向百度插件和扩展的几个值。
- 每个改动先写进日志 `journal.jsonl`，记下原来的状态和要写的状态，然后才动手。
- 被换掉的文件和目录先备份到 `backup\`（`-RemoveStagedUpdate` 删掉的更新包除外），删掉的注册表项导出到 `regkeys\`，都在 `%ProgramData%\AppDebloat\baidunetdisk`。
- 防火墙规则都放在 `AppDebloat-baidunetdisk` 组里。撤销时按组、程序和方向匹配，只删自己加的。

## 版本支持

只支持 8.2.1.102。脚本从 `BaiduNetdisk.exe` 的文件版本和卸载信息读版本号，不在支持列表里就拒绝 `-Apply`、`-Install`、`-Auto`，退出码 `2`，什么都不改。演练照样能跑。

只认一个版本，是因为文件名、组件和回装方式都会随版本变。拿 8.2.1.102 的分析去改别的版本，可能拦错东西，也可能拦不住。`-Force` 能强行执行，后果自负。

其他版本可以用 [app-debloat](https://github.com/DaizeDong/app-debloat) 重新分析，生成对应版本的脚本。

安装目录通过卸载信息、环境变量 `BAIDUNETDISK_ROOT` 或默认位置 `%APPDATA%\baidu\BaiduNetdisk` 查找。找到的目录里必须有 `BaiduNetdisk.exe` 和 `module\BrowserEngine\BaiduNetdiskUnite.exe`，`BaiduNetdisk.exe` 的签名必须属于北京度友科技（Beijing Duyou Science and Technology）。

## 常见问题

**百度网盘怎么关闭自动更新？设置里关了为什么还会升级？** 设置里的「开启自动升级」对应 `AutoUpdate\config.ini` 里的 `IsAutoUpdate`，更新程序会自己把 0 改回 1，然后照样下载强制更新包。脚本让更新程序本身起不来：`Autoupdate.exe` 和 `kernelUpdate.exe` 换成同名空目录，`AutoUpdate` 目录拒绝写入，它也没法把自己写回来。

**能去掉百度网盘里的广告吗？** 主窗口里的广告、会员推销和 AI 入口去不掉，它们是服务器下发、界面直接画出来的，本地没有开关。脚本能拿掉的是窗口外面的部分：Office 插件、推送通知（`-MuteNotifyService`）、右键菜单、打开方式、浏览器扩展。设置里那张表能再关掉一部分推荐和弹窗。

**会影响上传下载吗？** 传输内核原样保留，默认动作不碰任何传输连接。只有 `-BlockInboundP2P` 可能影响 P2P 加速的下载，这一点没测过，所以默认不开。

**`HelpUtility.exe` 是什么，能删吗？** 客户端启动约 10 分钟后会连着启动它好几次，上传旧的客户端日志和登录日志，登录令牌（BDUSS、stoken）明文写在它的命令行里。脚本默认给它加出站防火墙规则，程序还在，只是连不上网。它可能也负责帮助和反馈页的上传，那部分会一起失效。

**51413 端口是什么？** 传输内核里带了一个完整的 BT 引擎（Transmission 4.0.6），每次启动都在所有网卡上监听 TCP/UDP 51413，跑 DHT 和本地节点发现，还会向路由器申请端口映射。厂商装的 4 条防火墙规则对所有端口、所有网络类型放行入站，公用网络也不例外。`-BlockInboundP2P` 用阻止规则把它们盖过去。

**资源管理器里别的同步软件的状态图标不见了？** 资源管理器只加载按名字排序的前 15 个图标覆盖，扣掉系统自用的，留给软件的大约 11 个。百度网盘注册了 7 个，名字前加了 6 个空格，排序时挤到前面，排在后面的同步软件（比如 Google Drive）的图标就显示不出来了。不用百度同步盘的话加 `-RemoveOverlays`，然后重启资源管理器。

**右键菜单里的 baidunetdisk 怎么去掉？** 加 `-HideContextMenu`。经典菜单里的那一项会隐藏；Windows 11 顶层菜单里的那一项能不能一起藏掉还没测过。

**我的百度网盘已经升级到新版本了怎么办？** 脚本会拒绝执行，什么都不改。用 [app-debloat](https://github.com/DaizeDong/app-debloat) 对你的版本重新分析，生成对应的脚本。

**会不会把网盘弄坏？** 脚本不改二进制，每一步都先记日志，被换掉的文件先备份（`-RemoveStagedUpdate` 删掉的更新包除外），`-Uninstall` 能还原。先不加参数演练一次，看清会改什么再执行。

搜索关键词：百度网盘 关闭自动更新、百度网盘 强制更新、百度网盘 去广告、百度网盘 精简、百度网盘 纯净、百度网盘 右键菜单、百度网盘 51413 端口、HelpUtility.exe、kernelUpdate.exe、AutoUpdate.exe、YunOfficeAddin、BaiduNetdiskUtility、Baidu Netdisk debloat、disable auto update、block telemetry。

## 相关仓库

- [sogou-ime-debloat](https://github.com/SanJerry007/sogou-ime-debloat)：搜狗输入法的同类脚本，做法相同。
- [app-debloat](https://github.com/DaizeDong/app-debloat)：生成本脚本的 Claude Code skill，可以对其他版本或其他软件重新做一遍分析。

## 说明

- 本项目与百度、北京度友科技无关。脚本只修改你自己电脑上已安装的软件。
- 不绕过会员、限速或任何付费功能。
- 仅供在自己的电脑上个人使用。
- 风险自负。

---

## English Summary

A PowerShell script for Baidu Netdisk (百度网盘) 8.2.1.102 on Windows that stops the client updater and the transfer-kernel updater, blocks the crash-dump and log uploads, and removes the Office add-in that registers itself into Word, Excel and PowerPoint, without patching any binary. The updaters (`Autoupdate.exe`, `kernelUpdate.exe`) are replaced with same-named empty directories and the `AutoUpdate` folder gets a write-deny entry for the current user, because the in-app auto-update switch is rewritten by the updater itself. `crashpad_handler.exe` and the log uploader `HelpUtility.exe` (which carries the session token on its command line) get outbound firewall rules. Fourteen switches cover the rest: `-RemoveStagedUpdate`, `-RemoveSpeechSdk`, `-RemoveAiEngine`, `-DisableUtilityService`, `-BlockBugReport`, `-BlockDiagnoseUpdater`, `-BlockOfficeAddinPolicy`, `-RemoveOverlays` (seven sync-space overlay handlers that push other sync clients' icons out of Explorer's overlay limit), `-BlockInboundP2P` (the built-in BitTorrent engine on TCP/UDP 51413, exposed by the vendor's allow-all inbound rules), `-MuteNotifyService`, `-RemoveAiMsix`, `-HideContextMenu`, `-CleanOpenWith` and `-BlockBrowserExtension`.

Run `BaiduNetdiskDebloat.ps1` with no switch for a dry run, then `-Apply` from an elevated PowerShell. `-Install` applies, copies the script and your options to the admin-only `%ProgramData%\AppDebloat\baidunetdisk` and registers the task `BaiduNetdiskDebloat-Watch`, which re-asserts the blocking actions at every logon and every 6 hours; check `history.log` there for the result. Every change is journaled before it is made and originals are backed up (except the staged update package that `-RemoveStagedUpdate` deletes), so `-Uninstall` replays the journal in reverse and restores only what still looks the way the script left it.

What stays: sign-in, the file list, upload, download and sync, the browser hand-off service on 127.0.0.1:10000, the viewers and the transfer kernel (with its BitTorrent engine). What it cannot remove: ads, upsell and AI entry points inside the main window are drawn by the Electron UI from `pan.baidu.com` data with no local switch, and several settings live in encrypted files or on the account, so you change them in the client (see the settings table above). Only version 8.2.1.102 is supported; `-Apply`, `-Install` and `-Auto` refuse any other version (exit code 2). For another version, regenerate the script with [app-debloat](https://github.com/DaizeDong/app-debloat). Freezing the client also means no security updates for its embedded Chromium and Transmission. Not affiliated with Baidu or Beijing Duyou Science and Technology.
