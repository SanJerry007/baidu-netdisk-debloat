# 百度网盘 8.2.1.102 的启动链和回装路径

以下结论来自对 8.2.1.102 版的静态分析（导入表、字符串、`app.asar` 和 `core.asar` 里的脚本、DLL 自带的注册脚本），以及在一台装好的机器上做的观察：进程启动记录、模块列表、网络连接、文件变化和客户端自己的日志。路径和行为只对这个版本成立。

## 先说结论

客户端没有真正管用的本地开关。设置里的「自动升级」会被更新程序自己改回去；广告和 AI 入口由服务器下发、界面直接画；用户设置是加密存储的。

会把东西装回来的主要是两个更新程序。安装目录在 `%APPDATA%` 下面，用户自己就能写，所以它们都以普通用户身份运行，不弹 UAC：
- `AutoUpdate\Autoupdate.exe`：客户端启动约 30 秒后运行，先更新自己，再下载强制更新包，下次启动客户端就装上。
- `kernelUpdate.exe`：传输内核启动时拉起，常驻，负责更新 `kernel.dll`。

脚本因此不去找开关，直接让这两个程序起不来，再让 `AutoUpdate` 目录写不进去。这样用不着改任何二进制。

## 启动链

```
BaiduNetdisk.exe（主程序）
├─ YunDetectService.exe            网页下载接力，监听 127.0.0.1:10000，客户端退出后仍在        [保留]
├─ module\BrowserEngine\BaiduNetdiskUnite.exe（Electron 主窗口，app.asar / core.asar）        [保留]
│  │  广告位、会员推销、AI 侧栏都在这里画，内容来自 pan.baidu.com 的接口
│  ├─ baidunetdiskhost.exe --plugin_id=3      视频播放 vastplayer.dll                          [保留]
│  ├─ baidunetdiskhost.exe --plugin_id=1000   传输内核 kernel.dll、netbase.dll                 [保留]
│  │  ├─ 内置 BT 引擎 Transmission 4.0.6：TCP/UDP 51413、DHT、本地节点发现、UPnP/NAT-PMP     [-BlockInboundP2P]
│  │  ├─ P2P 上传共享（界面里叫「下载加速」）：UDP 8568，开关默认关
│  │  ├─ 上传测速：一次约 115 MB，配置间隔 12 小时                                           [拦不了]
│  │  └─ kernelUpdate.exe（getlatestversion，常驻）                                          [默认拦]
│  ├─ baidunetdiskhost.exe --plugin_id=1004   语音 module\SpeechSDK                            [-RemoveSpeechSdk]
│  └─ 每个插件宿主 → crashpad_handler.exe   上传崩溃转储，带账号标识                         [默认断网]
├─ 启动约 30 秒后：AutoUpdate\Autoupdate.exe -client_info ...                                   [默认拦]
│     先把自己换成新版（旧文件留成 Autoupdate.exe.old），再把强制完整包下到 AutoUpdate\Download\MainApp
├─ 启动约 10 分钟后：HelpUtility.exe -cmd report_log（33 秒里启动 8 次）                        [默认断网]
│     上传客户端日志和登录日志，命令行里带着 -bduss、-netdiskstoken
├─ 出异常时：BugReport.exe                                                                     [-BlockBugReport]
├─ 启动时：一个几秒就退出的 explorer.exe 子进程；商店应用 BaiduNetdiskService 随后运行          [未确认]
└─ 按路径登记：YunOfficeAddin.dll / YunOfficeAddin64.dll                                      [默认拦]
      自带注册脚本，写 HKCU\Software\Microsoft\Office\{Word,Excel,PowerPoint}\Addins，LoadBehavior 3，也写 WPS

explorer.exe（以及其他程序的打开、保存对话框）
└─ YunShellExtV164.dll：右键菜单「baidunetdisk」、同步盘的 7 个图标覆盖、同步盘命名空间

BaiduNetdiskUtility 服务（SYSTEM，按需启动）                                                   [-DisableUtilityService]
└─ YunUtilityService.exe：程序文件在用户可写的安装目录里，普通用户就能启动这个服务

autoDiagnose.dll → autoDiagnoseUpdate.exe、ServiceAssistans.exe                                [-BlockDiagnoseUpdater]
   只在字符串里出现，长时间监控里一次都没启动过
```

加上 `-RemoveSpeechSdk` 处理后，常驻进程从 17 个降到 14 个：少了 `kernelUpdate.exe`、语音插件宿主和它带的那个 `crashpad_handler.exe`。这 14 个进程在开着主窗口空闲时，工作集合计约 1.1 到 1.2 GB。只用默认动作时只少 `kernelUpdate.exe`。

## 真正的闸门

**自动升级开关不是闸门。** 设置里的「开启自动升级」对应 `AutoUpdate\config.ini` 的 `[AutoUpdate] IsAutoUpdate`。把它设成 0 以后，更新程序一分钟内就改回 1，照样下载。服务器在内核配置里也下发了 `open_auto_update=1`。

**更新不需要 UAC。** 安装目录是 `%APPDATA%\baidu\BaiduNetdisk`，在用户自己的配置文件里。更新程序以普通用户身份运行，能直接改写自己，再下载约 805 MB 的强制完整包（服务器把它标成 `level="force"`），下次启动客户端时装上。服务器目前还接受 8.2.1.102。

**广告和 AI 入口没有本地闸门。** 它们由 Electron 渲染器按 `pan.baidu.com` 接口返回的内容画出来。`app.asar` 里除了 `pan.baidu.com` 和百度的 CDN（`bdstatic`、`bcebos`），没有单独的广告域名，所以在 hosts 或 DNS 里屏蔽会连核心功能一起断。用户设置存在 `module\BrowserEngine\users\<编号>\PersonalSetting.xml`，是加密的；注册表里只有安装目录、版本和渠道。内核的网络参数（`upnp_on`、`enable_dht`、上传测速等）由服务器下发，加密缓存在 `%APPDATA%\BaiduYunKernel\Config\kernelb.cfg`。

**Office 插件谁也不依赖。** 安装目录下 320 个 PE 文件，没有一个在导入表或延迟导入表里引用 `YunOfficeAddin(64).dll`，只有 `BaiduNetdisk.exe` 把这两个文件名当字符串存着，用来按路径注册。所以把它们换成同名空目录后，客户端照常启动，Office 也照常运行，只是这个插件注册不进去了。

**Office 策略是第二道防线。** 在 `HKCU\Software\Policies\Microsoft\Office\16.0\<程序>\Resiliency\AddinList` 下，以插件名为值名、数据写 `"0"`，按 Office 的文档，这个插件会一直处于禁用状态，就算它把自己注册进去也一样。这是文档描述的行为，没有实测，所以做成选项 `-BlockOfficeAddinPolicy`。

## 回装路径和对应的堵法

| 回装路径 | 谁、什么时候 | 堵法 | 验证情况 |
|---|---|---|---|
| 客户端更新程序：先更新自己，再下载强制完整包 | 普通用户身份，每次启动约 30 秒后 | `Autoupdate.exe` 换成同名空目录；`AutoUpdate` 目录对当前用户拒绝写入（写数据、追加、写扩展属性、写属性，子目录和文件继承）；`-RemoveStagedUpdate` 删掉已下载的包 | 重启后、连续运行两天半、21 分钟监控里都没再启动，`AutoUpdate` 下没有任何文件变化 |
| 旧更新程序 `Autoupdate.exe.old` 被改回原名 | 厂商更新时的改名 | 同一条拒绝写入 | 以普通用户身份改名、复制、新建都被拒 |
| 传输内核更新程序 | 普通用户身份，内核插件启动时 | `kernelUpdate.exe` 换成同名空目录 | 处理后每次启动都没再出现 |
| BaiduNetdiskUtility 服务（推测会安装或修复组件） | SYSTEM，按需启动 | `-DisableUtilityService`，只禁用，不删除 | 禁用状态在重启后保持 |
| 普通进程删掉占位，或者去掉拒绝写入项 | 任何以你身份运行的程序都做得到：继承来的完全控制里包含删除和修改权限 | 计划任务在登录时和每 6 小时重做 | 拒绝写入项被去掉后，下一次运行把它和占位一起补了回来 |
| Office 插件自注册 | `BaiduNetdisk.exe` 按路径注册 | 两个 DLL 换成同名空目录；`-BlockOfficeAddinPolicy` 再加一道策略 | 处理后客户端正常启动 |
| 浏览器扩展侧载 | 用户在客户端弹窗里点了同意 | `-BlockBrowserExtension` | 没有触发过这个弹窗 |
| 通知服务从附带的 `BaiduNetdiskService.Msixbundle` 安装 | 没观察到 | `-MuteNotifyService` 占住安装包 | 假设 |
| AI 包 `sysres\BaiduNetdiskAI.msix` 被注册 | 没观察到 | `-RemoveAiMsix` | 假设 |
| 图标覆盖、右键菜单、打开方式被重新注册 | 修复、重装或更新时（没观察到） | 对应的选项，计划任务重做 | 没触发过 |
| 同步盘助手（商店应用 DesktopSyncClient）被重装 | 假设 | 不处理，保留 | 未确认 |
| 你自己重新运行官方安装包 | 你 | 版本没变：计划任务在下次登录或 6 小时内重做。版本变了：脚本停手，记一行日志 | 未触发 |

## 跨软件的劫持

除了自己的进程，百度网盘还往系统和别的软件里伸手。下面是逐项排查找到的东西和脚本的处理方式。默认只拦 Office 插件；其余的要么有人会用到，要么还没测完，做成选项。

**资源管理器图标覆盖。** 7 个覆盖标识 `      .WorkspaceExt0` 到 `.WorkspaceExt6`，名字前有 6 个空格，64 位和 32 位注册表视图里都有。资源管理器按名字排序，只加载前 15 个，扣掉系统自用的，留给软件的大约 11 个。名字前空格越多排得越靠前，这 7 个因此排到了大多数同步软件前面，排在后面的（比如 Google Drive）状态图标就显示不出来。它们服务的是同步盘，所以做成选项。`-RemoveOverlays` 删标识时要求名字和默认值同时匹配（默认值必须是百度的 7 个 CLSID 之一），再把这些 CLSID 加进资源管理器的黑名单，重新注册也加载不了。

**右键菜单。** 所有文件的右键菜单里都有「baidunetdisk」（`HKLM\SOFTWARE\Classes\*\shell\YunShellExplorerCommand`，处理程序 `{499C8928-90D7-411D-8DBD-E35B92529123}`），文件夹那一项厂商自己已经加了 `LegacyDisable`。Windows 11 顶层菜单里的那一项来自同步盘助手 AppX，用的是同一个 CLSID。`-HideContextMenu` 给文件那一项也加上 `LegacyDisable`，再把 CLSID 加进黑名单；黑名单能不能藏住 Windows 11 那一项没测过。右键上传是正经功能，所以做成选项。

**「此电脑」里的网盘入口。** 命名空间文件夹 `{679F137C-3162-45da-BE3C-2F9C3D093F64}` 已经注册，但没有挂到「此电脑」下面。设置里的「在我的电脑中显示网盘」会把它挂上去。脚本不处理，在设置里关着就行。

**Office 和 WPS 插件。** 见上一节，默认拦。

**打开方式和默认应用。** 26 种音视频和 PDF 扩展名下有 `OpenWithList\BaiduNetdisk.open`，「默认应用」里登记了网盘看图（20 种图片格式）。分析时它没有抢走任何默认程序，但更新后的引导页会默认勾上「设置「智能看图」为系统图片默认查看方式」。`-CleanOpenWith` 删掉这些项。每个 `BaiduNetdisk.open` 子项都只属于百度，删掉不会动到别的软件。

**浏览器扩展。** 安装目录里有 `netdisk-chrome-ext.crx`，ID `lfkkoabmfeaefiddnpbjeiofkmkpogim`，Manifest V3，权限有 cookies、webRequest、downloads、tabs，内容脚本覆盖所有网站。`browserengine.dll` 和 `cloudpic.dll` 里有 Chrome、Edge 侧载用的注册表路径，`core.asar` 里有安装弹窗。分析时它没有装进任何浏览器。`-BlockBrowserExtension` 把这个 ID 写进 Chrome、Edge 的 `ExtensionInstallBlocklist` 机器策略，用最小的空闲编号，不覆盖已有的值；侧载注册表项在每次运行时发现就删（装成自动模式后是登录时和每 6 小时）。代价是策略存在期间浏览器会显示「由你的组织管理」，所以做成选项。

**通知通道。** 第二个商店应用 `3AD8A913.49937325EFC59`（BaiduNetdiskService）带应用服务、完全信任的 `NotifyExtension.exe`、系统通知和一条出站放行规则，客户端启动时它也会跑起来。它具体推什么没观察到，传输完成的通知可能也走它，所以做成选项 `-MuteNotifyService`。

**AI 包。** `sysres\BaiduNetdiskAI.msix` 是看图程序的稀疏包，声明了 `systemAIModels` 能力（调用 Windows 本地 AI 模型）。分析时它没装，谁会装它也不知道。`-RemoveAiMsix` 把它换成同名空目录。

**入站端口。** 厂商装了 4 条入站放行规则，分别给 `BaiduNetdisk.exe`、`BaiduNetdiskRender.exe`、`baidunetdiskhost.exe`、`BaiduNetdiskUnite.exe`：任意协议、任意端口、所有网络类型都放行，公用网络也一样。传输内核里的 BT 引擎每次启动都在所有网卡上监听 51413。`-BlockInboundP2P` 给这 4 个程序加入站阻止规则，阻止规则优先于放行规则；出站传输和本机回环（10000、8475）不受影响。还没做过开着它的 BT 或磁力下载测试，所以默认不开。

**电源计划。** `BaiduNetdisk.exe` 导入了 `PowerWriteACValueIndex`、`PowerWriteDCValueIndex`，带着处理器加速模式的 GUID，设置里的「系统兼容模式」会改当前电源计划。没有能用的拦法，设置里别勾。

**剪贴板。** `core.asar` 在启动时读剪贴板，取出 `semclickid=`（搜索广告的点击编号）和渠道参数上报，再把这一段去掉写回剪贴板。「自动识别剪切板里的链接」默认打开，用于识别链接和 AI 收藏。这些都在核心进程里，能做的只有在设置里关掉后者。

## 证据等级

脚本里每个动作都标了证据等级：

| 等级 | 含义 |
|---|---|
| 实测（observed） | 在进程记录、模块列表、网络连接或文件变化里直接看到 |
| 有据（anchored） | 在文件里的确切位置找到（导入表、字符串、注册脚本、界面代码），但没看到它实际发生 |
| 推断（inferred） | 从上面两类证据推出来 |
| 假设（hypothesis） | 有可能，还缺证据 |

默认动作都是实测或有据的。推断和假设级的动作不会自己执行，必须加上对应的选项；演练时会把它们连同证据一起列出来。目前属于这一类的有 `-DisableUtilityService`、`-BlockBugReport`、`-BlockOfficeAddinPolicy`（推断），以及 `-BlockDiagnoseUpdater`、`-RemoveAiMsix` 和 `-MuteNotifyService` 里占住安装包的那一步（假设）。

## 做过的耐久性测试

下面这些都在 8.2.1.102 上跑过，结果符合预期：

- **重启之后。** 拒绝写入项、全部占位、防火墙规则、禁用的服务都还在，版本没变。登录时计划任务跑了一次，0 处改动、0 个失败。和处理前的快照对比，只有计划内的变化。
- **连续运行约两天半。** 客户端一直开着，`AutoUpdate` 下没有任何文件变化，没有更新程序运行过，期间计划任务的自动运行全部是 0 处改动。
- **21 分钟监控。** 开着进程启动跟踪，从启动客户端开始记录：`Autoupdate.exe`、`kernelUpdate.exe`、`autoDiagnoseUpdate.exe`、`BugReport.exe`、`ServiceAssistans.exe`、`YunUtilityService.exe` 一个都没启动；被拦的目录下没有文件事件；用户目录和临时目录里没有写出 20 MB 以上的新文件；`crashpad_handler.exe` 没有任何连接；语音插件宿主没再启动，传输内核照常加载；主窗口显示登录状态和文件列表。
- **以普通用户身份硬闯。** 在 `AutoUpdate` 里新建文件或目录、把 `Autoupdate.exe.old` 改名或复制回 `Autoupdate.exe`、以写方式打开 `config.ini`，全部被拒。占位目录被删掉以后也一样。
- **去掉拒绝写入项。** 以普通用户身份把它删掉，下一次计划任务运行把它和占位一起补了回来。
- **新增默认动作上线后。** 两个 Office 插件 DLL 变成占位目录、`HelpUtility.exe` 的出站规则生效以后，客户端正常启动并显示文件列表，更新程序没有出现，计划任务运行 0 处改动、0 个失败。

## 排查过但没发现问题的地方

- 没有 AppInit、AppCertDlls、IFEO、Winsock LSP、KnownDLLs 之类的注入，没有驱动，系统目录里没有它的文件。
- 没有计划任务和 WMI 订阅。
- 没有虚拟打印机、打印端口监视器、输入法组件、Outlook 插件。
- 分析时没有改写任何默认程序（http、https、magnet、.torrent、.pdf、图片、视频、压缩包）。
- 除了资源管理器和用过文件对话框的程序，没有别的进程加载百度的 DLL。

## 已知的限制

- 只支持 8.2.1.102。客户端和它自带的 Chromium、BT 库都停在这个版本，收不到安全更新。
- 服务器已经把新版本标成强制更新。哪天服务器不再接受 8.2.1.102，先 `-Uninstall` 让它升级，再用 app-debloat 对新版本重新分析。
- 主窗口里的广告、推销和 AI 入口去不掉。
- 上传测速、加密遥测、启动时读剪贴板、BT 引擎的 UPnP 请求都在核心进程里，拦不了。
- `Documents\BaiduNetdiskTmp` 里的临时文件会一直堆积，脚本不清理。
- 测试确认了登录、文件列表和传输内核加载正常，但没有跑过真实的上传和下载。`-DisableUtilityService` 关掉服务以后上传是否完全正常，也没测过。
- 拒绝写入只管写入。以你身份运行的程序仍能删掉占位、去掉拒绝写入项，计划任务最多 6 小时后补回；只要拒绝写入项还在，更新程序就写不回来。
- 被删掉的占位由计划任务补回，这是引擎的设计行为，但这一版脚本上还没有重做模拟删除的测试。
- 任务计划程序里的「上次运行结果」不反映失败（任务通过 `conhost.exe --headless` 启动 PowerShell），以 `history.log` 为准。
- `-RemoveSpeechSdk`、`-RemoveAiEngine` 执行前会检查有没有进程加载着这些模块。碰到查不了的进程，这一步会推迟到下一次运行。
- 如果更新程序在你跑脚本之前已经运行过，安装目录里可能留着它准备替换的文件（比如 `*.asar.new`、`*.dll.old`）。脚本不动它们；测试期间它们一直没有被换进去。
- `HelpUtility.exe` 可能也负责帮助和反馈页的上传，断网以后那部分也会失败（没有见过）。
