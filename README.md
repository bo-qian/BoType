# BoType

**BoType** 是一款专业的 Microsoft Word 插件 (VSTO Add-in)，专注于提供强大的公式排版、编号和引用管理功能。旨在让科研人员、学生及文字工作者能更高效地在 Word 中撰写包含数学公式的学术论文与报告。

## ✨ 核心特性

- **快速插入公式**：一键插入单行独立公式或行内公式。
- **自动化公式编号**：自动为公式添加编号，并支持多种编号样式（如 \(1)\、\(1.1)\、\(1-1)\ 等）。
- **默认样式管理**：支持一键将当前编号样式设为全局默认。
- **智能交叉引用**：
  - **插入引用**：快速在正文中插入对公式的交叉引用。
  - **更新引用**：当公式顺序发生变化时，一键刷新全文的公式引用，防止错乱。

## 💻 系统与环境要求

- **操作系统**：Windows 10 / 11
- **软件要求**：Microsoft Office Word 2016 及以上版本 (包括 Office 365)
- **运行环境**：.NET Framework 4.7.2 及更高版本

## 🚀 安装步骤

1. 前往本项目的 [Releases](https://github.com/bo-qian/BoType/releases) 页面，或者直接在当前代码仓库下载 \BoType_Release.zip\ 文件。
2. 将 \BoType_Release.zip\ 下载到你的电脑并解压到一个文件夹中。
3. 双击运行解压后的 \setup.exe\ 程序。如果在安装时遇到安全提示（如 Windows 保护你的电脑、智能屏幕拦截等），请点击“**更多信息**”并选择“**仍要运行**”，然后在弹出的安装窗口点击“**安装**”。
4. （如果在安装过程提示缺少环境组件，安装程序会自动帮你下载安装 VSTO 运行环境。）
5. 安装按完成后，重新启动你的 Microsoft Word。
6. 打开任意 Word 文档，即可在顶部菜单栏（功能区）看到新加入的 **BoType** 选项卡。

**遇到问题（如加载项被禁用）？**
如果打开 Word 后没有看到 BoType 选项卡：
- 点击 Word 左上角的 \文件\ -> \选项\ -> \加载项\。
- 在底部的【管理】下拉菜单选择 \COM 加载项\，点击 \转到...\。
- 在弹出的列表中找到 \BoType\，勾选它并点击 \确定\ 即可启用。

## 💡 使用指南

1. **插入公式**：点击菜单栏的 **BoType**，在“插入公式”组选择你需要创建的公式类别。
2. **给公式编号**：写完公式后，在“公式编号”组选择您需要的样式，并点击“给公式编号”。
3. **公式引用**：将光标放在正文中需要引用的位置，点击“插入引用”，从列表中选择对应的公式即可。在大量修改文档后，记得点击“更新引用”以保证所有编号准确无误。

## 📄 许可证 (License)

本项目采用 [MIT License](LICENSE) 许可协议。

## 🔐 安装签名证书（仅当安装提示证书不受信任时需要）

如果用户在运行 `setup.exe` 时收到“证书不受信任”或类似 SecurityException，可按以下方式快速信任自签名证书（适用于内部测试）。

1. 从本 Release 下载并解压所有文件，确保 `BoType_TemporaryKey.pfx` 与 `tools` 目录在同一目录下。
2. 以管理员身份打开 PowerShell（右键 PowerShell -> 以管理员身份运行）。
3. 运行一键脚本（会导入证书并启动安装程序）：

   ```powershell
   pwsh -ExecutionPolicy Bypass -File .\tools\install_and_run_setup.ps1 -PfxPath ".\BoType\BoType_TemporaryKey.pfx" -PfxPassword "123456" -SetupPath ".\BoType\publish\setup.exe"
   ```

4. 脚本执行完毕后，即可正常运行安装程序而不会再出现“不受信任证书”的错误。

如果你只想单独信任公钥（管理员可先信任证书再手动运行安装程序）：

```powershell
pwsh -ExecutionPolicy Bypass -File .\tools\install_public_cert.ps1 -CerPath ".\BoType\BoType_TemporaryKey.cer"
```

注意：以上为自签名证书的临时解决方案，适合内部或测试环境。对外发布请务必使用受信任 CA 签发的代码签名证书。
