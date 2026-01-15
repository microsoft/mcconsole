# McConsole - SSH & Multi-SUT Command Management Tool

McConsole is a Windows-based SSH/console management tool designed to organize, manage, and execute commands across multiple Systems Under Test (SUTs) efficiently. It combines an intuitive GUI with powerful multi-SUT capabilities and integrated terminal support.

## 🚀 Quick Start

### Key Features

- **🗂️ Node & Folder Management** - Organize SUTs into custom folders with favorites support
- **🔐 SSH & Connection Management** - Jumpbox support, MobaXterm integration, multi-console handling
- **📋 Command Helper Windows** - Execute single or multi-SUT commands with CSV-based management
- **🎮 Multi-SUT Console Control** - Batch command execution across multiple SUTs simultaneously
- **🔄 Auto-Refresh & Status Monitoring** - Real-time LED indicators for connection health
- **🎨 Customization & Themes** - Multiple UI themes, background music, and personalization options
- **🔧 Developer-Friendly** - Modular architecture with JSON config and CSV command databases

### Installation & Usage

1. **Add a Node**: Click "📁 New" → Create folder → Add node with console IPs
2. **Configure Commands**: Edit CSV files in `helper/` folder (e.g., `CAAA_CONTROLLER_commands.csv`)
3. **Connect**: Click "SSH" or "Control" button to launch console
4. **Execute Commands**: Double-click commands in helper window or right-click to edit
5. **Multi-SUT**: Hold Ctrl+Click to select multiple nodes, then open Multi-SUT Console

### System Requirements

- Windows 10/11 (64-bit)
- Python 3.8+
- MobaXterm (included) or SecureCRT
- TCP/IP network connectivity

### Security & Privacy

✅ No remote data collection  
✅ Local credential storage (MobaXterm-managed)  
✅ Jumpbox support for secure network access  
✅ Password masking in logs  

## Contributing

This project welcomes contributions and suggestions. Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit [Contributor License Agreements](https://cla.opensource.microsoft.com).

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft
trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

## Support & Documentation

For detailed information and support, see [SUPPORT.md](SUPPORT.md).