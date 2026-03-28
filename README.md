# Fix Office Apps Defaulting to ODS/ODT Format with Intune Proactive Remediation

> Detects and remediates Office 365 applications (Excel, Word, PowerPoint) defaulting to open document formats instead of native Office formats (XLSX, DOCX, PPTX) using Intune Proactive Remediation.

## Overview

This PowerShell script was extracted from an article on [MSEndpoint.com](https://msendpoint.com) and is part of the professional automation portfolio.

### Quick Start

```powershell
# Clone the repository
git clone https://github.com/Msendpoint/intune-fix-office-default-format-remediation.git
cd intune-fix-office-default-format-remediation

# One-click install & run (checks prerequisites, installs modules, launches script)
.\Install.ps1

# Or run the script directly
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
.\scripts\intune-fix-office-default-format-remediation.ps1
```

## Script Usage

```powershell
.\scripts\intune-fix-office-default-format-remediation.ps1
```

## Tech Stack

PowerShell, Microsoft Intune, Proactive Remediation, Windows Registry

## Requirements

- PowerShell 7.0 or later
- Microsoft Graph CLI (if applicable)
- Exchange Online Management module (if applicable)
- Appropriate administrative permissions

## About the Author

**Souhaiel Morhag** — Microsoft 365 & Intune Automation Specialist

- 🌐 **Blog**: [https://msendpoint.com](https://msendpoint.com) — Enterprise endpoint management insights
- 📚 **Academy**: [MSEndpoint Academy](https://app.msendpoint.com/academy) — Professional training & certifications
- 💼 **LinkedIn**: [Souhaiel Morhag](https://linkedin.com/in/souhaiel-morhag) — Connect for collaboration
- 🐙 **GitHub**: [@Msendpoint](https://github.com/Msendpoint)

## Source Article

This script originated from:
**[Fix Office Apps Defaulting to ODS/ODT Format with Intune Proactive Remediation](https://msendpoint.com/article/intune-detection-and-remediation-of-default-open-document-formats-on-office-365-apps)**

Published on MSEndpoint.com — Enterprise endpoint management blog

## Contributing

1. Fork this repository
2. Create a feature branch (`git checkout -b feature/improvement`)
3. Commit your changes (`git commit -m 'feat: add improvement'`)
4. Push to the branch (`git push origin feature/improvement`)
5. Open a Pull Request

## License

MIT License — see [LICENSE](LICENSE) for details.

## Support

- Report issues: [GitHub Issues](https://github.com/Msendpoint/intune-fix-office-default-format-remediation/issues)
- Questions: Comment on the original article or reach out on LinkedIn

---

*Developed with ❤️ by [MSEndpoint.com](https://msendpoint.com) — Your Microsoft 365 Endpoint Management Partner*