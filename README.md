# Useful PowerShell Scripts Collection

A comprehensive collection of PowerShell scripts for system administration, network monitoring, educational quizzes, and automation tasks.

## 📋 Table of Contents

- [Scripts Overview](#-scripts-overview)
- [Requirements](#-requirements)
- [Installation](#-installation)
- [Script Details](#-script-details)
- [Usage Examples](#-usage-examples)
- [Contributing](#-contributing)
- [License](#-license)

## 🚀 Scripts Overview

| Script | Description | Category |
|--------|-------------|----------|
| **Dynamic_ping.ps1** | Advanced network monitoring GUI with real-time charts | Network |
| **Automated_IT_Maintenance_Tool.ps1** | Comprehensive system maintenance and optimization | System Admin |
| **Driver_handler.ps1** | Driver management and installation utility | System Admin |
| **APA_Style_installer.ps1** | APA citation style formatter and installer | Productivity |
| **Start_Quiz.ps1** | Interactive quiz system for CCNA exam preparation | Education |

## 💻 Requirements

- **PowerShell 5.1** or higher (Windows PowerShell)
- **PowerShell 7.x** (PowerShell Core) for cross-platform scripts
- **Windows 10/11** or **Windows Server 2016+**
- **Administrator privileges** for system maintenance scripts
- **.NET Framework 4.5+** for GUI applications

### Additional Requirements by Script:
- **Dynamic_ping.ps1**: System.Windows.Forms.DataVisualization
- **Automated_IT_Maintenance_Tool.ps1**: Administrator rights
- **Driver_handler.ps1**: Windows Driver Kit (optional)

## 📦 Installation

1. **Clone the repository:**
```powershell
git clone https://github.com/merknu/Useful-PowerShell-Scripts.git
cd Useful-PowerShell-Scripts
```

2. **Set execution policy (if needed):**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

3. **Verify PowerShell version:**
```powershell
$PSVersionTable.PSVersion
```

## 📝 Script Details

### 🌐 Dynamic_ping.ps1
**Advanced Network Monitoring Tool**

Features:
- Real-time network latency monitoring
- Visual chart display with response times
- Failure detection and packet loss calculation
- Data usage tracking with configurable limits
- Customizable ping intervals
- Support for both IP addresses and domain names

**Usage:**
```powershell
.\Dynamic_ping.ps1
```

**GUI Features:**
- Enter IP or website to monitor
- Set ping interval (5-3600 seconds)
- Configure data usage limit
- View real-time statistics:
  - Response time graph
  - Failure counter
  - Packet loss percentage
  - Data usage tracking

### 🔧 Automated_IT_Maintenance_Tool.ps1
**Comprehensive System Maintenance Suite**

Features:
- Disk cleanup automation
- System information gathering
- Windows Update management
- Registry optimization
- Temporary file cleanup
- System health checks
- Performance optimization

**Usage:**
```powershell
# Run with administrator privileges
.\Automated_IT_Maintenance_Tool.ps1
```

**Available Functions:**
- Disk Cleanup
- System Information Report
- Driver Updates
- Windows Update Check
- Registry Cleanup
- Performance Optimization
- Security Scan

### 🚗 Driver_handler.ps1
**Driver Management Utility**

Features:
- List all installed drivers
- Check for driver updates
- Backup current drivers
- Install/Uninstall drivers
- Driver signature verification
- Compatibility checking

**Usage:**
```powershell
# Run with administrator privileges
.\Driver_handler.ps1
```

### 📚 APA_Style_installer.ps1
**APA Citation Style Formatter**

Features:
- Format citations in APA style
- Install APA templates for Word
- Bibliography management
- In-text citation generator

**Usage:**
```powershell
.\APA_Style_installer.ps1
```

### 🎓 Quizer/Start_Quiz.ps1
**CCNA Exam Quiz System**

Features:
- Interactive quiz interface
- Multiple question banks (CCNA 1 & 2)
- Score tracking
- Review incorrect answers
- Progress monitoring
- CSV-based question storage

**Usage:**
```powershell
cd Quizer
.\Start_Quiz.ps1
```

**Question Banks:**
- CCNA_1_Exam_Quiz.csv
- CCNA_2_Exam_Quiz.csv
- Custom Questions.csv

## 📊 Usage Examples

### Example 1: Monitor Network Connection
```powershell
# Start network monitoring
.\Dynamic_ping.ps1

# In the GUI:
# 1. Enter "google.com" or "8.8.8.8"
# 2. Set interval to 30 seconds
# 3. Set data limit to 1000 KB
# 4. Click "Start Monitoring"
```

### Example 2: System Maintenance
```powershell
# Run full system maintenance
.\Automated_IT_Maintenance_Tool.ps1

# Select tasks:
# - Disk Cleanup
# - System Info
# - Update Drivers
# - Optimize Performance
```

### Example 3: Study for CCNA
```powershell
# Start CCNA quiz
cd Quizer
.\Start_Quiz.ps1

# Select quiz:
# - CCNA 1 Exam
# - CCNA 2 Exam
# - Mixed Questions
```

## 🛡️ Security Considerations

- Always run scripts from trusted sources
- Review script contents before execution
- Use appropriate execution policies
- Run system maintenance scripts with admin privileges only when necessary
- Keep scripts updated for security patches

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/NewScript`
3. Commit your changes: `git commit -m 'Add NewScript.ps1'`
4. Push to the branch: `git push origin feature/NewScript`
5. Submit a Pull Request

### Contribution Ideas:
- Additional system maintenance scripts
- Network diagnostic tools
- Security automation scripts
- Performance monitoring tools
- Documentation improvements

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 📞 Support

For issues, questions, or contributions:
- Open an issue on GitHub
- Submit a pull request
- Contact: merknu@github.com

## 🙏 Acknowledgments

- PowerShell community for inspiration
- Contributors and users of these scripts
- Microsoft for PowerShell development

---

**Note:** Always test scripts in a safe environment before using in production. Some scripts require administrator privileges and may modify system settings.