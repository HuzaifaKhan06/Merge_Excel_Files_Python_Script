# 📊 Excel Merge & Email Automation Tool

A powerful yet simple Python desktop application that streamlines the process of merging multiple Excel files and sending them via email. Built with a user-friendly GUI, this tool eliminates repetitive manual work and automates data consolidation tasks.

---

## 🎯 Overview

This application provides a complete solution for Excel file management with:

- **Automated Excel Merging** - Combine multiple `.xlsx` or `.xls` files into a single master file
- **Real-time Data Preview** - View merged data with row and column statistics
- **Integrated Email Functionality** - Send merged files directly through Gmail
- **Intuitive GUI Interface** - No coding knowledge required to operate
- **Configuration Management** - Securely save email credentials for repeated use

**Ideal for:** Data analysts, accountants, office administrators, researchers, and anyone handling multiple Excel spreadsheets regularly.

---

## ✨ Key Features

| Feature | Description |
|---------|-------------|
| **Multi-File Merging** | Automatically combines multiple Excel files while preserving data integrity |
| **Data Statistics** | Displays total rows, columns, and sheet counts in real-time |
| **Email Integration** | Send merged files via Gmail with attachment support |
| **Smart Configuration** | Saves email settings securely using JSON configuration |
| **File Format Support** | Compatible with `.xlsx`, `.xls`, and `.xlsm` formats |
| **Error Handling** | Comprehensive validation and user-friendly error messages |
| **Cross-Platform** | Works on Windows, macOS, and Linux |

---

## 🛠️ Technology Stack

### Core Libraries

| Library | Version | Purpose |
|---------|---------|---------|
| **pandas** | Latest | Data manipulation, reading/writing Excel files, and DataFrame operations |
| **openpyxl** | Latest | Excel file engine for `.xlsx` format support |
| **PySimpleGUI** | Latest | Cross-platform GUI framework for desktop interface |
| **yagmail** | Latest | Simplified Gmail SMTP integration for email automation |
| **os / glob** | Built-in | File system operations and pattern matching |
| **json** | Built-in | Configuration file management |

---

## 📁 Project Structure

```
Merge_Excel_Files_Python_Script/
│
├── 📄 merge_excel.py              # Main application executable
├── 📄 excel_merge_app.ipynb       # Jupyter Notebook version (optional)
├── 📄 requirements.txt            # Python dependencies
├── 📄 config.json                 # Email configuration (auto-generated)
├── 📄 README.md                   # This file
└── 📄 LICENSE                     # License information
```

### File Descriptions

- **`merge_excel.py`** - Primary application file containing all GUI logic, data processing, and email functionality
- **`excel_merge_app.ipynb`** - Alternative Jupyter Notebook interface for step-by-step execution
- **`requirements.txt`** - Lists all required Python packages with version specifications
- **`config.json`** - Stores email credentials securely (created automatically on first configuration)

---

## 🚀 Installation

### Prerequisites

- **Python 3.7 or higher** installed on your system
- **pip** package manager
- **Git** (for cloning the repository)

### Step 1: Clone the Repository

```bash
git clone https://github.com/HuzaifaKhan06/Merge_Excel_Files_Python_Script.git
cd Merge_Excel_Files_Python_Script
```

### Step 2: Create Virtual Environment (Recommended)

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### Step 3: Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 4: Verify Installation

```bash
python merge_excel.py --version
```

---

## ▶️ Usage Guide

### Launching the Application

#### Method 1: Python Script

```bash
python merge_excel.py
```

#### Method 2: Jupyter Notebook

1. Launch Jupyter:
   ```bash
   jupyter notebook
   ```
2. Open `excel_merge_app.ipynb`
3. Run cells sequentially

### Application Workflow

1. **Select Files**
   - Click "Browse" to select the folder containing Excel files
   - Preview file list and data structure

2. **Merge Data**
   - Click "Merge Files" button
   - Review row/column counts
   - Preview merged data in the display window

3. **Export Results**
   - Save merged file to desired location
   - File automatically named with timestamp

4. **Send via Email** (Optional)
   - Configure email settings in Settings tab
   - Enter recipient address
   - Click "Send Email" to dispatch merged file

---

## 📧 Email Configuration

### Setting Up Gmail Integration

#### Step 1: Configure Application Settings

1. Navigate to the **Settings** tab in the application
2. Enter your Gmail address
3. Enter your Gmail App Password (not regular password)
4. Click **Save Configuration**

#### Step 2: Generate Gmail App Password

**Important:** App Passwords are required for security. Never use your regular Gmail password.

**Instructions:**

1. Visit [Google Account Security](https://myaccount.google.com/security)
2. Enable **2-Step Verification** if not already enabled
3. Search for **App Passwords** in the search bar
4. Select **Mail** and **Other (Custom name)**
5. Enter app name: "Excel Merge Tool"
6. Click **Generate**
7. Copy the 16-character password
8. Paste into the application settings

#### Step 3: Security Notes

- ✅ Compatible with **personal Gmail accounts only**
- ⚠️ Corporate/institutional Gmail may require additional SMTP configuration
- 🔒 App passwords are stored locally in `config.json`
- 🛡️ Never share your app password publicly
- 🔄 Revoke app passwords if compromised

### Troubleshooting Email Issues

| Issue | Solution |
|-------|----------|
| Authentication failed | Verify app password is correct and 2FA is enabled |
| Cannot send email | Check internet connection and Gmail server status |
| Attachment too large | Gmail limit is 25MB; compress or split files |
| Corporate email blocked | Contact IT department for SMTP configuration |

---

## 🔧 Advanced Configuration

### Custom SMTP Settings (For Corporate Email)

Edit the email configuration in `merge_excel.py`:

```python
# Example for custom SMTP server
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

smtp_server = "smtp.yourcompany.com"
smtp_port = 587
smtp_user = "your.email@company.com"
smtp_password = "your_password"
```

### Customizing Merge Behavior

Modify merge parameters in the code:

```python
# Ignore index column when merging
merged_df = pd.concat(dataframes, ignore_index=True)

# Specify sheet names to merge
df = pd.read_excel(file, sheet_name='Sheet1')

# Handle duplicate columns
merged_df = merged_df.loc[:, ~merged_df.columns.duplicated()]
```

---

## 📋 Requirements

### Python Packages

```txt
pandas>=1.3.0
openpyxl>=3.0.9
PySimpleGUI>=4.60.0
yagmail>=0.15.0
```

### System Requirements

- **OS:** Windows 10/11, macOS 10.14+, or Linux (Ubuntu 20.04+)
- **RAM:** 4GB minimum (8GB recommended for large files)
- **Storage:** 100MB free space
- **Network:** Internet connection for email functionality

---

## 🐛 Troubleshooting

### Common Issues

**Problem:** Application won't launch

- **Solution:** Verify Python version with `python --version` (must be 3.7+)
- **Solution:** Reinstall dependencies: `pip install -r requirements.txt --force-reinstall`

**Problem:** Excel files not merging properly

- **Solution:** Ensure all files have consistent column headers
- **Solution:** Check for corrupted Excel files
- **Solution:** Verify file permissions (read access required)

**Problem:** GUI display issues

- **Solution:** Update PySimpleGUI: `pip install --upgrade PySimpleGUI`
- **Solution:** Try running with administrator/sudo privileges

**Problem:** Memory errors with large files

- **Solution:** Process files in smaller batches
- **Solution:** Increase available system RAM
- **Solution:** Use 64-bit Python installation

---

## 🤝 Contributing

Contributions are welcome! Please follow these guidelines:

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/AmazingFeature`)
3. **Commit** your changes (`git commit -m 'Add some AmazingFeature'`)
4. **Push** to the branch (`git push origin feature/AmazingFeature`)
5. **Open** a Pull Request

### Development Guidelines

- Follow PEP 8 style guidelines
- Add unit tests for new features
- Update documentation for API changes
- Include comments for complex logic

---

## 👨‍💻 Author

**Muhammad Huzaifa Khan**

- **Role:** Data Analyst & Web Developer | Python Automation

---

## 🙏 Acknowledgments

- PySimpleGUI team for the excellent GUI framework
- Pandas community for robust data manipulation tools
- yagmail developers for simplified email integration
- All contributors and users who provide feedback

---

## 📞 Support

If you encounter any issues or have questions:
- 📧 **Email:** hk9349881@gmail.com
---

## ⭐ Show Your Support

If this project helped you, please consider:

- ⭐ **Starring** the repository
- 🍴 **Forking** for your own projects
- 📢 **Sharing** with colleagues
- 💬 **Providing feedback** and suggestions

---

## 📊 Project Statistics

![GitHub stars](https://img.shields.io/github/stars/HuzaifaKhan06/Merge_Excel_Files_Python_Script?style=social)
![GitHub forks](https://img.shields.io/github/forks/HuzaifaKhan06/Merge_Excel_Files_Python_Script?style=social)
![GitHub issues](https://img.shields.io/github/issues/HuzaifaKhan06/Merge_Excel_Files_Python_Script)

---

## 🔮 Future Enhancements

- [ ] Support for CSV and JSON file formats
- [ ] Scheduled automatic merging
- [ ] Cloud storage integration (Google Drive, Dropbox)
- [ ] Advanced filtering and transformation options
- [ ] Multi-language support
- [ ] Batch email sending with templates
- [ ] Data visualization dashboard

---

<div align="center">

**"Automation doesn't replace people it empowers them."** 🤖

Made with ❤️ by Muhammad Huzaifa Khan

</div>
