
# Office Suite 安装指南

---

## 📋 系统要求
- Python 版本：3.8 或更高
- 操作系统：Windows 10+/macOS 10.15+/Linux (Ubuntu 20.04+/CentOS 7+/OpenCloudOS 9+)
- 内存：至少 2GB RAM
- 磁盘空间：至少 500MB 可用空间

---

## 🔧 安装步骤
### 1. 克隆项目
```bash
git clone https://github.com/shynloc/ACKS-Office-Suite-Simplify.git
cd ACKS-Office-Suite-Simplify
```

### 2. 安装Python依赖
```bash
# 核心依赖（必须安装）
pip install python-docx openpyxl reportlab PyPDF2 python-pptx pandas xlrd Pillow

# 完整功能依赖（可选，推荐安装）
pip install Jinja2 click rich python-magic chardet tqdm
```

### 3. （可选）安装系统工具
用于高质量的格式转换（推荐安装）：

**Linux (Ubuntu/Debian/OpenCloudOS)**：
```bash
# Ubuntu/Debian
sudo apt update
sudo apt install -y libreoffice imagemagick ghostscript fonts-noto-cjk

# OpenCloudOS/CentOS
sudo dnf install -y libreoffice imagemagick ghostscript google-noto-cjk-fonts
```

**macOS**：
```bash
brew install libreoffice imagemagick ghostscript
```

**Windows**：
- 下载安装 [LibreOffice](https://www.libreoffice.org/)
- 下载安装 [ImageMagick](https://imagemagick.org/)

---

## ✅ 验证安装
运行测试脚本，确认所有功能正常：
```bash
python test_integration.py
```

如果输出 `✅ 所有功能测试通过`，说明安装成功！

---

## ❌ 常见问题
### 1. 中文显示乱码
- 确保系统安装了中文字体
- Linux系统安装 `fonts-noto-cjk` / `google-noto-cjk-fonts` 包
- Windows/macOS一般自带中文字体

### 2. 格式转换失败
- 检查是否安装了LibreOffice
- 确保LibreOffice在系统PATH中

### 3. 依赖安装失败
- 升级pip版本：`pip install --upgrade pip`
- 使用国内镜像源：`pip install -i https://pypi.tuna.tsinghua.edu.cn/simple 包名`
