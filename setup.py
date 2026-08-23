
from setuptools import setup, find_packages
from src.office_suite import __version__, __author__

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith("#")]

setup(
    name="acks-office-suite-simplify",
    version=__version__,
    author=__author__,
    author_email="mail@jintao.uk",
    description="ACKS Office Suite Simplify —— 全能 Python 办公自动化工具包，内置 ACKS 设计规范，单个 API 搞定 Word/Excel/PDF/PPT 四种格式，中文友好，适配 Hermes / DSH / MCP 等 AI Agent。",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/shynloc/ACKS-Office-Suite-Simplify",
    project_urls={
        "Bug Tracker": "https://github.com/shynloc/ACKS-Office-Suite-Simplify/issues",
        "Source": "https://github.com/shynloc/ACKS-Office-Suite-Simplify",
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Natural Language :: Chinese (Simplified)",
    ],
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    package_data={
        "office_suite.design_system": ["tokens.json", "assets/*.png"],
    },
    python_requires=">=3.9",
    install_requires=requirements,
    extras_require={
        "full": [
            "Jinja2>=3.1.0",
            "click>=8.1.0",
            "rich>=13.0.0",
            "python-magic>=0.4.27",
            "python-dotenv>=1.0.0",
        ],
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
        ],
    },
    keywords=[
        "office", "excel", "word", "powerpoint", "pdf",
        "docx", "xlsx", "pptx", "办公自动化", "automation",
        "report-generator", "hermes", "ai-agent", "design-system",
        "acks", "mcp", "dsh", "deepseek", "office-suite", "document-generation"
    ],
    include_package_data=True,
)
