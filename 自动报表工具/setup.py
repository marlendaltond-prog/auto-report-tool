from setuptools import setup

setup(
    name="auto-report-generator",
    version="1.0.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="自动化报表生成工具",
    long_description=open("README.md", "r", encoding="utf-8").read() if open("README.md", "r", encoding="utf-8") else "自动化报表生成工具",
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/auto-report-generator",
    py_modules=['auto_report', 'update_manager'],
    package_data={'': ['*.json', '*.yaml', 'templates/*']},
    include_package_data=True,
    install_requires=[
        "pandas>=2.0.0",
        "numpy>=1.20.0",
        "openpyxl>=3.0.0",
        "matplotlib>=3.0.0",
        "seaborn>=0.10.0",
        "sqlalchemy>=2.0.0",
        "requests>=2.0.0",
        "pydantic>=2.0.0",
        "python-crontab>=2.0.0",
        "PyYAML>=6.0.0"
    ],
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "flake8>=6.0.0",
            "black>=23.0.0"
        ]
    },
    entry_points={
        "console_scripts": [
            "auto-report=auto_report:main",
        ],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: Business Owners",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Topic :: Office/Business",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    python_requires=">=3.10",
    keywords="report, excel, automation, data-analysis, business-intelligence",
    project_urls={
        "Documentation": "https://github.com/yourusername/auto-report-generator/wiki",
        "Source": "https://github.com/yourusername/auto-report-generator",
        "Tracker": "https://github.com/yourusername/auto-report-generator/issues",
    },
)
