"""
Setup configuration for ContaFlow.
"""

from setuptools import setup, find_packages

setup(
    name="contaflow",
    version="2.0.0",
    description="Bot de automatización corporativa para procesamiento de emails y datos",
    author="ContaFlow Team",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    python_requires=">=3.7",
    install_requires=[
        "openpyxl",
        "pandas",
        "pdfplumber",
        "pywin32; platform_system=='Windows'",
    ],
    entry_points={
        "console_scripts": [
            "contaflow=contaflow.main:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: OS Independent",
    ],
)
