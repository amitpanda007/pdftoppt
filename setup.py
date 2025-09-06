"""Setup script for pdftoppt package."""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="pdftoppt",
    version="1.0.0",
    author="Amit Panda",
    author_email="amitpanda007@gmail.com",
    description="Advanced PDF to PowerPoint converter with high fidelity",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/amitpanda007/pdftoppt",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Multimedia :: Graphics :: Graphics Conversion",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    python_requires=">=3.7",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=6.0",
            "pytest-cov>=2.0",
            "black>=21.0",
            "flake8>=3.8",
            "mypy>=0.800",
        ],
    },
    entry_points={
        "console_scripts": [
            "pdftoppt=pdftoppt.cli:main",
        ],
    },
    keywords="pdf powerpoint converter presentation office",
    project_urls={
        "Bug Reports": "https://github.com/amitpanda007/pdftoppt/issues",
        "Source": "https://github.com/amitpanda007/pdftoppt",
        "Documentation": "https://github.com/amitpanda007/pdftoppt#readme",
    },
)
