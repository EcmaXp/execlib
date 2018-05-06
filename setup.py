from setuptools import setup, find_packages

setup(
    name="execlib",
    version="0.1.0",
    author="EcmaXp",
    author_email="module-execlib@ecmaxp.pe.kr",
    description="",
    keywords="excel",
    url="http://github.com/EcmaXp/execlib",
    packages=find_packages(),
    package_data={
        "execlib": [
            "PyExcelLibrary.dll",
        ],
    },
    install_requires=[
        "pythonnet",
        "dataclasses",
    ],
    classifiers=[
        "Development Status :: 2 - Pre-Alpha",
        "Environment :: Win32 (MS Windows)",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: C#",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
    ],
)
