from distutils.core import setup
from setuptools import find_packages
 
with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()
 
setup(
    name="exceltidy",
    version="0.0.1",
    author="cmacckk",
    author_email="emailforgty@163.com",
    description="Based on openpyxl and xlwings to facilitate Excel processing",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/cmacckk/exceltidy",
    packages=find_packages(),
    install_requires=[
        'xlwings',
        'openpyxl'
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: GNU General Public License v3.0",
        "Operating System :: Windows Mac",
    ],
    python_requires='>=3.6',    #对python的最低版本要求
)