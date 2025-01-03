# 2024.10.10 BK Choi

from setuptools import setup, find_packages

try:
    import chardet
except ImportError:
    chardet = None

# 파일 인코딩을 감지하는 함수
def detect_encoding(filename):
    if chardet:
        with open(filename, "rb") as file:
            raw_data = file.read()
            result = chardet.detect(raw_data)
            return result['encoding']
    return 'utf-8'  # chardet이 없으면 기본 인코딩으로 utf-8 사용

# requirements.txt 파일에서 의존성 목록을 읽어오는 함수
def parse_requirements(filename, encoding='utf-8'):
    try:
        with open(filename, "r", encoding=encoding) as file:
            install_packages = [
                line.strip() for line in file if line.strip() and not line.startswith("#")
            ]
        return install_packages
    except FileNotFoundError:
        print(f"Warning: {filename} not found.")
        return []

# requirements.txt 파일 인코딩 감지
requirements_encoding = detect_encoding("requirements.txt")

# long_description 파일 예외 처리
try:
    with open("README.md", "r", encoding="utf-8") as fh:
        long_description = fh.read()
except FileNotFoundError:
    long_description = "This program helps enhance user convenience when saving a pandas DataFrame to an Excel file by adding Excel sheets and Excel charts."

setup(
    name="bkexcel",
    version="0.1.3",
    author="BK Choi",
    author_email="stsboongkee@gmail.com",
    description="This program helps enhance user convenience when saving a pandas DataFrame to an Excel file by adding Excel sheets and Excel charts.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/skyboong/bkexcel",
    packages=find_packages(),
    install_requires=parse_requirements("requirements.txt", encoding=requirements_encoding),
    classifiers=[
        "Development Status :: 2 - Pre-Alpha",
        "Intended Audience :: Developers",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.9',
)
