# .devcontainer/Dockerfile

# ベースイメージとして公式のPythonイメージを使用
FROM python:3.10-slim

# 必要なシステムパッケージをインストール
RUN apt-get update && apt-get install -y \
    build-essential \
    libgl1-mesa-glx \
    tk \
    tk-dev \
    libx11-dev \
    libxext-dev \
    libxrender-dev \
    libxft-dev \
    libjpeg-dev \
    libfreetype6-dev \
    locales \
    fonts-ipafont \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# ロケールの設定
RUN sed -i '/en_US.UTF-8/s/^# //g' /etc/locale.gen && \
    sed -i '/ja_JP.UTF-8/s/^# //g' /etc/locale.gen && \
    locale-gen

ENV LANG ja_JP.UTF-8
ENV LANGUAGE ja_JP:ja
ENV LC_ALL ja_JP.UTF-8

# 作業ディレクトリを設定
WORKDIR /workspace

# pipを最新バージョンに更新
RUN pip install --upgrade pip

# 必要なPythonパッケージをインストール
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーションのソースコードをコピー
COPY . .

# デフォルトのコマンド
CMD ["python", "main.py"]
