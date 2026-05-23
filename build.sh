#!/usr/bin/env bash
set -e

# Install system packages yang dibutuhkan
apt-get update -qq
apt-get install -y -qq \
    poppler-utils \
    ffmpeg \
    libgl1

# Install Python dependencies
pip install -r requirements.txt
