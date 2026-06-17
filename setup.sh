#!/bin/bash
set -e
pip install -r requirements.txt
playwright install chromium
echo "Setup complete. Copy .env.example to .env and add your Google API key."
