#!/bin/bash
export GOOGLE_SERVICE_ACCOUNT_JSON=/root/.credentials/creekside_service_account.json
cd "$(dirname "$0")"
python build_dashboard.py
