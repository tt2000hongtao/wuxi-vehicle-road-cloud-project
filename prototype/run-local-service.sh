#!/bin/bash
set -euo pipefail

cd "$(dirname "$0")/.."
exec /opt/homebrew/bin/node prototype/server.js
