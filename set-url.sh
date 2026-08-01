#!/usr/bin/env bash
# Injects your GitHub Pages base URL into manifest.xml.
# Usage:  ./set-url.sh <github-username> <repo-name>
# Example: ./set-url.sh shubich followup-addin
#   -> https://shubich.github.io/followup-addin
set -e
if [ $# -ne 2 ]; then echo "Usage: ./set-url.sh <github-username> <repo-name>"; exit 1; fi
USER="$1"; REPO="$2"
BASEURL="https://${USER}.github.io/${REPO}"
# portable in-place edit (Linux & macOS)
if sed --version >/dev/null 2>&1; then
  sed -i "s#__BASEURL__#${BASEURL}#g" manifest.xml
else
  sed -i '' "s#__BASEURL__#${BASEURL}#g" manifest.xml
fi
echo "manifest.xml base URL set to: ${BASEURL}"
echo "Remaining __BASEURL__ placeholders: $(grep -c '__BASEURL__' manifest.xml || true)"
