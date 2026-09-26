#!/usr/bin/env bash
set -euo pipefail

repo_root="${BUILD_WORKSPACE_DIRECTORY:?Run this script through Bazel}/services/doc-mcp/external/docx-editor"
cd "$repo_root"
if ! command -v bun >/dev/null 2>&1; then
  echo "bun is required to build the DOCX workspace packages" >&2
  exit 1
fi
export NODE_OPTIONS="${NODE_OPTIONS:---max-old-space-size=16384}"
exec bun run build:packages
