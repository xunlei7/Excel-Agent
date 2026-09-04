#!/bin/zsh
set -e

runtime_root="${CODEX_RUNTIME_ROOT:-${HOME}/.cache/codex-runtimes/codex-primary-runtime/dependencies}"
node_modules="${CODEX_NODE_MODULES:-$runtime_root/node/node_modules}"
python_binary="${WEALTH_PYTHON:-}"

if [[ -z "$python_binary" ]]; then
  for candidate in \
    "${HOME}/miniconda3/envs/money/bin/python" \
    "${HOME}/anaconda3/envs/money/bin/python" \
    "${PWD}/.venv/bin/python"; do
    if [[ -x "$candidate" ]]; then
      python_binary="$candidate"
      break
    fi
  done
fi

if [[ ! -d "$node_modules" ]]; then
  echo "找不到 Codex 表格运行依赖。请先打开一次 ChatGPT/Codex 桌面版以安装运行时。" >&2
  exit 1
fi

if [[ -z "$python_binary" || ! -x "$python_binary" ]]; then
  echo "找不到可用的 Python。请设置 WEALTH_PYTHON，或创建 Conda money/.venv 环境。" >&2
  exit 1
fi

if ! "$python_binary" -c 'import yfinance' >/dev/null 2>&1; then
  echo "Conda money 环境缺少 yfinance：$python_binary" >&2
  exit 1
fi

export CODEX_NODE_MODULES="$node_modules"
export WEALTH_PYTHON="$python_binary"
export PATH="$runtime_root/node/bin:$runtime_root/bin/override:$runtime_root/bin/fallback:$PATH"

codex_binary="$(command -v codex || true)"
if [[ -z "$codex_binary" ]]; then
  codex_binary="/Applications/ChatGPT.app/Contents/Resources/codex"
fi

if [[ ! -x "$codex_binary" ]]; then
  echo "找不到 Codex CLI。请先安装 Codex CLI。" >&2
  exit 1
fi

exec "$codex_binary" "$@"
