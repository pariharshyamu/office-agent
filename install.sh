#!/bin/sh
# officecli installer for Linux and macOS.
#
#   curl -fsSL https://github.com/pariharshyamu/office-agent/releases/latest/download/install.sh | sh
#
# Options (environment variables):
#   OFFICECLI_VERSION   install a specific tag (e.g. v0.2.0); default: latest
#   OFFICECLI_INSTALL   install directory; default: ~/.local/bin
set -eu

REPO="pariharshyamu/office-agent"
VERSION="${OFFICECLI_VERSION:-latest}"
INSTALL_DIR="${OFFICECLI_INSTALL:-$HOME/.local/bin}"

os="$(uname -s)"
arch="$(uname -m)"
case "$os" in
  Linux)
    case "$arch" in
      x86_64 | amd64) asset="officecli-linux-x64" ;;
      aarch64 | arm64) asset="officecli-linux-arm64" ;;
      *) echo "error: unsupported Linux architecture '$arch'" >&2; exit 1 ;;
    esac
    ;;
  Darwin)
    case "$arch" in
      arm64) asset="officecli-mac-arm64" ;;
      x86_64) asset="officecli-mac-x64" ;;
      *) echo "error: unsupported macOS architecture '$arch'" >&2; exit 1 ;;
    esac
    ;;
  *)
    echo "error: unsupported OS '$os' (on Windows, use install.ps1)" >&2
    exit 1
    ;;
esac

if [ "$VERSION" = "latest" ]; then
  url="https://github.com/$REPO/releases/latest/download/$asset"
else
  url="https://github.com/$REPO/releases/download/$VERSION/$asset"
fi

echo "Downloading $asset ($VERSION) ..."
tmp="$(mktemp)"
trap 'rm -f "$tmp"' EXIT
if command -v curl >/dev/null 2>&1; then
  curl -fSL --progress-bar -o "$tmp" "$url"
elif command -v wget >/dev/null 2>&1; then
  wget -qO "$tmp" "$url"
else
  echo "error: need curl or wget" >&2
  exit 1
fi

mkdir -p "$INSTALL_DIR"
install -m 755 "$tmp" "$INSTALL_DIR/officecli"

echo "Installed $("$INSTALL_DIR/officecli" --version) to $INSTALL_DIR/officecli"

case ":$PATH:" in
  *":$INSTALL_DIR:"*) ;;
  *)
    echo
    echo "NOTE: $INSTALL_DIR is not on your PATH. Add it with:"
    echo "  export PATH=\"$INSTALL_DIR:\$PATH\""
    ;;
esac

echo
echo "Get started:   officecli help"
echo "MCP server:    claude mcp add officecli -- $INSTALL_DIR/officecli mcp"
