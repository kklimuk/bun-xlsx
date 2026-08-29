#!/bin/sh
# Install a checksummed xlsx standalone binary from a GitHub release.
# Usage: sh install.sh
#        PREFIX=/usr/local/bin VERSION=v0.1.0 sh install.sh

set -eu

REPO="kklimuk/bun-xlsx"
PREFIX="${PREFIX:-${HOME-}/.local/bin}"
PREFIX="${PREFIX%/}"
VERSION="${VERSION:-latest}"
REQUIRE_CHECKSUM="${REQUIRE_CHECKSUM:-0}"

version_without_v="${VERSION#v}"
if [ "$VERSION" != "latest" ]; then
  echo "$VERSION" | awk '/^v?[0-9]+\.[0-9]+\.[0-9]+([.-][0-9A-Za-z.-]+)?$/ { ok=1 } END { exit !ok }' || {
    echo "Error: VERSION must be latest or a release tag like v0.2.0." >&2
    exit 1
  }
fi

npm_hint() {
  echo "  Install from npm instead: bun add -g @sageling/xlsx" >&2
  echo "  Or with npm: npm install -g @sageling/xlsx" >&2
}
prefix_hint() { echo "  Pick a writable location: PREFIX=/usr/local/bin sh install.sh" >&2; }

os="$(uname -s)"
arch="$(uname -m)"

detect_target() {
  case "$os" in
    Linux)
      case "$arch" in
        x86_64|amd64) echo "xlsx-linux-x64" ;;
        aarch64|arm64) echo "xlsx-linux-arm64" ;;
        *) return 1 ;;
      esac
      ;;
    Darwin)
      case "$arch" in
        x86_64) echo "xlsx-darwin-x64" ;;
        arm64) echo "xlsx-darwin-arm64" ;;
        *) return 1 ;;
      esac
      ;;
    MINGW*|MSYS*|CYGWIN*) echo "xlsx-windows-x64.exe" ;;
    *) return 1 ;;
  esac
}

if command -v curl >/dev/null 2>&1; then
  download() { curl -fsSL "$1" -o "$2"; }
elif command -v wget >/dev/null 2>&1; then
  download() { wget -qO "$2" "$1"; }
else
  echo "Error: need curl or wget to install." >&2
  npm_hint
  exit 1
fi

verify=1
if command -v sha256sum >/dev/null 2>&1; then
  sha256_of() { sha256sum "$1" | awk '{print $1}'; }
elif command -v shasum >/dev/null 2>&1; then
  sha256_of() { shasum -a 256 "$1" | awk '{print $1}'; }
elif command -v openssl >/dev/null 2>&1; then
  sha256_of() { openssl dgst -sha256 "$1" | awk '{print $NF}'; }
else
  verify=0
fi

target="$(detect_target)" || {
  echo "Error: unsupported platform: ${os} ${arch}." >&2
  echo "  Supported: linux/x64, linux/arm64, darwin/x64, darwin/arm64, windows/x64." >&2
  npm_hint
  exit 1
}
binary_name="xlsx"
case "$target" in *.exe) binary_name="xlsx.exe" ;; esac

if [ "$verify" = 0 ]; then
  if [ "$REQUIRE_CHECKSUM" = "1" ]; then
    echo "Error: no sha256sum, shasum, or openssl; refusing an unverified install." >&2
    exit 1
  fi
  echo "Warning: installing without SHA-256 verification (no checksum tool found)." >&2
fi

if [ "$VERSION" = "latest" ]; then
  base="https://github.com/${REPO}/releases/latest/download"
else
  base="https://github.com/${REPO}/releases/download/${VERSION}"
fi

fetch_asset() {
  download "${base}/$1" "$2" && [ -s "$2" ] || {
    echo "Error: could not download ${base}/$1." >&2
    exit 1
  }
}

mkdir -p "$PREFIX" || { echo "Error: could not create ${PREFIX}." >&2; prefix_hint; exit 1; }
stage_dir="$(mktemp -d "${PREFIX}/.xlsx-install.XXXXXX")" || {
  echo "Error: ${PREFIX} is not writable." >&2
  prefix_hint
  exit 1
}
bin_tmp="${stage_dir}/${binary_name}"
sums_tmp="${stage_dir}/SHA256SUMS"
trap 'rm -rf "$stage_dir"' EXIT INT TERM

expected=""
if [ "$verify" = 1 ]; then
  fetch_asset SHA256SUMS "$sums_tmp"
  expected="$(awk -v f="$target" '$2 == f || $2 == "*"f { print $1; exit }' "$sums_tmp")"
  [ -n "$expected" ] || { echo "Error: no checksum for ${target}." >&2; exit 1; }
fi

echo "Downloading ${target} from ${base}/${target}"
fetch_asset "$target" "$bin_tmp"
if [ -n "$expected" ]; then
  actual="$(sha256_of "$bin_tmp")"
  [ "$actual" = "$expected" ] || {
    echo "Error: SHA-256 mismatch for ${target}." >&2
    exit 1
  }
  echo "Verified SHA-256 (${target})"
fi

chmod +x "$bin_tmp"
staged_version="$("$bin_tmp" --version)" || {
  echo "Error: downloaded ${target} cannot run on this system; existing install is unchanged." >&2
  exit 1
}
case "$staged_version" in
  "xlsx "*) ;;
  *) echo "Error: downloaded binary returned an invalid version; existing install is unchanged." >&2; exit 1 ;;
esac
if [ "$VERSION" != "latest" ]; then
  expected_version="xlsx ${version_without_v}"
  [ "$staged_version" = "$expected_version" ] || {
    echo "Error: requested ${VERSION}, but downloaded binary reports ${staged_version}; existing install is unchanged." >&2
    exit 1
  }
fi

mv "$bin_tmp" "${PREFIX}/${binary_name}"
echo "Installed: ${PREFIX}/${binary_name}"

case ":${PATH}:" in
  *":$PREFIX:"*) ;;
  *) echo "Note: ${PREFIX} is not on PATH. Add it to your shell profile." ;;
esac

echo "$staged_version"
