#!/usr/bin/env bash
#
# Publish batdoc and batdoc-bin to AUR + Homebrew tap.
#
# Reads the version from Cargo.toml, downloads the release assets,
# computes checksums, updates PKGBUILDs/.SRCINFOs/Formula, and pushes.
#
# Prerequisites:
#   - A GitHub release for the version must already exist
#   - SSH keys registered with aur.archlinux.org (or skip AUR)
#   - Push access to daemonp/homebrew-tap (or skip Homebrew)
#   - gh CLI authenticated
#
# Usage:
#   ./pkg/publish-aur.sh          # uses version from Cargo.toml
#   ./pkg/publish-aur.sh 1.2.3    # explicit version override

set -euo pipefail

REPO="daemonp/batdoc"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

# ---------------------------------------------------------------------------
# Resolve version
# ---------------------------------------------------------------------------
if [[ ${1:-} ]]; then
    VERSION="$1"
else
    VERSION=$(sed -n 's/^version = "\(.*\)"/\1/p' "$SCRIPT_DIR/../Cargo.toml")
fi
TAG="v$VERSION"

echo "==> Publishing batdoc $VERSION ($TAG)"

# ---------------------------------------------------------------------------
# Verify the GitHub release exists
# ---------------------------------------------------------------------------
echo "--- Checking GitHub release $TAG ..."
if ! gh release view "$TAG" --repo "$REPO" >/dev/null 2>&1; then
    echo "ERROR: GitHub release $TAG not found." >&2
    echo "Tag and push first:  git tag $TAG && git push origin master $TAG" >&2
    exit 1
fi

# ---------------------------------------------------------------------------
# Download release assets & source tarball, compute checksums
# ---------------------------------------------------------------------------
echo "--- Downloading assets ..."

# Source tarball (for batdoc + homebrew)
curl -sL "https://github.com/$REPO/archive/$TAG.tar.gz" -o "$WORK/source.tar.gz"

# Binary tarballs (for batdoc-bin)
gh release download "$TAG" --repo "$REPO" \
    --pattern "batdoc_${VERSION}_x86_64.tar.gz" \
    --pattern "batdoc_${VERSION}_aarch64.tar.gz" \
    --dir "$WORK"

echo "--- Computing checksums ..."

SOURCE_B2=$(b2sum "$WORK/source.tar.gz"    | awk '{print $1}')
SOURCE_SHA256=$(sha256sum "$WORK/source.tar.gz" | awk '{print $1}')

BIN_X86_SHA256=$(sha256sum "$WORK/batdoc_${VERSION}_x86_64.tar.gz"  | awk '{print $1}')
BIN_ARM_SHA256=$(sha256sum "$WORK/batdoc_${VERSION}_aarch64.tar.gz" | awk '{print $1}')

echo "  source b2sum:       $SOURCE_B2"
echo "  source sha256:      $SOURCE_SHA256"
echo "  bin x86_64 sha256:  $BIN_X86_SHA256"
echo "  bin aarch64 sha256: $BIN_ARM_SHA256"

# ---------------------------------------------------------------------------
# Helper: clone an AUR repo, apply changes, commit, push
# ---------------------------------------------------------------------------
aur_publish() {
    local pkg="$1" msg="$2"
    local aur_dir="$WORK/aur-$pkg"

    echo "--- Updating AUR: $pkg ..."

    if ! git clone "ssh://aur@aur.archlinux.org/$pkg.git" "$aur_dir" 2>/dev/null; then
        echo "  WARNING: could not clone AUR repo (missing SSH key?) — skipping $pkg"
        return 0
    fi

    # Copy updated files (caller must have written them to $WORK/$pkg/)
    cp "$WORK/$pkg/PKGBUILD"  "$aur_dir/PKGBUILD"
    cp "$WORK/$pkg/.SRCINFO"  "$aur_dir/.SRCINFO"

    git -C "$aur_dir" add PKGBUILD .SRCINFO
    if git -C "$aur_dir" diff --cached --quiet; then
        echo "  (no changes — skipping)"
        return
    fi
    git -C "$aur_dir" commit -m "$msg"
    git -C "$aur_dir" push origin master
    echo "  pushed $pkg"
}

# ---------------------------------------------------------------------------
# 1. AUR: batdoc (build-from-source)
# ---------------------------------------------------------------------------
mkdir -p "$WORK/batdoc"

sed -e "s/^pkgver=.*/pkgver=$VERSION/" \
    -e "s/^pkgrel=.*/pkgrel=1/" \
    -e "s/^b2sums=.*/b2sums=('$SOURCE_B2')/" \
    "$SCRIPT_DIR/arch/PKGBUILD" > "$WORK/batdoc/PKGBUILD"

cat > "$WORK/batdoc/.SRCINFO" << EOF
pkgbase = batdoc
	pkgdesc = cat(1) for doc, docx, xls, xlsx, pptx, and pdf -- renders to markdown with bat
	pkgver = $VERSION
	pkgrel = 1
	url = https://github.com/$REPO
	arch = x86_64
	license = MIT
	makedepends = cargo
	depends = gcc-libs
	depends = glibc
	source = batdoc-${VERSION}.tar.gz::https://github.com/$REPO/archive/$TAG.tar.gz
	b2sums = $SOURCE_B2

pkgname = batdoc
EOF

aur_publish "batdoc" "batdoc ${VERSION}-1: update to $TAG"

# ---------------------------------------------------------------------------
# 2. AUR: batdoc-bin (pre-compiled)
# ---------------------------------------------------------------------------
mkdir -p "$WORK/batdoc-bin"

sed -e "s/^pkgver=.*/pkgver=$VERSION/" \
    -e "s/^pkgrel=.*/pkgrel=1/" \
    -e "s/^sha256sums_x86_64=.*/sha256sums_x86_64=('$BIN_X86_SHA256')/" \
    -e "s/^sha256sums_aarch64=.*/sha256sums_aarch64=('$BIN_ARM_SHA256')/" \
    "$SCRIPT_DIR/arch-bin/PKGBUILD" > "$WORK/batdoc-bin/PKGBUILD"

cat > "$WORK/batdoc-bin/.SRCINFO" << EOF
pkgbase = batdoc-bin
	pkgdesc = cat(1) for doc, docx, xls, xlsx, pptx, and pdf -- renders to markdown with bat. Pre-compiled.
	pkgver = $VERSION
	pkgrel = 1
	url = https://github.com/$REPO
	arch = x86_64
	arch = aarch64
	license = MIT
	provides = batdoc
	conflicts = batdoc
	conflicts = batdoc-debug
	options = !debug
	source_x86_64 = https://github.com/$REPO/releases/download/$TAG/batdoc_${VERSION}_x86_64.tar.gz
	sha256sums_x86_64 = $BIN_X86_SHA256
	source_aarch64 = https://github.com/$REPO/releases/download/$TAG/batdoc_${VERSION}_aarch64.tar.gz
	sha256sums_aarch64 = $BIN_ARM_SHA256

pkgname = batdoc-bin
EOF

aur_publish "batdoc-bin" "batdoc-bin ${VERSION}-1: update to $TAG"

# ---------------------------------------------------------------------------
# 3. Homebrew tap
# ---------------------------------------------------------------------------
echo "--- Updating Homebrew tap ..."
TAP_DIR="$WORK/homebrew-tap"

if ! git clone "git@github.com:daemonp/homebrew-tap.git" "$TAP_DIR" 2>/dev/null; then
    echo "  WARNING: could not clone homebrew-tap (missing credentials?) — skipping"
else
    sed -e "s|url \".*\"|url \"https://github.com/$REPO/archive/refs/tags/$TAG.tar.gz\"|" \
        -e "s|sha256 \".*\"|sha256 \"$SOURCE_SHA256\"|" \
        "$SCRIPT_DIR/homebrew/batdoc.rb" > "$TAP_DIR/Formula/batdoc.rb"

    git -C "$TAP_DIR" add Formula/batdoc.rb
    if git -C "$TAP_DIR" diff --cached --quiet; then
        echo "  (no changes — skipping)"
    else
        git -C "$TAP_DIR" commit -m "batdoc $VERSION"
        git -C "$TAP_DIR" push origin main
        echo "  pushed homebrew-tap"
    fi
fi

# ---------------------------------------------------------------------------
# Done
# ---------------------------------------------------------------------------
echo ""
echo "==> Done publishing batdoc $VERSION"
