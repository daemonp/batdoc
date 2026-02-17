# Multi-stage build: produces .deb, .rpm, .apk, and Arch .pkg.tar.zst
# from a single source tree.
#
# Usage:
#   docker build --build-arg VERSION=1.0.0 -o pkg/out .
#
# Each stage builds batdoc on its native distro, then the final stage
# collects all packages into /out for export.

ARG VERSION=1.0.0

# ---------------------------------------------------------------------------
# Stage 1: Debian .deb
# ---------------------------------------------------------------------------
FROM debian:bookworm-slim AS deb

ARG VERSION
ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    curl \
    ca-certificates \
    debhelper \
    fakeroot \
    && rm -rf /var/lib/apt/lists/*

# Install Rust via rustup (distro cargo is too old on bookworm)
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs \
    | sh -s -- -y --default-toolchain stable --profile minimal
ENV PATH="/root/.cargo/bin:${PATH}"

WORKDIR /build/batdoc-${VERSION}
COPY . .

# Build the binary
RUN cargo build --release --locked

# Assemble the .deb with fpm-free dpkg-deb
RUN mkdir -p dpkg/usr/bin dpkg/usr/share/doc/batdoc dpkg/usr/share/man/man1 dpkg/DEBIAN /out \
    && install -m755 target/release/batdoc dpkg/usr/bin/batdoc \
    && install -m644 target/man/batdoc.1 dpkg/usr/share/man/man1/batdoc.1 \
    && gzip -9 dpkg/usr/share/man/man1/batdoc.1 \
    && install -m644 README.md dpkg/usr/share/doc/batdoc/README \
    && install -m644 LICENSE dpkg/usr/share/doc/batdoc/copyright \
    && printf 'Package: batdoc\nVersion: %s\nSection: utils\nPriority: optional\nArchitecture: amd64\nMaintainer: Damon Petta <d@disassemble.net>\nDescription: cat(1) for doc, docx, xls, xlsx, pptx, and pdf -- renders to markdown with bat\n Reads legacy .doc and .xls, modern .docx, .xlsx, and .pptx, and PDF files\n and dumps their text to stdout as syntax-highlighted markdown via bat.\n' "${VERSION}" > dpkg/DEBIAN/control \
    && dpkg-deb --build dpkg /out/batdoc_${VERSION}-1_amd64.deb

# ---------------------------------------------------------------------------
# Stage 2: Fedora .rpm
# ---------------------------------------------------------------------------
FROM fedora:41 AS rpm

ARG VERSION

RUN dnf install -y gcc rpm-build curl && dnf clean all

# Install Rust via rustup (Fedora 41 ships 1.91, too old for bat deps)
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs \
    | sh -s -- -y --default-toolchain stable --profile minimal
ENV PATH="/root/.cargo/bin:${PATH}"

WORKDIR /build
COPY . src/

# Create rpmbuild tree with source in place
RUN mkdir -p rpmbuild/{BUILD,RPMS,SOURCES,SPECS,SRPMS} \
    && tar czf rpmbuild/SOURCES/batdoc-${VERSION}.tar.gz \
       --transform="s,^src,batdoc-${VERSION}," src/ \
    && cp src/pkg/rpm/batdoc.spec rpmbuild/SPECS/ \
    && sed -i "s|^Version:.*|Version:        ${VERSION}|" rpmbuild/SPECS/batdoc.spec

RUN rpmbuild --define "_topdir /build/rpmbuild" \
    --nodeps -bb rpmbuild/SPECS/batdoc.spec

RUN mkdir -p /out && cp rpmbuild/RPMS/x86_64/*.rpm /out/

# ---------------------------------------------------------------------------
# Stage 3: Alpine .apk
# ---------------------------------------------------------------------------
FROM alpine:3.21 AS apk

ARG VERSION

RUN apk add --no-cache \
    alpine-sdk \
    curl \
    sudo

# abuild needs a non-root user in the abuild group
RUN adduser -D builder && addgroup builder abuild \
    && echo "builder ALL=(ALL) NOPASSWD: ALL" >> /etc/sudoers

# Install Rust via rustup (Alpine 3.21 ships 1.83, too old for bat deps)
USER builder
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs \
    | sh -s -- -y --default-toolchain stable --profile minimal
ENV PATH="/home/builder/.cargo/bin:${PATH}"

WORKDIR /home/builder

# Generate a throwaway signing key (required by abuild)
RUN abuild-keygen -ain

# Set up the build directory
RUN mkdir -p package/src
COPY --chown=builder:builder . package/src/

# Create a local tarball so APKBUILD source= resolves without network
RUN cd package && tar czf batdoc-${VERSION}.tar.gz \
    --transform="s,^src,batdoc-${VERSION}," src/

# Write an APKBUILD that uses the local tarball instead of fetching from GitHub
RUN printf '# Maintainer: Damon Petta <d@disassemble.net>\n\
pkgname=batdoc\n\
pkgver=%s\n\
pkgrel=0\n\
pkgdesc="cat(1) for doc, docx, xls, xlsx, pptx, and pdf -- renders to markdown with bat"\n\
url="https://github.com/daemonp/batdoc"\n\
arch="x86_64"\n\
license="MIT"\n\
makedepends=""\n\
source="batdoc-$pkgver.tar.gz"\n\
\n\
build() {\n\
\tcargo build --release --locked\n\
}\n\
\n\
check() {\n\
\tcargo test --locked\n\
}\n\
\n\
package() {\n\
\tinstall -Dm755 target/release/batdoc -t "$pkgdir"/usr/bin/\n\
\tinstall -Dm644 target/man/batdoc.1 "$pkgdir"/usr/share/man/man1/batdoc.1\n\
\tgzip -9 "$pkgdir"/usr/share/man/man1/batdoc.1\n\
\tinstall -Dm644 LICENSE -t "$pkgdir"/usr/share/licenses/$pkgname/\n\
}\n\
\n\
sha512sums=""\n' "${VERSION}" > package/APKBUILD

WORKDIR /home/builder/package
RUN abuild checksum && abuild -F -r

RUN sudo mkdir -p /out && sudo find ~/packages -name 'batdoc-*.apk' -exec cp {} /out/ \;

# ---------------------------------------------------------------------------
# Stage 4: Arch Linux .pkg.tar.zst
# ---------------------------------------------------------------------------
FROM archlinux:base-devel AS arch

ARG VERSION

RUN pacman -Syu --noconfirm rust && pacman -Scc --noconfirm

# makepkg refuses to run as root
RUN useradd -m builder \
    && echo "builder ALL=(ALL) NOPASSWD: ALL" >> /etc/sudoers

USER builder
WORKDIR /home/builder/build

COPY --chown=builder:builder . src/

# Create source tarball matching what PKGBUILD expects
RUN tar czf batdoc-${VERSION}.tar.gz \
    --transform="s,^src,batdoc-${VERSION}," src/

# Write PKGBUILD that uses local tarball with correct version
COPY --chown=builder:builder pkg/arch/PKGBUILD .
RUN sed -i \
    -e "s|^pkgver=.*|pkgver=${VERSION}|" \
    -e "s|source=.*|source=(\"batdoc-${VERSION}.tar.gz\")|" \
    PKGBUILD

RUN makepkg -sf --noconfirm --skipchecksums

RUN sudo mkdir -p /out && sudo cp batdoc-*.pkg.tar.zst /out/

# ---------------------------------------------------------------------------
# Final stage: collect all artifacts
# ---------------------------------------------------------------------------
FROM scratch AS output

COPY --from=deb  /out/ /
COPY --from=rpm  /out/ /
COPY --from=apk  /out/ /
COPY --from=arch /out/ /
