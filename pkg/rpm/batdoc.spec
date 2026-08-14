Name:           batdoc
Version:        1.0.0
Release:        1%{?dist}
Summary:        cat(1) for doc, docx, xls, xlsx, pptx, pdf, and image files (OCR) -- renders to markdown with bat

License:        MIT
URL:            https://github.com/daemonp/batdoc
Source0:        %{url}/archive/v%{version}/%{name}-%{version}.tar.gz

BuildRequires:  cargo
BuildRequires:  rust
BuildRequires:  gcc

%description
Reads legacy .doc and .xls, modern .docx, .xlsx, and .pptx, PDF files, and
raster images (.png, .jpg, .gif, .webp, .bmp) and dumps their text to stdout.
Image files are OCR'd with the ocrs engine; --ocr extends OCR to embedded
DOCX/PPTX images and textless PDF pages. When stdout is a terminal the output
is pretty-printed as syntax-highlighted markdown via bat; when piped, plain
text is emitted. Format is detected by file signature, not extension.

Spiritual successor to catdoc. Pure Rust: no C, no system libs.

%prep
%autosetup -n %{name}-%{version} -p1

%build
cargo build --release --locked

%install
install -Dpm 0755 target/release/%{name} %{buildroot}%{_bindir}/%{name}
install -Dpm 0644 target/man/%{name}.1 %{buildroot}%{_mandir}/man1/%{name}.1

%check
cargo test --locked

%files
%license LICENSE
%doc README.md
%{_bindir}/%{name}
%{_mandir}/man1/%{name}.1*

%changelog
* Fri Aug 14 2026 Damon Petta <d@disassemble.net> - 1.5.0-1
- OCR support: raster image input always OCR'd; --ocr extends OCR to
  embedded DOCX/PPTX images and textless PDF pages
- Extract PPTX speaker notes; DOCX comments, footnotes, and endnotes
- Drop the -o short flag; image OCR renders as plain text on TTY

* Sat Feb 14 2026 Damon Petta <d@disassemble.net> - 1.0.0-1
- Initial package
