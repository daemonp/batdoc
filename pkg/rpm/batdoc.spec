Name:           batdoc
Version:        1.0.0
Release:        1%{?dist}
Summary:        cat(1) for doc, docx, xls, xlsx, pptx, and pdf -- renders to markdown with bat

License:        MIT
URL:            https://github.com/daemonp/batdoc
Source0:        %{url}/archive/v%{version}/%{name}-%{version}.tar.gz

BuildRequires:  cargo
BuildRequires:  rust
BuildRequires:  gcc

%description
Reads legacy .doc and .xls, modern .docx, .xlsx, and .pptx, and PDF files
and dumps their text to stdout. When stdout is a terminal the output is
pretty-printed as syntax-highlighted markdown via bat; when piped, plain text
is emitted. Format is detected by file signature, not extension.

Spiritual successor to catdoc. Seven crates, no C, no system libs.

%prep
%autosetup -n %{name}-%{version} -p1

%build
cargo build --release --locked

%install
install -Dpm 0755 target/release/%{name} %{buildroot}%{_bindir}/%{name}

%check
cargo test --locked

%files
%license LICENSE
%doc README.md
%{_bindir}/%{name}

%changelog
* Sat Feb 14 2026 Damon Petta <d@disassemble.net> - 1.0.0-1
- Initial package
