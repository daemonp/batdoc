//! CLI integration tests for `--password` and the interactive prompt.
//!
//! The encrypted fixture is built here (rather than reused from `batdoc-core`
//! tests) so the binary is exercised end to end: detect → decrypt → extract.

use std::collections::BTreeMap;
use std::process::{Command, Stdio};
use std::sync::Arc;

/// A one-page PDF with a known text layer, encrypted with user password "user".
fn encrypted_pdf() -> Vec<u8> {
    use lopdf::encryption::crypt_filters::{Aes128CryptFilter, CryptFilter};
    use lopdf::{dictionary, Document, Object, Stream};

    let mut doc = Document::with_version("1.4");
    let pages_node_id = doc.new_object_id();
    let font_id = doc.add_object(dictionary! {
        "Type" => "Font", "Subtype" => "Type1", "BaseFont" => "Helvetica",
        "Encoding" => "WinAnsiEncoding",
    });
    let resources_id = doc.add_object(dictionary! { "Font" => dictionary! { "F1" => font_id } });
    let content_id = doc.add_object(Stream::new(
        lopdf::Dictionary::new(),
        b"BT /F1 12 Tf 72 720 Td (SecretCliText) Tj ET".to_vec(),
    ));
    let page_id = doc.add_object(dictionary! {
        "Type" => "Page", "Parent" => pages_node_id, "Contents" => content_id,
        "Resources" => resources_id,
        "MediaBox" => vec![0.into(), 0.into(), 612.into(), 792.into()],
    });
    doc.objects.insert(
        pages_node_id,
        Object::Dictionary(dictionary! {
            "Type" => "Pages", "Kids" => vec![Object::Reference(page_id)], "Count" => 1,
        }),
    );
    let catalog_id = doc.add_object(dictionary! { "Type" => "Catalog", "Pages" => pages_node_id });
    doc.trailer.set("Root", catalog_id);

    // lopdf derives the file encryption key from the first /ID element, and
    // this document has no trailer /ID yet. Without it `EncryptionState`
    // construction fails with `MissingFileID`.
    let file_id = Object::string_literal("batdoc-cli-file-id");
    doc.trailer.set("ID", vec![file_id.clone(), file_id]);

    let filter: Arc<dyn CryptFilter> = Arc::new(Aes128CryptFilter);
    let version = lopdf::EncryptionVersion::V4 {
        document: &doc,
        encrypt_metadata: true,
        crypt_filters: BTreeMap::from([(b"StdCF".to_vec(), filter)]),
        stream_filter: b"StdCF".to_vec(),
        string_filter: b"StdCF".to_vec(),
        owner_password: "owner",
        user_password: "user",
        permissions: lopdf::Permissions::all(),
    };
    let state = lopdf::EncryptionState::try_from(version).unwrap();
    doc.encrypt(&state).unwrap();
    let mut out = Vec::new();
    doc.save_to(&mut out).unwrap();
    out
}

fn fixture_path(tag: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("batdoc-cli-{tag}-{}.pdf", std::process::id()));
    std::fs::write(&path, encrypted_pdf()).unwrap();
    path
}

#[test]
fn password_flag_extracts() {
    let path = fixture_path("ok");
    let out = Command::new(env!("CARGO_BIN_EXE_batdoc"))
        .args(["--plain", "--password", "user"])
        .arg(&path)
        .output()
        .unwrap();
    let _ = std::fs::remove_file(&path);
    assert!(
        out.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&out.stderr)
    );
    assert!(String::from_utf8_lossy(&out.stdout).contains("SecretCliText"));
}

#[test]
fn missing_password_non_tty_fails() {
    let path = fixture_path("nopw");
    let out = Command::new(env!("CARGO_BIN_EXE_batdoc"))
        .arg(&path)
        .stdin(Stdio::null())
        .output()
        .unwrap();
    let _ = std::fs::remove_file(&path);
    assert!(!out.status.success());
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(stderr.contains("password-protected"), "stderr: {stderr}");
}

#[test]
fn password_without_value_is_usage_error() {
    let out = Command::new(env!("CARGO_BIN_EXE_batdoc"))
        .arg("--password")
        .output()
        .unwrap();
    assert_eq!(out.status.code(), Some(2));
}

fn secret_docx() -> Vec<u8> {
    use std::io::Write;
    let mut buf = std::io::Cursor::new(Vec::new());
    let mut z = zip::ZipWriter::new(&mut buf);
    z.start_file(
        "word/document.xml",
        zip::write::SimpleFileOptions::default(),
    )
    .unwrap();
    z.write_all(b"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>SecretOfficeCli</w:t></w:r></w:p></w:body></w:document>")
        .unwrap();
    z.finish().unwrap();
    buf.into_inner()
}

fn office_fixture(tag: &str) -> std::path::PathBuf {
    let path = std::env::temp_dir().join(format!("batdoc-cli-{tag}-{}.docx", std::process::id()));
    let enc = msoffice_crypto::encrypt_ooxml(&secret_docx(), "pw").unwrap();
    std::fs::write(&path, enc).unwrap();
    path
}

#[test]
fn office_password_flag_extracts() {
    let path = office_fixture("office-ok");
    let out = Command::new(env!("CARGO_BIN_EXE_batdoc"))
        .args(["--plain", "--password", "pw"])
        .arg(&path)
        .stdin(Stdio::null())
        .output()
        .unwrap();
    let _ = std::fs::remove_file(&path);
    assert!(
        out.status.success(),
        "stderr: {}",
        String::from_utf8_lossy(&out.stderr)
    );
    assert!(String::from_utf8_lossy(&out.stdout).contains("SecretOfficeCli"));
}

#[test]
fn office_missing_password_non_tty_fails() {
    let path = office_fixture("office-nopw");
    let out = Command::new(env!("CARGO_BIN_EXE_batdoc"))
        .args(["--plain"])
        .arg(&path)
        .stdin(Stdio::null())
        .output()
        .unwrap();
    let _ = std::fs::remove_file(&path);
    assert!(!out.status.success());
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(stderr.contains("password-protected"), "stderr: {stderr}");
}

#[test]
fn office_wrong_password_exits_without_prompting() {
    let path = office_fixture("office-bad");
    let out = Command::new(env!("CARGO_BIN_EXE_batdoc"))
        .args(["--plain", "--password", "nope"])
        .arg(&path)
        .stdin(Stdio::null())
        .output()
        .unwrap();
    let _ = std::fs::remove_file(&path);
    assert_eq!(out.status.code(), Some(1));
    let stderr = String::from_utf8_lossy(&out.stderr);
    assert!(stderr.contains("incorrect password"), "stderr: {stderr}");
    assert!(!stderr.contains("Password:"), "stderr: {stderr}");
    assert!(!stderr.contains("nope") && !stderr.contains("pw"));
}
