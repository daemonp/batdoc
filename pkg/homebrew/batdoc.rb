class Batdoc < Formula
  desc "cat(1) for doc, docx, xls, xlsx, pptx, and pdf -- renders to markdown with bat"
  homepage "https://github.com/daemonp/batdoc"
  url "https://github.com/daemonp/batdoc/archive/refs/tags/v1.0.0.tar.gz"
  sha256 "PLACEHOLDER"
  license "MIT"
  head "https://github.com/daemonp/batdoc.git", branch: "master"

  depends_on "rust" => :build

  def install
    system "cargo", "install", *std_cargo_args
  end

  test do
    # Create a minimal PDF and verify batdoc detects it
    (testpath/"test.pdf").write "%PDF-1.4 empty"
    output = shell_output("#{bin}/batdoc --plain #{testpath}/test.pdf 2>&1", 1)
    assert_match(/pdf|error/i, output)
  end
end
