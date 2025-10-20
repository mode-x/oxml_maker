# frozen_string_literal: true

require "test_helper"
require "fileutils"
require "tempfile"
require "zip"

class TestDocument < Minitest::Test
  def setup
    @temp_dir = Dir.mktmpdir
    @sample_params = {
      sections: [
        { paragraph: { text: "Hello, World!" } }
      ],
      page_size: { width: 12_240, height: 15_840 },
      page_margin: {
        top: 1440, right: 1440, bottom: 1440, left: 1440,
        header: 720, footer: 720, gutter: 0
      }
    }

    # Create mock docx template structure
    @mock_docx_root = File.join(@temp_dir, "docx")
    create_mock_docx_structure
  end

  def teardown
    FileUtils.rm_rf(@temp_dir) if @temp_dir && Dir.exist?(@temp_dir)
  end

  def test_initialize_with_filename
    doc = OxmlMaker::Document.new(filename: "test.docx")

    assert_equal "test.docx", doc.filename
  end

  def test_initialize_with_params
    doc = OxmlMaker::Document.new(params: @sample_params)

    assert_equal @sample_params, doc.params
  end

  def test_initialize_with_block
    result = nil
    doc = OxmlMaker::Document.new do |d|
      result = d
    end

    assert_equal doc, result
  end

  def test_template_includes_xml_declaration
    doc = OxmlMaker::Document.new(params: @sample_params)
    template = doc.template

    assert_includes template, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  end

  def test_template_includes_document_structure
    doc = OxmlMaker::Document.new(params: @sample_params)
    template = doc.template

    assert_includes template, '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    assert_includes template, "<w:body>"
    assert_includes template, "</w:body>"
    assert_includes template, "</w:document>"
  end

  def test_template_includes_page_settings
    doc = OxmlMaker::Document.new(params: @sample_params)
    template = doc.template

    assert_includes template, "<w:sectPr>"
    assert_includes template, "<w:pgSz w:w='12240' w:h='15840'/>"
    assert_includes template, "w:top='1440'"
    assert_includes template, "w:right='1440'"
    assert_includes template, "w:bottom='1440'"
    assert_includes template, "w:left='1440'"
  end

  def test_template_processes_paragraph_sections
    doc = OxmlMaker::Document.new(params: @sample_params)
    template = doc.template

    assert_includes template, "<w:t>Hello, World!</w:t>"
  end

  def test_template_with_empty_sections
    params = @sample_params.dup
    params[:sections] = []

    doc = OxmlMaker::Document.new(params: params)
    template = doc.template

    assert_includes template, "<w:body>"
    assert_includes template, "</w:body>"
  end

  def test_prepare_paths_sets_instance_variables
    doc = OxmlMaker::Document.new
    doc.send(:prepare_paths)

    assert doc.instance_variable_get(:@docx_root)
    assert doc.instance_variable_get(:@docx_word_path)
    assert doc.instance_variable_get(:@zip_path)
    assert doc.instance_variable_get(:@docx_path)
  end

  def test_ensure_directories_creates_word_folder
    doc = OxmlMaker::Document.new
    doc.send(:prepare_paths)

    # Mock the word path to use temp directory
    doc.instance_variable_set(:@docx_word_path, File.join(@temp_dir, "word"))

    refute Dir.exist?(File.join(@temp_dir, "word"))
    doc.send(:ensure_directories)
    assert Dir.exist?(File.join(@temp_dir, "word"))
  end

  def test_write_document_xml_creates_file
    doc = OxmlMaker::Document.new(params: @sample_params)
    doc.send(:prepare_paths)
    doc.instance_variable_set(:@docx_word_path, @temp_dir)

    doc.send(:write_document_xml)

    document_path = File.join(@temp_dir, "document.xml")
    assert File.exist?(document_path)

    content = File.read(document_path)
    assert_includes content, "Hello, World!"
  end

  def test_cleanup_temp_files_removes_files
    doc = OxmlMaker::Document.new
    doc.send(:prepare_paths)

    # Create temporary files
    zip_path = File.join(@temp_dir, "document.zip")
    docx_path = File.join(@temp_dir, "document.docx")

    File.write(zip_path, "test")
    File.write(docx_path, "test")

    doc.instance_variable_set(:@zip_path, zip_path)
    doc.instance_variable_set(:@docx_path, docx_path)

    assert File.exist?(zip_path)
    assert File.exist?(docx_path)

    doc.send(:cleanup_temp_files)

    refute File.exist?(zip_path)
    refute File.exist?(docx_path)
  end

  def test_zip_docx_folder_creates_valid_zip
    doc = OxmlMaker::Document.new
    doc.send(:prepare_paths)

    # Use our mock docx structure
    doc.instance_variable_set(:@docx_root, @mock_docx_root)
    doc.instance_variable_set(:@zip_path, File.join(@temp_dir, "test.zip"))

    # First write the document.xml file
    doc.instance_variable_set(:@docx_word_path, File.join(@mock_docx_root, "word"))
    FileUtils.mkdir_p(File.join(@mock_docx_root, "word"))
    File.write(File.join(@mock_docx_root, "word", "document.xml"), "<test>content</test>")

    doc.send(:zip_docx_folder)

    zip_path = doc.instance_variable_get(:@zip_path)
    assert File.exist?(zip_path)

    # Debug: List all entries in the zip
    entries = []
    Zip::File.open(zip_path) do |zip|
      entries = zip.entries.map(&:name)
    end

    # Verify zip contains expected files
    Zip::File.open(zip_path) do |zip|
      assert zip.find_entry("[Content_Types].xml"), "Should contain [Content_Types].xml. Found: #{entries}"
      assert zip.find_entry("_rels/.rels"), "Should contain _rels/.rels. Found: #{entries}"
      assert zip.find_entry("word/document.xml"), "Should contain word/document.xml. Found: #{entries}"
    end
  end

  def test_copy_to_public_creates_file_in_destination
    doc = OxmlMaker::Document.new(filename: "test.docx")
    doc.send(:prepare_paths)

    # Create source file
    source_path = File.join(@temp_dir, "source.docx")
    File.write(source_path, "test content")
    doc.instance_variable_set(:@docx_path, source_path)

    # Create public directory
    public_dir = File.join(@temp_dir, "public")
    FileUtils.mkdir_p(public_dir)

    doc.send(:copy_to_public, public_dir)

    destination_path = File.join(public_dir, "test.docx")
    assert File.exist?(destination_path)
    assert_equal "test content", File.read(destination_path)
  end

  def test_detect_public_dir_without_rails
    doc = OxmlMaker::Document.new

    # Ensure Rails is not defined
    original_rails = Object.const_defined?(:Rails) ? Rails : nil
    Object.send(:remove_const, :Rails) if Object.const_defined?(:Rails)

    public_dir = doc.send(:detect_public_dir)
    assert_equal "public", public_dir

    # Restore Rails if it existed
    Object.const_set(:Rails, original_rails) if original_rails
  end

  def test_detect_public_dir_with_mock_rails
    doc = OxmlMaker::Document.new

    # Mock Rails
    rails_mock = Object.new
    def rails_mock.root
      "/fake/rails/root"
    end

    # Temporarily define Rails
    original_rails = Object.const_defined?(:Rails) ? Rails : nil
    Object.send(:remove_const, :Rails) if Object.const_defined?(:Rails)
    Object.const_set(:Rails, rails_mock)

    public_dir = doc.send(:detect_public_dir)
    assert_equal "/fake/rails/root/public", public_dir

    # Restore original state
    Object.send(:remove_const, :Rails)
    Object.const_set(:Rails, original_rails) if original_rails
  end

  def test_full_create_workflow
    doc = OxmlMaker::Document.new(filename: "test_workflow.docx", params: @sample_params)

    # Test individual components of the create workflow
    doc.send(:prepare_paths)

    # Mock paths to use temp directory
    doc.instance_variable_set(:@docx_root, @mock_docx_root)
    doc.instance_variable_set(:@docx_word_path, File.join(@mock_docx_root, "word"))
    doc.instance_variable_set(:@zip_path, File.join(@temp_dir, "document.zip"))
    doc.instance_variable_set(:@docx_path, File.join(@temp_dir, "document.docx"))

    # Test cleanup
    doc.send(:cleanup_temp_files)

    # Test directory creation
    doc.send(:ensure_directories)
    assert Dir.exist?(File.join(@mock_docx_root, "word"))

    # Test document writing
    doc.send(:write_document_xml)
    document_xml_path = File.join(@mock_docx_root, "word", "document.xml")
    assert File.exist?(document_xml_path)

    content = File.read(document_xml_path)
    assert_includes content, "Hello, World!"

    # Test zip creation
    doc.send(:zip_docx_folder)
    assert File.exist?(doc.instance_variable_get(:@zip_path))

    # Test copy step (create the source file first)
    FileUtils.cp(doc.instance_variable_get(:@zip_path), doc.instance_variable_get(:@docx_path))
    assert File.exist?(doc.instance_variable_get(:@docx_path))

    # Test copy to public with temp directory
    doc.send(:copy_to_public, @temp_dir)
    assert File.exist?(File.join(@temp_dir, "test_workflow.docx"))

    # Final cleanup
    doc.send(:cleanup_temp_files)
  end

  def test_ensure_public_directory_creates_directory
    doc = OxmlMaker::Document.new

    test_public_dir = File.join(@temp_dir, "public")
    refute Dir.exist?(test_public_dir)

    doc.send(:ensure_public_directory, test_public_dir)
    assert Dir.exist?(test_public_dir)
  end

  def test_detect_public_dir_in_non_rails_environment
    doc = OxmlMaker::Document.new

    # Ensure Rails is not defined for this test
    if Object.const_defined?(:Rails)
      rails_backup = Rails
      Object.send(:remove_const, :Rails)
    end

    public_dir = doc.send(:detect_public_dir)
    assert_equal "public", public_dir

    # Restore Rails if it was defined
    Object.const_set(:Rails, rails_backup) if defined?(rails_backup)
  end

  def test_copy_to_public_creates_public_directory_automatically
    doc = OxmlMaker::Document.new(filename: "auto_public_test.docx")

    # Create a source file
    source_file = File.join(@temp_dir, "source.docx")
    File.write(source_file, "test content")
    doc.instance_variable_set(:@docx_path, source_file)

    # Test directory that doesn't exist yet
    public_dir = File.join(@temp_dir, "auto_created_public")
    refute Dir.exist?(public_dir)

    # This should create the directory and copy the file
    doc.send(:copy_to_public, public_dir)

    assert Dir.exist?(public_dir)
    assert File.exist?(File.join(public_dir, "auto_public_test.docx"))
    assert_equal "test content", File.read(File.join(public_dir, "auto_public_test.docx"))
  end

  private

  def create_mock_docx_structure
    # Create basic DOCX structure for testing
    FileUtils.mkdir_p(File.join(@mock_docx_root, "_rels"))
    FileUtils.mkdir_p(File.join(@mock_docx_root, "word"))

    # Create [Content_Types].xml
    content_types = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>
    XML
    File.write(File.join(@mock_docx_root, "[Content_Types].xml"), content_types)

    # Create _rels/.rels
    rels = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    XML
    File.write(File.join(@mock_docx_root, "_rels", ".rels"), rels)
  end
end
