# frozen_string_literal: true

require "test_helper"
require "zip"
require "fileutils"

class TestZipFunctionality < Minitest::Test
  def setup
    @temp_dir = Dir.mktmpdir
    @sample_params = {
      sections: [
        { paragraph: { text: "Test Document Content" } },
        {
          table: {
            columns: [
              { name: "Name", width: 2000 },
              { name: "Value", width: 1500 }
            ],
            rows: [
              {
                cells: [
                  { value: :name, width: 2000 },
                  { value: :value, width: 1500 }
                ]
              }
            ],
            data: {
              0 => [
                create_test_object(name: "Item 1", value: "100"),
                create_test_object(name: "Item 2", value: "200")
              ]
            },
            font_size: 22
          }
        }
      ],
      page_size: { width: 12_240, height: 15_840 },
      page_margin: {
        top: 1440, right: 1440, bottom: 1440, left: 1440,
        header: 720, footer: 720, gutter: 0
      }
    }

    create_complete_docx_structure
  end

  def teardown
    FileUtils.rm_rf(@temp_dir) if @temp_dir && Dir.exist?(@temp_dir)
  end

  def test_zip_creation_with_rubyzip
    # Test that we can create a zip file using the rubyzip gem
    test_files = {
      "test1.txt" => "Content of file 1",
      "test2.txt" => "Content of file 2",
      "folder/test3.txt" => "Content of file 3 in folder"
    }

    # Create test files
    test_files.each do |filename, content|
      filepath = File.join(@temp_dir, filename)
      FileUtils.mkdir_p(File.dirname(filepath))
      File.write(filepath, content)
    end

    zip_path = File.join(@temp_dir, "test.zip")

    # Create zip using same logic as Document class
    entries = Dir.glob(File.join(@temp_dir, "**", "*"), File::FNM_DOTMATCH).select { |f| File.file?(f) }
    entries.reject! { |f| f.end_with?(".zip") } # Don't include the zip file itself

    Zip::File.open(zip_path, create: true) do |zipfile|
      entries.each do |file|
        rel_path = file.sub("#{@temp_dir}/", "")
        zipfile.add(rel_path, file)
      end
    end

    assert File.exist?(zip_path)

    # Verify zip contents
    Zip::File.open(zip_path) do |zip|
      assert zip.find_entry("test1.txt")
      assert zip.find_entry("test2.txt")
      assert zip.find_entry("folder/test3.txt")

      # Verify content
      assert_equal "Content of file 1", zip.read("test1.txt")
      assert_equal "Content of file 2", zip.read("test2.txt")
      assert_equal "Content of file 3 in folder", zip.read("folder/test3.txt")
    end
  end

  def test_docx_zip_structure
    doc = OxmlMaker::Document.new(filename: "test.docx", params: @sample_params)

    # Use our complete docx structure
    docx_root = File.join(@temp_dir, "docx")
    doc.instance_variable_set(:@docx_root, docx_root)
    doc.instance_variable_set(:@docx_word_path, File.join(docx_root, "word"))
    doc.instance_variable_set(:@zip_path, File.join(@temp_dir, "result.zip"))

    # Write the document XML
    doc.send(:write_document_xml)

    # Create the zip
    doc.send(:zip_docx_folder)

    zip_path = doc.instance_variable_get(:@zip_path)
    assert File.exist?(zip_path)

    # Verify the zip has proper DOCX structure
    Zip::File.open(zip_path) do |zip|
      # Check required DOCX files
      assert zip.find_entry("[Content_Types].xml"), "Missing [Content_Types].xml"
      assert zip.find_entry("_rels/.rels"), "Missing _rels/.rels"
      assert zip.find_entry("word/document.xml"), "Missing word/document.xml"

      # Verify document content
      document_content = zip.read("word/document.xml")
      assert_includes document_content, "Test Document Content"
      assert_includes document_content, "<w:t>Item 1</w:t>"
      assert_includes document_content, "<w:t>Item 2</w:t>"
      assert_includes document_content, "<w:t>Name</w:t>"
      assert_includes document_content, "<w:t>Value</w:t>"
    end
  end

  def test_created_zip_is_valid_docx
    doc = OxmlMaker::Document.new(filename: "valid_test.docx", params: @sample_params)

    # Use our complete docx structure
    docx_root = File.join(@temp_dir, "docx")
    doc.instance_variable_set(:@docx_root, docx_root)
    doc.instance_variable_set(:@docx_word_path, File.join(docx_root, "word"))
    doc.instance_variable_set(:@zip_path, File.join(@temp_dir, "valid_test.zip"))
    doc.instance_variable_set(:@docx_path, File.join(@temp_dir, "valid_test.docx"))

    # Simulate the full creation process (without Rails dependency)
    doc.send(:prepare_paths)
    doc.send(:cleanup_temp_files)
    doc.send(:ensure_directories)
    doc.send(:write_document_xml)
    doc.send(:zip_docx_folder)

    # Copy zip to docx
    FileUtils.cp(doc.instance_variable_get(:@zip_path), doc.instance_variable_get(:@docx_path))

    docx_path = doc.instance_variable_get(:@docx_path)
    assert File.exist?(docx_path)

    # Verify it's a valid zip file that can be opened as DOCX
    Zip::File.open(docx_path) do |zip|
      # Verify DOCX structure
      content_types = zip.read("[Content_Types].xml")
      assert_includes content_types, "wordprocessingml"

      rels = zip.read("_rels/.rels")
      assert_includes rels, "officeDocument"

      document = zip.read("word/document.xml")
      assert_includes document, "Test Document Content"
      assert_includes document, "w:document"
    end
  end

  def test_large_document_zip_performance
    # Test with a larger document to ensure zip works efficiently
    large_params = {
      sections: Array.new(50) do |i|
        { paragraph: { text: "Paragraph #{i + 1}: " + ("Lorem ipsum " * 20) } }
      end + [
        {
          table: {
            columns: Array.new(5) { |i| { name: "Column #{i + 1}", width: 2000 } },
            rows: [
              {
                cells: Array.new(5) { |i| { value: :"field_#{i}", width: 2000 } }
              }
            ],
            data: {
              0 => Array.new(100) do |i|
                data = {}
                5.times { |j| data[:"field_#{j}"] = "Data #{i}-#{j}" }
                create_test_object(data)
              end
            },
            font_size: 22
          }
        }
      ],
      page_size: { width: 12_240, height: 15_840 },
      page_margin: {
        top: 1440, right: 1440, bottom: 1440, left: 1440,
        header: 720, footer: 720, gutter: 0
      }
    }

    doc = OxmlMaker::Document.new(filename: "large_test.docx", params: large_params)

    docx_root = File.join(@temp_dir, "docx")
    doc.instance_variable_set(:@docx_root, docx_root)
    doc.instance_variable_set(:@docx_word_path, File.join(docx_root, "word"))
    doc.instance_variable_set(:@zip_path, File.join(@temp_dir, "large_test.zip"))

    # Ensure directories exist
    FileUtils.mkdir_p(File.join(docx_root, "word"))
    FileUtils.mkdir_p(File.join(docx_root, "_rels"))

    # Copy our template files
    FileUtils.cp(File.join(@temp_dir, "docx_template", "[Content_Types].xml"),
                 File.join(docx_root, "[Content_Types].xml"))
    FileUtils.cp(File.join(@temp_dir, "docx_template", "_rels", ".rels"), File.join(docx_root, "_rels", ".rels"))

    start_time = Time.now

    doc.send(:write_document_xml)
    doc.send(:zip_docx_folder)

    end_time = Time.now
    processing_time = end_time - start_time

    # Should complete in reasonable time (less than 5 seconds for this test)
    assert processing_time < 5, "Zip creation took too long: #{processing_time} seconds"

    zip_path = doc.instance_variable_get(:@zip_path)
    assert File.exist?(zip_path)

    # Verify the large content is present
    Zip::File.open(zip_path) do |zip|
      document_content = zip.read("word/document.xml")
      assert_includes document_content, "Paragraph 1:"
      assert_includes document_content, "Paragraph 50:"
      assert_includes document_content, "Data 0-0"
      assert_includes document_content, "Data 99-4"
    end
  end

  def test_zip_error_handling
    doc = OxmlMaker::Document.new

    # Test with non-existent directory
    doc.instance_variable_set(:@docx_root, "/non/existent/path")
    doc.instance_variable_set(:@zip_path, File.join(@temp_dir, "error_test.zip"))

    # Should handle gracefully without raising an error
    begin
      doc.send(:zip_docx_folder)
      success = true
    rescue StandardError => e
      success = false
      error = e
    end

    assert success, "zip_docx_folder should handle missing directory gracefully: #{error}"

    # Zip file should be created but empty/minimal
    zip_path = doc.instance_variable_get(:@zip_path)
    return unless File.exist?(zip_path)

    Zip::File.open(zip_path) do |zip|
      # Should be empty or have no entries
      assert zip.entries.empty?, "Zip should be empty when source directory doesn't exist"
    end
  end

  private

  def create_complete_docx_structure
    docx_template_dir = File.join(@temp_dir, "docx_template")
    docx_dir = File.join(@temp_dir, "docx")

    # Create template directory
    FileUtils.mkdir_p(File.join(docx_template_dir, "_rels"))
    FileUtils.mkdir_p(File.join(docx_template_dir, "word"))

    # Create actual docx directory
    FileUtils.mkdir_p(File.join(docx_dir, "_rels"))
    FileUtils.mkdir_p(File.join(docx_dir, "word"))

    # Create [Content_Types].xml
    content_types = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>
    XML

    # Create _rels/.rels
    rels = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
      </Relationships>
    XML

    # Write template files
    File.write(File.join(docx_template_dir, "[Content_Types].xml"), content_types)
    File.write(File.join(docx_template_dir, "_rels", ".rels"), rels)

    # Copy to actual docx directory
    File.write(File.join(docx_dir, "[Content_Types].xml"), content_types)
    File.write(File.join(docx_dir, "_rels", ".rels"), rels)
  end
end
