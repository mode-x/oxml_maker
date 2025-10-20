# frozen_string_literal: true

require "test_helper"
require "fileutils"

class TestFullWorkflowIntegration < Minitest::Test
  def setup
    @temp_dir = Dir.mktmpdir
    @test_public_dir = File.join(@temp_dir, "public")
    @complex_params = {
      sections: [
        { paragraph: { text: "Document Title - Integration Test" } },
        {
          table: {
            columns: [
              { name: "Product", width: 3000 },
              { name: "Price", width: 2000 },
              { name: "Quantity", width: 1500 }
            ],
            rows: [
              {
                cells: [
                  { value: :name, width: 3000 },
                  { value: :price, width: 2000 },
                  { value: :quantity, width: 1500 }
                ]
              }
            ],
            data: {
              0 => [
                create_test_object(name: "Widget A", price: "$10.99", quantity: "5"),
                create_test_object(name: "Gadget B", price: "$25.50", quantity: "3"),
                create_test_object(name: "Tool C", price: "$45.00", quantity: "1")
              ]
            },
            font_size: 24
          }
        },
        { paragraph: { text: "End of Report" } }
      ],
      page_size: { width: 12_240, height: 15_840 },
      page_margin: {
        top: 1440, right: 1440, bottom: 1440, left: 1440,
        header: 720, footer: 720, gutter: 0
      }
    }
  end

  def teardown
    FileUtils.rm_rf(@temp_dir) if @temp_dir && Dir.exist?(@temp_dir)
  end

  def test_complete_docx_creation_workflow
    # This is the full integration test - from params to final DOCX file
    # Create a test-specific document class to override just the public directory
    test_doc = Class.new(OxmlMaker::Document) do
      def initialize(filename:, params:, test_public_dir:)
        super(filename: filename, params: params)
        @test_public_dir = test_public_dir
      end

      def detect_public_dir
        @test_public_dir
      end
    end

    doc = test_doc.new(
      filename: "integration_test.docx",
      params: @complex_params,
      test_public_dir: @test_public_dir
    )

    # Execute the complete workflow
    begin
      doc.create
      success = true
      error = nil
    rescue StandardError => e
      success = false
      error = e
    end

    # Verify the workflow completed successfully
    assert success, "Complete DOCX creation should succeed: #{error}"

    # Verify the final DOCX file was created in public directory
    final_docx_path = File.join(@test_public_dir, "integration_test.docx")
    assert File.exist?(final_docx_path), "Final DOCX file should exist in public directory"

    # Verify the file is a valid ZIP/DOCX
    require "zip"
    Zip::File.open(final_docx_path) do |zip|
      # Check DOCX structure
      assert zip.find_entry("[Content_Types].xml"), "Should contain Content Types"
      assert zip.find_entry("_rels/.rels"), "Should contain relationships"
      assert zip.find_entry("word/document.xml"), "Should contain main document"

      # Verify content
      document_xml = zip.read("word/document.xml")
      assert_includes document_xml, "Document Title - Integration Test"
      assert_includes document_xml, "Widget A"
      assert_includes document_xml, "$10.99"
      assert_includes document_xml, "End of Report"
      assert_includes document_xml, "<w:tbl>", "Should contain table XML"
      assert_includes document_xml, "<w:p>", "Should contain paragraph XML"
    end

    # Verify file size is reasonable (should be > 1KB for this content)
    file_size = File.size(final_docx_path)
    assert file_size > 1024, "DOCX file should be substantial size, got #{file_size} bytes"
  end

  def test_workflow_with_rails_environment_simulation
    # Simulate a Rails environment
    rails_mock = Class.new do
      def self.root
        @root ||= "/fake/rails/app"
      end

      def self.respond_to?(method)
        method == :root || super
      end
    end

    # Temporarily replace Rails constant
    original_rails = Object.const_defined?(:Rails) ? Rails : nil
    Object.send(:remove_const, :Rails) if Object.const_defined?(:Rails)
    Object.const_set(:Rails, rails_mock)

    doc = OxmlMaker::Document.new(filename: "rails_test.docx", params: @complex_params)

    # Test that Rails.root is detected correctly
    detected_dir = doc.send(:detect_public_dir)
    assert_equal "/fake/rails/app/public", detected_dir

    # Restore original Rails constant
    Object.send(:remove_const, :Rails)
    Object.const_set(:Rails, original_rails) if original_rails
  end

  def test_workflow_handles_complex_content
    # Test with complex content including special characters
    complex_params = {
      sections: [
        { paragraph: { text: "Test with special chars: & < > \" '" } },
        {
          table: {
            columns: [{ name: "Data & Info", width: 3000 }],
            rows: [{ cells: [{ value: :content, width: 3000 }] }],
            data: {
              0 => [
                create_test_object(content: "Content with <tags> & symbols")
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

    # Create test document class
    test_doc = Class.new(OxmlMaker::Document) do
      def initialize(filename:, params:, test_public_dir:)
        super(filename: filename, params: params)
        @test_public_dir = test_public_dir
      end

      def detect_public_dir
        @test_public_dir
      end
    end

    doc = test_doc.new(
      filename: "complex_content.docx",
      params: complex_params,
      test_public_dir: @test_public_dir
    )

    # This should handle special characters correctly
    begin
      doc.create
      success = true
    rescue StandardError => e
      success = false
      error = e
    end

    assert success, "Should handle complex content: #{error}"

    # Verify the content in the final file
    final_file = File.join(@test_public_dir, "complex_content.docx")
    assert File.exist?(final_file)

    Zip::File.open(final_file) do |zip|
      document_xml = zip.read("word/document.xml")

      # Paragraph content should be preserved (note: not escaped in current implementation)
      assert_includes document_xml, "Test with special chars: & < > \" '"

      # Table content should be escaped
      assert_includes document_xml, "Content with &lt;tags&gt; &amp; symbols"
    end
  end

  def test_workflow_error_recovery
    # Test what happens when there are issues in the workflow
    doc = OxmlMaker::Document.new(filename: "error_test.docx", params: @complex_params)

    # Simulate a permission error by using a read-only directory
    readonly_dir = File.join(@temp_dir, "readonly")
    FileUtils.mkdir_p(readonly_dir)
    FileUtils.chmod(0o444, readonly_dir) # Read-only

    doc.define_singleton_method(:detect_public_dir) { readonly_dir }

    # This should fail gracefully
    begin
      doc.create
      success = true
    rescue StandardError => e
      success = false
      error = e
    end

    # Clean up the readonly directory
    FileUtils.chmod(0o755, readonly_dir)

    # The test should show that errors are handled appropriately
    # (In a production system, you might want to log errors or handle them differently)
    refute success, "Should fail when public directory is not writable"
    assert_kind_of StandardError, error
  end

  def test_multiple_documents_in_sequence
    # Test creating multiple documents to ensure no state pollution
    filenames = ["doc1.docx", "doc2.docx", "doc3.docx"]

    # Create test document class
    test_doc = Class.new(OxmlMaker::Document) do
      def initialize(filename:, params:, test_public_dir:)
        super(filename: filename, params: params)
        @test_public_dir = test_public_dir
      end

      def detect_public_dir
        @test_public_dir
      end
    end

    filenames.each_with_index do |filename, index|
      params = {
        sections: [
          { paragraph: { text: "Document #{index + 1} content" } }
        ],
        page_size: { width: 12_240, height: 15_840 },
        page_margin: {
          top: 1440, right: 1440, bottom: 1440, left: 1440,
          header: 720, footer: 720, gutter: 0
        }
      }

      doc = test_doc.new(
        filename: filename,
        params: params,
        test_public_dir: @test_public_dir
      )

      doc.create

      final_file = File.join(@test_public_dir, filename)
      assert File.exist?(final_file), "Document #{filename} should be created"

      # Verify unique content
      Zip::File.open(final_file) do |zip|
        document_xml = zip.read("word/document.xml")
        assert_includes document_xml, "Document #{index + 1} content"
      end
    end

    # All three files should exist
    filenames.each do |filename|
      assert File.exist?(File.join(@test_public_dir, filename))
    end
  end
end
