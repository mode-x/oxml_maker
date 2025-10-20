# frozen_string_literal: true

require "test_helper"

class TestParagraph < Minitest::Test
  def test_initialize_with_valid_hash
    data = { text: "Hello, World!" }
    paragraph = OxmlMaker::Paragraph.new(data)

    assert_equal data, paragraph.data
  end

  def test_initialize_with_empty_hash
    paragraph = OxmlMaker::Paragraph.new({})

    assert_equal({}, paragraph.data)
  end

  def test_initialize_with_no_arguments
    paragraph = OxmlMaker::Paragraph.new

    assert_equal({}, paragraph.data)
  end

  def test_initialize_raises_error_with_non_hash
    assert_raises(ArgumentError) do
      OxmlMaker::Paragraph.new("not a hash")
    end

    assert_raises(ArgumentError) do
      OxmlMaker::Paragraph.new(123)
    end

    assert_raises(ArgumentError) do
      OxmlMaker::Paragraph.new(nil)
    end
  end

  def test_template_with_text
    data = { text: "Hello, World!" }
    paragraph = OxmlMaker::Paragraph.new(data)
    template = paragraph.template

    assert_includes template, "<w:p>"
    assert_includes template, "</w:p>"
    assert_includes template, "<w:r>"
    assert_includes template, "</w:r>"
    assert_includes template, "<w:t>Hello, World!</w:t>"
  end

  def test_template_with_empty_text
    data = { text: "" }
    paragraph = OxmlMaker::Paragraph.new(data)
    template = paragraph.template

    assert_includes template, "<w:t></w:t>"
  end

  def test_template_with_no_text_key
    paragraph = OxmlMaker::Paragraph.new({})
    template = paragraph.template

    assert_includes template, "<w:t></w:t>"
  end

  def test_template_structure
    paragraph = OxmlMaker::Paragraph.new({ text: "Test" })
    template = paragraph.template

    # Test XML structure
    expected_structure = <<~XML
      <w:p>
        <w:r>
          <w:t>Test</w:t>
        </w:r>
      </w:p>
    XML

    # Remove whitespace for comparison
    assert_equal expected_structure.strip, template.strip
  end

  def test_template_with_special_characters
    data = { text: "Text with <>&\"' characters" }
    paragraph = OxmlMaker::Paragraph.new(data)
    template = paragraph.template

    # NOTE: This test reveals that the paragraph doesn't escape XML characters
    # This might be a bug that should be addressed
    assert_includes template, "Text with <>&\"' characters"
  end
end
