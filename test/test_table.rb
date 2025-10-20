# frozen_string_literal: true

require "test_helper"

class TestTable < Minitest::Test
  def setup
    @sample_table_data = {
      columns: [
        { name: "Name", width: 2000 },
        { name: "Age", width: 1500 }
      ],
      rows: [
        {
          cells: [
            { value: :name, width: 2000 },
            { value: :age, width: 1500 }
          ]
        }
      ],
      data: {
        0 => [
          OpenStruct.new(name: "John", age: 30),
          OpenStruct.new(name: "Jane", age: 25)
        ]
      },
      font_size: 22
    }
  end

  def test_initialize
    table = OxmlMaker::Table.new(@sample_table_data)

    assert_equal @sample_table_data, table.table
  end

  def test_template_includes_table_structure
    table = OxmlMaker::Table.new(@sample_table_data)
    template = table.template

    assert_includes template, "<w:tbl>"
    assert_includes template, "</w:tbl>"
    assert_includes template, "<w:tblPr>"
    assert_includes template, "<w:tblGrid>"
    assert_includes template, "<w:tr>"
  end

  def test_template_includes_headers
    table = OxmlMaker::Table.new(@sample_table_data)
    template = table.template

    assert_includes template, "<w:t>Name</w:t>"
    assert_includes template, "<w:t>Age</w:t>"
  end

  def test_template_includes_data_rows
    table = OxmlMaker::Table.new(@sample_table_data)
    template = table.template

    assert_includes template, "<w:t>John</w:t>"
    assert_includes template, "<w:t>30</w:t>"
    assert_includes template, "<w:t>Jane</w:t>"
    assert_includes template, "<w:t>25</w:t>"
  end

  def test_template_with_custom_font_size
    custom_data = @sample_table_data.dup
    custom_data[:font_size] = 16

    table = OxmlMaker::Table.new(custom_data)
    template = table.template

    assert_includes template, '<w:sz w:val="16"/>'
  end

  def test_template_with_default_font_size
    data_without_font_size = @sample_table_data.dup
    data_without_font_size.delete(:font_size)

    table = OxmlMaker::Table.new(data_without_font_size)
    template = table.template

    assert_includes template, '<w:sz w:val="22"/>'
  end

  def test_grid_col_with_custom_width
    table = OxmlMaker::Table.new(@sample_table_data)

    # Use send to call private method for testing
    grid_col = table.send(:grid_col, { width: 3000 })

    assert_equal '<w:gridCol w:w="3000"/>', grid_col
  end

  def test_grid_col_with_default_width
    table = OxmlMaker::Table.new(@sample_table_data)

    grid_col = table.send(:grid_col, {})

    assert_equal '<w:gridCol w:w="2000"/>', grid_col
  end

  def test_cell_value_escapes_xml_characters
    table = OxmlMaker::Table.new(@sample_table_data)

    # Test XML character escaping
    assert_equal "John &amp; Jane", table.send(:cell_value, "John & Jane")
    assert_equal "&lt;tag&gt;", table.send(:cell_value, "<tag>")
    assert_equal "&quot;quoted&quot;", table.send(:cell_value, '"quoted"')
    assert_equal "&apos;single&apos;", table.send(:cell_value, "'single'")
  end

  def test_v_merge_tag_restart
    table = OxmlMaker::Table.new(@sample_table_data)

    merge_tag = table.send(:v_merge_tag, { v_merge: true }, 0)

    assert_equal '<w:vMerge w:val="restart"/>', merge_tag
  end

  def test_v_merge_tag_continue
    table = OxmlMaker::Table.new(@sample_table_data)

    merge_tag = table.send(:v_merge_tag, { v_merge: true }, 1)

    assert_equal '<w:vMerge w:val="continue"/>', merge_tag
  end

  def test_v_merge_tag_none
    table = OxmlMaker::Table.new(@sample_table_data)

    merge_tag = table.send(:v_merge_tag, {}, 0)

    assert_equal "", merge_tag
  end

  def test_table_with_empty_data
    empty_data = {
      columns: [],
      rows: [],
      data: {},
      font_size: 22
    }

    table = OxmlMaker::Table.new(empty_data)
    template = table.template

    assert_includes template, "<w:tbl>"
    assert_includes template, "</w:tbl>"
  end

  def test_multi_line_cell_content_with_blank_value
    table = OxmlMaker::Table.new(@sample_table_data)

    content = table.send(:multi_line_cell_content, "")

    assert_includes content, "<w:p>"
    assert_includes content, "<w:t></w:t>"
  end

  def test_multi_line_cell_content_with_comma_separated_values
    table = OxmlMaker::Table.new(@sample_table_data)

    content = table.send(:multi_line_cell_content, "Line 1, Line 2, Line 3")

    assert_includes content, "<w:t>Line 1</w:t>"
    assert_includes content, "<w:t>Line 2</w:t>"
    assert_includes content, "<w:t>Line 3</w:t>"
  end
end
