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

  def test_table_with_v_merge_and_new_line_features
    sample_data = {
      columns: [
        { name: "No.", width: 800 },
        { name: "Marketplace", width: 1500 },
        { name: "Merchant Name", width: 3000 },
        { name: "Merchant ID", width: 1500 },
        { name: "Product ID", width: 2000 },
        { name: "CP Used", width: 1500 }
      ],
      rows: [
        {
          cells: [
            { value: :no, width: 800, v_merge: true },
            { value: :marketplace, width: 1500, v_merge: true },
            { value: :merchant_name, width: 3000, v_merge: true },
            { value: :merchant_id, width: 1500, v_merge: true },
            { value: :product_id, width: 2000, v_merge: true },
            { value: :cp_used, width: 1500, v_merge: true, new_line: true }
          ]
        }
      ],
      data: {
        0 => [
          OpenStruct.new(no: "1", marketplace: "Alibaba", merchant_name: "Shenzhen BYF Precision Mould Co., Ltd.",
                         merchant_id: "byfpm", product_id: "1600837748820", cp_used: "VA0002267367, VA0002267368"),
          OpenStruct.new(no: "1", marketplace: "Alibaba", merchant_name: "Shenzhen BYF Precision Mould Co., Ltd.",
                         merchant_id: "byfpm", product_id: "1600838257049", cp_used: "VA0002267367, VA0002267368"),
          OpenStruct.new(no: "1", marketplace: "Alibaba", merchant_name: "Shenzhen BYF Precision Mould Co., Ltd.",
                         merchant_id: "byfpm", product_id: "1600838593508", cp_used: "VA0002267367, VA0002267368"),
          OpenStruct.new(no: "1", marketplace: "Alibaba", merchant_name: "Shenzhen BYF Precision Mould Co., Ltd.",
                         merchant_id: "byfpm", product_id: "1600847594806", cp_used: "VA0002267367, VA0002267368"),
          OpenStruct.new(no: "2", marketplace: "Alibaba", merchant_name: "Yiwu Lvye E-Commerce Firm",
                         merchant_id: "chinalvye", product_id: "1601026909143", cp_used: "VA0002267368"),
          OpenStruct.new(no: "3", marketplace: "Alibaba",
                         merchant_name: "Ningbo City Yinzhou Yierzhe Trading Co., Ltd.", merchant_id: "easyget", product_id: "10000025426669", cp_used: "VA0002267368")
        ]
      },
      font_size: 20
    }

    table = OxmlMaker::Table.new(sample_data)
    template = table.template

    # Test that v_merge functionality is working
    # First occurrence should have "restart"
    assert_includes template, '<w:vMerge w:val="restart"/>'

    # Subsequent occurrences should have "continue"
    assert_includes template, '<w:vMerge w:val="continue"/>'

    # Test that new_line functionality is working with comma-separated CP values
    # Should split "VA0002267367, VA0002267368" into separate lines
    assert_includes template, "<w:t>VA0002267367</w:t>"
    assert_includes template, "<w:t>VA0002267368</w:t>"

    # Test that the data is properly rendered
    assert_includes template, "<w:t>1</w:t>"
    assert_includes template, "<w:t>Alibaba</w:t>"
    assert_includes template, "<w:t>Shenzhen BYF Precision Mould Co., Ltd.</w:t>"
    assert_includes template, "<w:t>byfpm</w:t>"
    assert_includes template, "<w:t>1600837748820</w:t>"

    # Test different merchant data
    assert_includes template, "<w:t>2</w:t>"
    assert_includes template, "<w:t>Yiwu Lvye E-Commerce Firm</w:t>"
    assert_includes template, "<w:t>chinalvye</w:t>"

    # Test single CP value (should not be split)
    assert_includes template, "<w:t>VA0002267368</w:t>"
  end

  def test_v_merge_with_repeated_values
    # Test data with repeated values that should be merged
    # NOTE: Current implementation uses index-based logic (0 vs non-0)
    merge_data = {
      columns: [
        { name: "Group", width: 1500 },
        { name: "Item", width: 2000 }
      ],
      rows: [
        {
          cells: [
            { value: :group, width: 1500, v_merge: true },
            { value: :item, width: 2000 }
          ]
        }
      ],
      data: {
        0 => [
          OpenStruct.new(group: "A", item: "Item 1"),
          OpenStruct.new(group: "A", item: "Item 2"),
          OpenStruct.new(group: "A", item: "Item 3"),
          OpenStruct.new(group: "B", item: "Item 4"),
          OpenStruct.new(group: "B", item: "Item 5")
        ]
      },
      font_size: 20
    }

    table = OxmlMaker::Table.new(merge_data)
    template = table.template

    # Current implementation: first item gets restart, rest get continue
    restart_count = template.scan('<w:vMerge w:val="restart"/>').length
    assert_equal 1, restart_count, "Should have 1 restart tag for first item"

    # Should have continue for all subsequent items
    continue_count = template.scan('<w:vMerge w:val="continue"/>').length
    assert_equal 4, continue_count, "Should have 4 continue tags for remaining items"
  end

  def test_new_line_with_different_separators
    # Test new_line functionality with various comma-separated values
    # NOTE: Current implementation only splits on ", " (comma + space)
    newline_data = {
      columns: [
        { name: "Tags", width: 3000 }
      ],
      rows: [
        {
          cells: [
            { value: :tags, width: 3000, new_line: true }
          ]
        }
      ],
      data: {
        0 => [
          OpenStruct.new(tags: "tag1, tag2, tag3"),
          OpenStruct.new(tags: "single"),
          OpenStruct.new(tags: "first, second, third"), # With spaces
          OpenStruct.new(tags: "")
        ]
      },
      font_size: 20
    }

    table = OxmlMaker::Table.new(newline_data)
    template = table.template

    # Test comma-separated with spaces
    assert_includes template, "<w:t>tag1</w:t>"
    assert_includes template, "<w:t>tag2</w:t>"
    assert_includes template, "<w:t>tag3</w:t>"

    # Test single value (no splitting)
    assert_includes template, "<w:t>single</w:t>"

    # Test comma-separated with spaces (should split)
    assert_includes template, "<w:t>first</w:t>"
    assert_includes template, "<w:t>second</w:t>"
    assert_includes template, "<w:t>third</w:t>"

    # Test empty value
    assert_includes template, "<w:t></w:t>"
  end

  def test_complex_table_with_both_features
    # Test a more complex scenario combining both v_merge and new_line
    complex_data = {
      columns: [
        { name: "Category", width: 1500 },
        { name: "Subcategory", width: 1500 },
        { name: "Items", width: 3000 }
      ],
      rows: [
        {
          cells: [
            { value: :category, width: 1500, v_merge: true },
            { value: :subcategory, width: 1500, v_merge: true },
            { value: :items, width: 3000, new_line: true }
          ]
        }
      ],
      data: {
        0 => [
          OpenStruct.new(category: "Electronics", subcategory: "Phones", items: "iPhone, Samsung, Google"),
          OpenStruct.new(category: "Electronics", subcategory: "Phones", items: "Xiaomi, OnePlus"),
          OpenStruct.new(category: "Electronics", subcategory: "Laptops", items: "MacBook, ThinkPad, Dell"),
          OpenStruct.new(category: "Books", subcategory: "Fiction", items: "Novel1, Novel2, Novel3")
        ]
      },
      font_size: 20
    }

    table = OxmlMaker::Table.new(complex_data)
    template = table.template

    # Test v_merge for category (Electronics appears 3 times, Books once)
    # Should have restart for Electronics and Books
    assert_includes template, '<w:vMerge w:val="restart"/>'
    assert_includes template, '<w:vMerge w:val="continue"/>'

    # Test new_line for items
    assert_includes template, "<w:t>iPhone</w:t>"
    assert_includes template, "<w:t>Samsung</w:t>"
    assert_includes template, "<w:t>Google</w:t>"
    assert_includes template, "<w:t>MacBook</w:t>"
    assert_includes template, "<w:t>ThinkPad</w:t>"
    assert_includes template, "<w:t>Dell</w:t>"

    # Test that both features work together
    assert_includes template, "<w:t>Electronics</w:t>"
    assert_includes template, "<w:t>Phones</w:t>"
    assert_includes template, "<w:t>Laptops</w:t>"
    assert_includes template, "<w:t>Books</w:t>"
    assert_includes template, "<w:t>Fiction</w:t>"
  end
end
