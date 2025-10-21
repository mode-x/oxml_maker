# OxmlMaker

A Ruby gem for generating Microsoft Word DOCX files using OpenXML. Create professional documents with tables, paragraphs, and custom formatting programmatically.

## Features

- ✅ **Generate Valid DOCX Files**: Creates Microsoft Word-compatible documents
- ✅ **Tables with Dynamic Data**: Populate tables from Ruby objects 
- ✅ **Paragraphs and Text**: Simple text content with proper XML structure
- ✅ **Page Configuration**: Control page size, margins, headers/footers
- ✅ **Rails Integration**: Automatically detects Rails environment for file placement
- ✅ **ZIP-based Structure**: Properly formatted DOCX files using rubyzip
- ✅ **XML Safety**: Handles special characters and escaping
- ✅ **Public Directory Management**: Auto-creates output directories

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'oxml_maker'
```

And then execute:

```bash
bundle install
```

Or install it yourself as:

```bash
gem install oxml_maker
```

## Quick Start

```ruby
require 'oxml_maker'

# Define document parameters
params = {
  sections: [
    { paragraph: { text: "Welcome to OxmlMaker" } },
    { 
      table: {
        columns: [
          { name: "Product", width: 3000 },
          { name: "Price", width: 2000 }
        ],
        rows: [
          {
            cells: [
              { value: :name, width: 3000 },
              { value: :price, width: 2000 }
            ]
          }
        ],
        data: {
          0 => [
            OpenStruct.new(name: "Widget", price: "$10.99"),
            OpenStruct.new(name: "Gadget", price: "$25.50")
          ]
        },
        font_size: 24
      }
    },
    { paragraph: { text: "End of document" } }
  ],
  page_size: { width: 12240, height: 15840 },
  page_margin: {
    top: 1440, right: 1440, bottom: 1440, left: 1440,
    header: 720, footer: 720, gutter: 0
  }
}

# Create and generate the document
doc = OxmlMaker::Document.new(filename: "example.docx", params: params)
doc.create

# The file will be created in:
# - Rails apps: Rails.root/public/example.docx
# - Non-Rails: ./public/example.docx (auto-created)
```

## Usage Examples

### Creating Tables

Tables can be populated with dynamic data from Ruby objects:

```ruby
# Define table structure
table_config = {
  columns: [
    { name: "Name", width: 2000 },
    { name: "Age", width: 1500 },
    { name: "Email", width: 3000 }
  ],
  rows: [
    {
      cells: [
        { value: :name, width: 2000 },
        { value: :age, width: 1500 },
        { value: :email, width: 3000 }
      ]
    }
  ],
  data: {
    0 => [
      OpenStruct.new(name: "John Doe", age: 30, email: "john@example.com"),
      OpenStruct.new(name: "Jane Smith", age: 25, email: "jane@example.com")
    ]
  },
  font_size: 12
}

params = {
  sections: [
    { table: table_config }
  ],
  page_size: { width: 12240, height: 15840 },
  page_margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 }
}
```

### Adding Paragraphs

Simple text content with proper XML formatting:

```ruby
params = {
  sections: [
    { paragraph: { text: "Document Title" } },
    { paragraph: { text: "This is the first paragraph." } },
    { paragraph: { text: "This is the second paragraph with more content." } }
  ],
  page_size: { width: 12240, height: 15840 },
  page_margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 }
}
```

### Output Location

The gem intelligently handles output location:

- **In Rails apps**: Files are saved to `Rails.root/public/`
- **Outside Rails**: Files are saved to `./public/` (created automatically)  
- **Custom location**: Pass a custom directory to `copy_to_public(custom_dir)`

## Development

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Testing

### Test Framework

The gem uses **Minitest** as its testing framework, which is included in Ruby's standard library. The test structure follows Ruby conventions:

- `test/test_helper.rb` - Common setup and helper methods
- `test/test_*.rb` - Individual test files for each class
- Tests run with `bundle exec rake test`

### Current Test Coverage

- **63 tests, 274 assertions, 0 failures, 0 errors**
- Unit tests for all classes (Document, Paragraph, Table)
- Integration tests for complete workflows
- ZIP functionality tests with rubyzip
- Rails environment detection tests
- Error handling and edge case tests

### Test Files

1. **test_oxml_maker.rb** - Tests for the main module
2. **test_paragraph.rb** - Tests for the Paragraph class  
3. **test_table.rb** - Tests for the Table class
4. **test_document.rb** - Tests for the Document class
5. **test_integration.rb** - Integration tests combining multiple classes
6. **test_zip_functionality.rb** - ZIP creation and DOCX structure tests
7. **test_full_workflow_integration.rb** - Complete end-to-end workflow tests

### Running Tests

```bash
# Run all tests
bundle exec rake test

# Run a specific test file
bundle exec ruby -Itest test/test_paragraph.rb

# Run a specific test method
bundle exec ruby -Itest test/test_paragraph.rb -n test_initialize_with_valid_hash
```

### Types of Tests

#### 1. Unit Tests
Test individual methods and classes in isolation:

```ruby
def test_initialize_with_valid_hash
  data = { text: "Hello, World!" }
  paragraph = OxmlMaker::Paragraph.new(data)
  
  assert_equal data, paragraph.data
end
```

#### 2. Template/Output Tests
Verify that XML output contains expected elements:

```ruby
def test_template_includes_xml_declaration
  doc = OxmlMaker::Document.new(params: @sample_params)
  template = doc.template

  assert_includes template, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
end
```

#### 3. Integration Tests
Test multiple classes working together:

```ruby
def test_document_with_mixed_content
  params = {
    sections: [
      { paragraph: { text: "Document Title" } },
      { table: table_data }
    ],
    page_size: { width: 12240, height: 15840 }
  }

  document = OxmlMaker::Document.new(params: params)
  xml = document.template

  assert_includes xml, "<w:t>Document Title</w:t>"
end
```

#### 4. ZIP Functionality Tests
Verify DOCX file creation and structure:

```ruby
def test_created_zip_is_valid_docx
  doc = OxmlMaker::Document.new(filename: "test.docx", params: @params)
  doc.create

  # Verify ZIP structure matches DOCX requirements
  Zip::File.open(@docx_path) do |zip|
    assert zip.find_entry("[Content_Types].xml"), "Should contain Content Types"
    assert zip.find_entry("_rels/.rels"), "Should contain relationships"
    assert zip.find_entry("word/document.xml"), "Should contain main document"
  end
end
```

### Testing Best Practices

1. **Use Setup and Teardown** for consistent test environments
2. **Test Edge Cases** including empty inputs, nil values, and special characters
3. **Use Descriptive Test Names** that explain what is being tested
4. **Test Both Success and Failure Paths** to ensure robust error handling
5. **Mock External Dependencies** when needed (Rails environment, file systems)

### Common Assertions

- `assert_equal expected, actual` - Test equality
- `assert_includes collection, item` - Test inclusion
- `assert_raises(ExceptionClass) { code }` - Test exceptions
- `assert condition` - Test truthiness
- `refute condition` - Test falsiness
- `assert_kind_of Class, object` - Test object type

## Why OxmlMaker is Different: The Ruby Object Revolution

### The Problem with Traditional DOCX Gems

Most Ruby DOCX libraries force you into rigid patterns:

```ruby
# Traditional approach - manual, inflexible
builder.table do |t|
  t.row ["Name", "Age"]        # Static arrays only
  t.row ["John", "30"]         # Manual string building
  t.row ["Jane", "25"]         # Can't use objects directly
end
```

### OxmlMaker's Revolutionary Approach

**Direct Ruby Object Mapping** - Use ANY Ruby object with methods:

```ruby
# Revolutionary - works with ANY objects!
data: {
  0 => [
    OpenStruct.new(name: "John", age: 30),     # OpenStruct
    User.find(1),                             # ActiveRecord model  
    JSON.parse('{"name": "Jane"}'),           # JSON object
    api_response.data,                        # API response
    custom_object                             # Any object with methods
  ]
}

# Table automatically calls methods dynamically
{ value: :name }  # Calls object.name on each object
{ value: :email } # Calls object.email on each object
```

### Perfect for Modern Architectures

#### 1. **JSON-Native Design**
Built for API-first and headless architectures:

```ruby
# HTML → JSON → DOCX pipeline
html_content = "<div><h1>Title</h1><table>...</table></div>"
json_data = html_to_json_parser(html_content)

# Direct consumption - no transformation needed!
doc = OxmlMaker::Document.new(
  filename: "converted.docx", 
  params: json_data  # Pure JSON input
)
```

#### 2. **Polymorphic Data Handling**
Mix any data sources in the same document:

```ruby
# Different object types in the same table!
data: {
  0 => [
    user_model,                    # ActiveRecord object
    { name: "Jane" },              # Hash
    OpenStruct.new(name: "Bob"),   # OpenStruct
    json_api_response              # JSON object
  ]
}
```

#### 3. **Configuration-Driven Architecture**
Documents are pure data structures - serializable and cacheable:

```ruby
# Store templates as JSON in database
template = DocumentTemplate.find_by(name: "invoice")
live_data = Invoice.includes(:line_items).find(params[:id])

# Merge template with live data
params = template.structure.deep_merge({
  sections: [{
    table: {
      data: { 0 => live_data.line_items.to_a }  # Direct AR relation!
    }
  }]
})
```

### Comparison with Other Ruby DOCX Gems

| Feature | OxmlMaker | Caracal | ruby-docx | docx | sablon |
|---------|-----------|---------|-----------|------|--------|
| **Ruby Object Mapping** | ✅ Dynamic | ❌ Manual | ❌ Manual | ❌ Template | ❌ Template |
| **JSON Serializable** | ✅ 100% | ❌ Code-based | ❌ Code-based | ❌ Template | ❌ Template |
| **Any Ruby Object** | ✅ Polymorphic | ❌ Arrays only | ❌ Limited | ❌ Static | ❌ Mail merge |
| **HTML→JSON→DOCX** | ✅ Native | ❌ Complex | ❌ N/A | ❌ N/A | ❌ Template only |
| **API-Friendly** | ✅ Pure data | ❌ Code required | ❌ Code required | ❌ Files | ❌ Templates |
| **Microservices Ready** | ✅ Stateless | ❌ Complex | ❌ Complex | ❌ File-based | ❌ Template-based |

### Perfect Use Cases

#### **CMS/Blog Export**
```ruby
# Blog post with mixed content types
post_json = {
  sections: [
    { paragraph: { text: post.title } },
    { table: { 
        data: { 0 => post.comments.approved }  # Direct relation!
      }
    }
  ]
}
```

#### **API Report Generation**
```ruby
# Consume external APIs directly
api_response = HTTParty.get("https://api.example.com/reports/#{id}")
json_data = JSON.parse(api_response.body, object_class: OpenStruct)

# No transformation needed!
doc = OxmlMaker::Document.new(params: { sections: json_data.sections })
```

#### **Dynamic Form Processing**
```ruby
# Form submission → DOCX
form_data = params[:form_responses]  # Frontend JSON

document_params = {
  sections: [{
    table: {
      data: { 0 => form_data.map { |item| OpenStruct.new(item) } }
    }
  }]
}
```

### The Architectural Advantage

**Zero Impedance Mismatch** - Your data flows directly into documents without transformation layers, making OxmlMaker perfect for:

- **Headless CMS** systems
- **API-first** applications  
- **Microservice** architectures
- **HTML→DOCX** conversion pipelines
- **Real-time report** generation
- **JSON-driven** document templates

This isn't just a different API - it's a **fundamentally superior architecture** for modern Ruby applications.

## Dependencies

- **rubyzip (~> 3.2)** - For ZIP file creation and DOCX generation
- **Standard Ruby libraries** (FileUtils, etc.)
- **Optional: Rails** - For automatic public directory detection

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/mode-x/oxml_maker. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/mode-x/oxml_maker/blob/master/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the OxmlMaker project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/mode-x/oxml_maker/blob/master/CODE_OF_CONDUCT.md).
