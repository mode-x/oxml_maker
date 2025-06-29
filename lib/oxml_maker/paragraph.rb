# frozen_string_literal: true

# OxmlMaker module
module OxmlMaker
  # Represents a paragraph in an OXML document.
  # It can be used to create structured text content within a Word document.
  # The class is initialized with data that will be used to populate the paragraph.
  # @param data [Hash] The data to be used in the paragraph.
  # @example
  #   paragraph = Oxml::Paragraph.new(data: {text: "Hello, World!"})
  #   puts paragraph.template
  #   # Output: <w:p><w:r><w:t>Hello, World!</w:t></w:r></w:p>
  class Paragraph
    attr_reader :data

    def initialize(data = {})
      raise ArgumentError, "Data must be a Hash" unless data.is_a?(Hash)

      @data = data
    end

    def template
      <<~XML
        <w:p>
          <w:r>
            <w:t>#{data[:text]}</w:t>
          </w:r>
        </w:p>
      XML
    end
  end
end
