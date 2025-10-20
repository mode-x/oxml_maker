# frozen_string_literal: true

# OxmlMaker module
module OxmlMaker
  # Table class represents a table in an OXML document.
  # Initialized with a hash containing :columns, :rows, :data, and :font_size.
  class Table
    attr_reader :table

    def initialize(table)
      @table = table
    end

    def template
      <<~XML
        <w:tbl>
          <w:tblPr>
            <w:tblStyle w:val="DefaultTable"/>
            <w:tblW w:w="10468" w:type="dxa"/>
            <w:tblInd w:w="0" w:type="dxa"/>
            <w:tblBorders>
              <w:top w:color="666666" w:val="single" w:sz="4" w:space="0"/>
              <w:left w:color="666666" w:val="single" w:sz="4" w:space="0"/>
              <w:bottom w:color="666666" w:val="single" w:sz="4" w:space="0"/>
              <w:right w:color="666666" w:val="single" w:sz="4" w:space="0"/>
              <w:insideH w:color="666666" w:val="single" w:sz="4" w:space="0"/>
              <w:insideV w:color="666666" w:val="single" w:sz="4" w:space="0"/>
            </w:tblBorders>
            <w:tblLayout w:type="fixed"/>
            <w:tblLook w:val="0600"/>
          </w:tblPr>
          <w:tblGrid>
            #{table[:columns].map { |col| grid_col(col) }.join("\n")}
          </w:tblGrid>
          <w:tr>
            <w:tblPrEx>
              <w:tblCellMar>
                <w:top w:w="0" w:type="dxa"/>
                <w:bottom w:w="0" w:type="dxa"/>
              </w:tblCellMar>
            </w:tblPrEx>
            <w:trPr>
              <w:tblAlign w:val="center"/>
            </w:trPr>
            #{table[:columns].map { |column| header_row(column) }.join("\n")}
          </w:tr>
          #{table[:data].map { |_k, data| body_row(data) }.flatten.join("\n")}
        </w:tbl>
      XML
    end

    private

    def grid_col(column)
      width = column[:width] || 2000
      "<w:gridCol w:w=\"#{width}\"/>"
    end

    def header_row(column)
      <<~XML
        <w:tc>
          <w:tcPr>
            <w:tcW w:w="#{column[:width] || 2000}"/>
          </w:tcPr>
          <w:p>
            <w:pPr>
              <w:jc w:val="center"/>
              <w:rPr>
                <w:b/>
              </w:rPr>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:b/>
                <w:sz w:val="#{table[:font_size] || 22}"/>
              </w:rPr>
              <w:t>#{column[:name]}</w:t>
            </w:r>
          </w:p>
        </w:tc>
      XML
    end

    def body_row(data)
      data.map.with_index do |item, index|
        <<~XML
          <w:tr>
            #{table[:rows][0][:cells].map { |cell| cell_xml(cell, item, index) }.join("\n")}
          </w:tr>
        XML
      end.join("\n")
    end

    def cell_xml(cell, item, index)
      width = cell[:width] || 2000
      value = begin
        item.send(cell[:value])
      rescue StandardError
        ""
      end
      if cell[:new_line]
        <<~XML
          <w:tc>
            <w:tcPr>
              <w:tcW w:w='#{width}' w:type='dxa'/>
              <w:vAlign w:val="center"/>
            </w:tcPr>
            #{multi_line_cell_content(value)}
          </w:tc>
        XML
      else
        <<~XML
          <w:tc>
            <w:tcPr>
              <w:tcW w:w='#{width}' w:type='dxa'/>
              #{v_merge_tag(cell, index)}
              <w:vAlign w:val="center"/>
            </w:tcPr>
            <w:p>
              <w:pPr>
                <w:jc w:val='center'/>
              </w:pPr>
              <w:r>
                <w:rPr><w:sz w:val='#{table[:font_size] || 22}'/></w:rPr>
                <w:t>#{cell_value(value)}</w:t>
              </w:r>
            </w:p>
          </w:tc>
        XML
      end
    end

    def multi_line_cell_content(value)
      if value.nil? || value.to_s.strip.empty?
        <<~XML
          <w:p>
            <w:pPr>
              <w:spacing w:after='0'/>
              <w:jc w:val='center'/>
            </w:pPr>
            <w:r>
              <w:rPr><w:sz w:val='22'/></w:rPr>
              <w:t></w:t>
            </w:r>
          </w:p>
        XML
      else
        value.split(", ").map do |line|
          <<~XML
            <w:p>
              <w:pPr>
                <w:spacing w:after='0'/>
                <w:jc w:val='center'/>
              </w:pPr>
              <w:r>
                <w:rPr><w:sz w:val='22'/></w:rPr>
                <w:t>#{cell_value(line)}</w:t>
              </w:r>
            </w:p>
          XML
        end.join("\n")
      end
    end

    def v_merge_tag(cell, index)
      return "" unless cell[:v_merge]

      index.zero? ? '<w:vMerge w:val="restart"/>' : '<w:vMerge w:val="continue"/>'
    end

    def cell_value(value)
      value.to_s
           .gsub("&", "&amp;")
           .gsub("<", "&lt;")
           .gsub(">", "&gt;")
           .gsub('"', "&quot;")
           .gsub("'", "&apos;")
    end
  end
end
