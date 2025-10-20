# frozen_string_literal: true

require "zip"
require "fileutils"

# OxmlMaker module
module OxmlMaker
  class Document
    attr_accessor :params
    attr_reader :filename

    def initialize(filename: nil, params: {}, &block)
      @filename = filename
      @params = params
      yield(self) if block_given?
    end

    def template
      <<~XML
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            #{generate_sections}
            <w:sectPr>
              <w:pgSz w:w='#{params[:page_size][:width]}' w:h='#{params[:page_size][:height]}'/>
              <w:pgMar w:top='#{params[:page_margin][:top]}' w:right='#{params[:page_margin][:right]}' w:bottom='#{params[:page_margin][:bottom]}' w:left='#{params[:page_margin][:left]}' w:header='#{params[:page_margin][:header]}' w:footer='#{params[:page_margin][:footer]}' w:gutter='#{params[:page_margin][:gutter]}'/>
            </w:sectPr>
          </w:body>
        </w:document>
      XML
    end

    def create
      prepare_paths
      cleanup_temp_files
      ensure_directories
      write_document_xml
      zip_docx_folder
      convert_zip_to_docx
      copy_to_public
      cleanup_temp_files
    end

    private

    def generate_sections
      params[:sections].map do |section|
        if section[:paragraph]
          OxmlMaker::Paragraph.new(section[:paragraph]).template
        elsif section[:table]
          OxmlMaker::Table.new(section[:table]).template
        end
      end.join("\n")
    end

    def prepare_paths
      @docx_root = File.join(__dir__, "docx")
      @docx_word_path = File.join(@docx_root, "word")
      @zip_path = File.join(__dir__, "document.zip")
      @docx_path = File.join(__dir__, "document.docx")
    end

    def cleanup_temp_files
      FileUtils.rm_f(@zip_path)
      FileUtils.rm_f(@docx_path)
    end

    def ensure_directories
      FileUtils.mkdir_p(@docx_word_path)
    end

    def write_document_xml
      File.write(File.join(@docx_word_path, "document.xml"), template)
    end

    def zip_docx_folder
      entries = Dir.glob(File.join(@docx_root, "**", "*"), File::FNM_DOTMATCH).select { |f| File.file?(f) }
      Zip::File.open(@zip_path, create: true) do |zipfile|
        entries.each do |file|
          rel_path = file.sub("#{@docx_root}/", "")
          zipfile.add(rel_path, file)
        end
      end
    end

    def convert_zip_to_docx
      FileUtils.cp(@zip_path, @docx_path)
    end

    def copy_to_public(destination_dir = nil)
      destination_dir ||= detect_public_dir
      ensure_public_directory(destination_dir)
      FileUtils.cp(@docx_path, File.join(destination_dir, filename))
    end

    def detect_public_dir
      if defined?(Rails) && Rails.respond_to?(:root) && Rails.root
        File.join(Rails.root, "public")
      else
        # For non-Rails environments (like testing), use local public directory
        "public"
      end
    end

    def ensure_public_directory(dir)
      FileUtils.mkdir_p(dir)
    end
  end
end
