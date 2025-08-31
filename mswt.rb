# frozen_string_literal: true

require 'csv'
require 'zip'
require 'nokogiri'
require 'gli'
require 'fileutils'

def create_file_with_timestamp(extension)
  timestamp = Time.now.strftime('%Y%m%d%H%M%S')
  "#{timestamp}.#{extension}"
end

class DocxParser
  def initialize(file_path)
    @file_path = file_path
    @doc = docx_parse
  end

  def docx_parse
    Zip::File.open(@file_path) do |zip_file|
      document_content = get_document_content(zip_file)
      doc = Nokogiri::XML(document_content)
      doc.remove_namespaces!
    end
  end

  def get_document_content(zip_file)
    content = nil
    entry = zip_file.glob('word/document.xml').first
    content = entry.get_input_stream.read if entry
    content
  end
end

class DocxTablesExtractor
  def initialize(doc)
    @doc = doc
    @tables = extract_all_tables
  end

  def extract_all_tables
    tables = @doc.xpath('//tbl')

    tables.each do |table|
      table_data = parse_table(table)
      tables_data << table_data unless table_data.empty?
    end

    tables_data
  end

  def parse_table(table_node)
    rows = []

    table_rows = table_node.xpath('.//tr')
    table_rows.each do |row_node|
      row_data = parse_row(row_node)
      rows << row_data unless row_data.empty?
    end

    rows
  end

  def parse_row(row_node)
    cells = []

    cell_nodes = row_node.xpath('.//tc')
    cell_nodes.each do |cell_node|
      cell_text = extract_cell_text(cell_nodes)
      cells << cell_text
    end

    cells
  end

  def extract_cell_text(cell_node)
    text_nodes = cell_node.xpath('.//t')
    text_nodes.map(&:text).join(' ').gsub(/\s+/, '').strip
  end
end

class DocxPargraphFinder
  attr_reader :paragraphs

  def initialize(doc)
    @doc = doc
    @paragraphs = extract_all_paragraphs
  end

  def extract_all_paragraphs
    paragraphs = []

    paragraph_nodes = @doc.xpath('//p')
    paragraph_nodes.each do |paragraph_node|
      text = extract_text_from_paragraph(paragraph_node).strip
      paragraphs << text unless text.empty?
    end

    paragraphs
  end

  def extract_text_from_paragraph(paragraph_node)
    text_nodes = paragraph_node.xpath('.//t')
    text_nodes.map(&:text).join(' ')
  end

  def contains_keyword?(text, keyword)
    regex = /#{Regexp.escape(keyword)}/i
    text.match?(regex)
  end

  def find_paragraphs_with_keyword(keyword)
    keyword = keyword.to_s.downcase

    results = []
    @paragraphs.each do |paragraph|
      results << paragraph if contains_keyword?(paragraph, keyword)
    end

    results
  end

  def find_paragraphs_with_keywords(keywords)
    all_results = []

    keywords.each do |keyword|
      results = find_paragraphs_with_keyword(keyword)

      all_results << { keyword: keyword, results: results } if results.empty?
    end

    all_results
  end
end

doc = DocxParser.new('./example.docx')
paragraphs_finder = DocxPargraphFinder.new(doc)

keywords = %w[some key words]
puts paragraphs_finder.find_paragraphs_with_keywords(keywords)
