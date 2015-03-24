require 'nokogiri'
require 'zip/zip'

require 'docxer/relationships'
require 'docxer/content_types'
require 'docxer/properties'
require 'docxer/word'
require 'docxer/document'

module Docxer

  def self.to_xml(document)
    document.to_xml(:save_with => Nokogiri::XML::Node::SaveOptions::AS_XML).gsub("\n", "\r\n").strip
  end

  def self.test
    word = Docxer::Document.new

    word.document.break
    word.document.break
    word.document.break

    file = word.render
    File.open('test.docx', 'wb') do |f|
      f.write(file.string)
    end
  end

end
