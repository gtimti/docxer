require 'docxer/relationships/relationship'

module Docxer
  class Relationships
    @@prefix = 'rId'

    attr_accessor :relationships

    def initialize
      @relationships = []
      @counter = 0
      add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "word/document.xml")
      add("http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", "docProps/core.xml")
      add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties", "docProps/app.xml")
    end

    def add(type, target)
      relationship = create_relationship(sequence, type, target)
      @relationships << relationship if relationship
      relationship
    end

    def render(zip)
      zip.put_next_entry('_rels/.rels')
      zip.write(Docxer.to_xml(document))
    end

    private
    def document
      Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
        xml.Relationships(xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships') do
          @relationships.each do |relationship|
            relationship.build(xml)
          end
        end
      end
    end

    def create_relationship(id, type, target)
      Relationship.new(id, type, target)
    end

    def sequence
      @counter += 1
      @@prefix + @counter.to_s
    end

  end
end
