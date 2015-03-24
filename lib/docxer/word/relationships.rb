require 'docxer/word/relationships/relationship'

module Docxer
  module Word
    class Relationships
      @@prefix = 'rId'

      attr_accessor :relationships

      def initialize
        @relationships = []
        @counter = 0
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", "styles.xml")
        add("http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects", "stylesWithEffects.xml")
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings", "settings.xml")
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings", "webSettings.xml")
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable", "fontTable.xml")
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme", "theme/theme1.xml")
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes", "endnotes.xml")
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes", "footnotes.xml")
        add("http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering", "numbering.xml")
      end

      def add(type, target)
        relationship = create_relationship(sequence, type, target)
        @relationships << relationship if relationship
        relationship
      end

      def render(zip)
        zip.put_next_entry('word/_rels/document.xml.rels')
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
end
