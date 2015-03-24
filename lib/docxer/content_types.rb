require 'docxer/content_types/default'
require 'docxer/content_types/override'

module Docxer
  class ContentTypes

    attr_accessor :content_types

    def initialize
      @content_types = []
      add(:default, "rels", "application/vnd.openxmlformats-package.relationships+xml")
      add(:default, "xml", "application/xml")
      add(:default, "png", "image/png")
      add(:override, "/word/document.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
      add(:override, "/word/styles.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
      add(:override, "/word/stylesWithEffects.xml", "application/vnd.ms-word.stylesWithEffects+xml")
      add(:override, "/word/settings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml")
      add(:override, "/word/webSettings.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml")
      add(:override, "/word/fontTable.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml")
      add(:override, "/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml")
      add(:override, "/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")
      add(:override, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
      add(:override, "/word/numbering.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
      add(:override, "/word/footnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml")
      add(:override, "/word/endnotes.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml")
      add(:override, "/word/footer1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
      add(:override, "/word/header1.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml")
    end

    def add(type, attr, content)
      content_type = create_content_type(type, attr, content)
      @content_types << content_type if content_type
      content_type
    end

    def render(zip)
      zip.put_next_entry('[Content_Types].xml')
      zip.write(Docxer.to_xml(document))
    end

    private
    def document
      Nokogiri::XML::Builder.with(Nokogiri::XML('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')) do |xml|
        xml.Types(xmlns: 'http://schemas.openxmlformats.org/package/2006/content-types') do
          @content_types.each do |content_type|
            content_type.build(xml)
          end
        end
      end
    end

    # TODO: Need to rewrite create_content_type method
    def create_content_type(type, attr, content)
      send(:"create_#{type.to_s}", attr, content)
    rescue
      nil
    end

    def create_default(attr, content)
      Default.new(attr, content)
    end

    def create_override(attr, content)
      Override.new(attr, content)
    end

  end
end
